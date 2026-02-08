// formats/epub.rs - EPUB変換モジュール
//
// EPUB (Electronic Publication) ファイルを解析し、テキストを抽出して
// ドキュメントモデルに変換します。

use crate::converter::{
    ConvertError, Document, DocumentConverter, FontStyle, Metadata, Page, PageElement, TextAlign,
};

/// EPUBコンバーター
pub struct EpubConverter;

impl EpubConverter {
    pub fn new() -> Self {
        Self
    }
}

impl DocumentConverter for EpubConverter {
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError> {
        let cursor = std::io::Cursor::new(input);
        let mut archive = zip::ZipArchive::new(cursor)
            .map_err(|e| ConvertError::new("EPUB", &format!("ZIPアーカイブを開けません: {}", e)))?;

        // container.xml からルートファイルパスを取得
        let container_xml =
            read_zip_entry(&mut archive, "META-INF/container.xml")?;
        let rootfile_path = parse_rootfile_path(&container_xml)
            .ok_or_else(|| ConvertError::new("EPUB", "container.xmlからrootfileが見つかりません"))?;

        // OPF (content.opf) を読み込む
        let opf_xml = read_zip_entry(&mut archive, &rootfile_path)?;

        // メタデータ
        let metadata = parse_epub_metadata(&opf_xml);

        // コンテンツファイルパスをspine順で取得
        let base_dir = rootfile_path
            .rfind('/')
            .map(|i| &rootfile_path[..i + 1])
            .unwrap_or("");
        let content_paths = parse_spine_items(&opf_xml, base_dir);

        // 各コンテンツファイルからテキストを抽出
        let mut all_paragraphs: Vec<String> = Vec::new();
        for path in &content_paths {
            if let Ok(xhtml) = read_zip_entry(&mut archive, path) {
                let paragraphs = extract_xhtml_text(&xhtml);
                all_paragraphs.extend(paragraphs);
            }
        }

        if all_paragraphs.is_empty() {
            all_paragraphs.push("(コンテンツなし)".to_string());
        }

        // ドキュメントモデルに変換
        let style = FontStyle::default();
        let margin = 50.0;
        let line_height = style.font_size * 1.5;
        let page_width = 595.28;
        let page_height = 841.89;
        let usable_height = page_height - margin * 2.0;
        let max_lines_per_page = (usable_height / line_height) as usize;

        let mut doc = Document::new();
        doc.metadata = metadata;

        for chunk in all_paragraphs.chunks(max_lines_per_page.max(1)) {
            let mut page = Page::a4();
            let mut y = margin;

            for para in chunk {
                page.elements.push(PageElement::Text {
                    x: margin,
                    y,
                    width: page_width - margin * 2.0,
                    text: para.clone(),
                    style: style.clone(),
                    align: TextAlign::Left,
                });
                y += line_height;
            }

            doc.pages.push(page);
        }

        if doc.pages.is_empty() {
            doc.pages.push(Page::a4());
        }

        Ok(doc)
    }

    fn supported_extensions(&self) -> &[&str] {
        &["epub"]
    }

    fn format_name(&self) -> &str {
        "EPUB"
    }
}

/// ZIPアーカイブからエントリーを読み込む
fn read_zip_entry(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    path: &str,
) -> Result<String, ConvertError> {
    use std::io::Read;
    let mut file = archive
        .by_name(path)
        .map_err(|e| ConvertError::new("EPUB", &format!("{}が見つかりません: {}", path, e)))?;

    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| ConvertError::new("EPUB", &format!("{}の読み込みエラー: {}", path, e)))?;

    Ok(content)
}

/// container.xml から rootfile パスを抽出
fn parse_rootfile_path(xml: &str) -> Option<String> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e))
            | Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"rootfile" {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"full-path" {
                            return Some(
                                String::from_utf8_lossy(&attr.value).to_string(),
                            );
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    None
}

/// OPF からspine順のコンテンツファイルパスを取得
fn parse_spine_items(opf_xml: &str, base_dir: &str) -> Vec<String> {
    use std::collections::HashMap;

    let mut reader = quick_xml::Reader::from_str(opf_xml);
    let mut buf = Vec::new();

    // manifest: id → href
    let mut manifest: HashMap<String, String> = HashMap::new();
    // spine: idref 順序
    let mut spine_idrefs: Vec<String> = Vec::new();

    let mut in_manifest = false;
    let mut in_spine = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"manifest" => in_manifest = true,
                    b"spine" => in_spine = true,
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"item" && in_manifest {
                    let mut id = String::new();
                    let mut href = String::new();
                    let mut media_type = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"id" => id = String::from_utf8_lossy(&attr.value).to_string(),
                            b"href" => href = String::from_utf8_lossy(&attr.value).to_string(),
                            b"media-type" => {
                                media_type = String::from_utf8_lossy(&attr.value).to_string()
                            }
                            _ => {}
                        }
                    }
                    if media_type.contains("xhtml") || media_type.contains("html") {
                        manifest.insert(id, href);
                    }
                } else if local.as_ref() == b"itemref" && in_spine {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"idref" {
                            spine_idrefs
                                .push(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"manifest" => in_manifest = false,
                    b"spine" => in_spine = false,
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // spine 順で href パスを返す
    spine_idrefs
        .iter()
        .filter_map(|idref| {
            manifest.get(idref).map(|href| format!("{}{}", base_dir, href))
        })
        .collect()
}

/// XHTMLからテキスト段落を抽出
fn extract_xhtml_text(xhtml: &str) -> Vec<String> {
    let mut paragraphs = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xhtml);
    let mut buf = Vec::new();
    let mut current_text = String::new();
    let mut in_block = false;
    let mut depth: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");
                match name {
                    "p" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6" | "li" | "div" | "td"
                    | "th" | "dt" | "dd" | "blockquote" => {
                        if in_block && !current_text.trim().is_empty() {
                            paragraphs.push(current_text.trim().to_string());
                        }
                        current_text.clear();
                        in_block = true;
                        depth = 1;
                    }
                    _ => {
                        if in_block {
                            depth += 1;
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");
                match name {
                    "p" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6" | "li" | "div" | "td"
                    | "th" | "dt" | "dd" | "blockquote" => {
                        if !current_text.trim().is_empty() {
                            paragraphs.push(current_text.trim().to_string());
                        }
                        current_text.clear();
                        in_block = false;
                        depth = 0;
                    }
                    _ => {
                        if in_block && depth > 0 {
                            depth -= 1;
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) => {
                if in_block {
                    if let Ok(text) = e.unescape() {
                        if !current_text.is_empty() && !current_text.ends_with(' ') {
                            current_text.push(' ');
                        }
                        current_text.push_str(text.trim());
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    if !current_text.trim().is_empty() {
        paragraphs.push(current_text.trim().to_string());
    }

    paragraphs
}

/// EPUBメタデータを解析
fn parse_epub_metadata(opf_xml: &str) -> Metadata {
    let mut metadata = Metadata::default();
    let mut reader = quick_xml::Reader::from_str(opf_xml);
    let mut buf = Vec::new();
    let mut current_tag = String::new();
    let mut in_metadata = false;
    let mut in_tag = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");
                match name {
                    "metadata" => in_metadata = true,
                    _ if in_metadata => {
                        current_tag = name.to_string();
                        in_tag = true;
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) => {
                if in_tag && in_metadata {
                    if let Ok(text) = e.unescape() {
                        match current_tag.as_str() {
                            "title" => metadata.title = Some(text.to_string()),
                            "creator" => metadata.author = Some(text.to_string()),
                            "subject" => metadata.subject = Some(text.to_string()),
                            _ => {}
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"metadata" {
                    in_metadata = false;
                }
                in_tag = false;
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    metadata.creator = Some("WASM Document Converter".to_string());
    metadata
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_rootfile_path() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <container xmlns="urn:oasis:names:tc:opendocument:xmlns:container"
                   version="1.0">
            <rootfiles>
                <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
            </rootfiles>
        </container>"#;
        let path = parse_rootfile_path(xml);
        assert_eq!(path, Some("OEBPS/content.opf".to_string()));
    }

    #[test]
    fn test_extract_xhtml_text() {
        let xhtml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <html xmlns="http://www.w3.org/1999/xhtml">
        <body>
            <h1>Chapter 1</h1>
            <p>Hello World</p>
            <p>こんにちは</p>
        </body>
        </html>"#;
        let paragraphs = extract_xhtml_text(xhtml);
        assert!(paragraphs.iter().any(|p| p.contains("Chapter 1")));
        assert!(paragraphs.iter().any(|p| p.contains("Hello World")));
    }

    #[test]
    fn test_format_name() {
        let converter = EpubConverter::new();
        assert_eq!(converter.format_name(), "EPUB");
        assert_eq!(converter.supported_extensions(), &["epub"]);
    }
}
