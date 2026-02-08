// formats/odt.rs - ODT変換モジュール
//
// ODT (OpenDocument Text) ファイルを解析し、テキストと基本的な構造を
// 抽出してドキュメントモデルに変換します。

use crate::converter::{
    ConvertError, Document, DocumentConverter, FontStyle, Metadata, Page, PageElement, TextAlign,
};

/// ODTコンバーター
pub struct OdtConverter;

impl OdtConverter {
    pub fn new() -> Self {
        Self
    }
}

impl DocumentConverter for OdtConverter {
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError> {
        let cursor = std::io::Cursor::new(input);
        let mut archive = zip::ZipArchive::new(cursor)
            .map_err(|e| ConvertError::new("ODT", &format!("ZIPアーカイブを開けません: {}", e)))?;

        // content.xml を読み込む
        let content_xml = read_zip_entry(&mut archive, "content.xml")?;

        // テキスト段落を抽出
        let paragraphs = parse_odt_content(&content_xml);

        // メタデータの取得
        let metadata = read_odt_metadata(&mut archive);

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

        for chunk in paragraphs.chunks(max_lines_per_page.max(1)) {
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
        &["odt"]
    }

    fn format_name(&self) -> &str {
        "ODT"
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
        .map_err(|e| ConvertError::new("ODT", &format!("{}が見つかりません: {}", path, e)))?;

    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| ConvertError::new("ODT", &format!("{}の読み込みエラー: {}", path, e)))?;

    Ok(content)
}

/// ODT content.xml からテキスト段落を抽出
fn parse_odt_content(xml: &str) -> Vec<String> {
    let mut paragraphs = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut current_paragraph = String::new();
    let mut in_text = false;
    let mut in_paragraph = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"p" | b"h" => {
                        in_paragraph = true;
                        current_paragraph.clear();
                    }
                    b"span" if in_paragraph => {
                        in_text = true;
                    }
                    _ => {
                        if in_paragraph {
                            in_text = true;
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"p" | b"h" => {
                        paragraphs.push(current_paragraph.clone());
                        current_paragraph.clear();
                        in_paragraph = false;
                        in_text = false;
                    }
                    b"span" => {
                        in_text = false;
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) => {
                if in_paragraph {
                    if let Ok(text) = e.unescape() {
                        current_paragraph.push_str(&text);
                    }
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local_name = e.local_name();
                if local_name.as_ref() == b"tab" && in_paragraph {
                    current_paragraph.push('\t');
                } else if local_name.as_ref() == b"line-break" && in_paragraph {
                    current_paragraph.push('\n');
                } else if local_name.as_ref() == b"s" && in_paragraph {
                    current_paragraph.push(' ');
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // 空行のみの段落をフィルタリング（先頭・末尾の空段落を除去）
    let _ = in_text; // suppress unused warning
    paragraphs
}

/// ODTメタデータを読み取る
fn read_odt_metadata(archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> Metadata {
    let mut metadata = Metadata::default();

    if let Ok(meta_xml) = read_zip_entry(archive, "meta.xml") {
        let mut reader = quick_xml::Reader::from_str(&meta_xml);
        let mut buf = Vec::new();
        let mut current_tag = String::new();
        let mut in_tag = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let name = String::from_utf8_lossy(e.local_name().as_ref()).to_string();
                    current_tag = name;
                    in_tag = true;
                }
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    if in_tag {
                        if let Ok(text) = e.unescape() {
                            match current_tag.as_str() {
                                "title" => metadata.title = Some(text.to_string()),
                                "initial-creator" | "creator" => {
                                    metadata.author = Some(text.to_string())
                                }
                                "subject" => metadata.subject = Some(text.to_string()),
                                _ => {}
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::End(_)) => {
                    in_tag = false;
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    metadata.creator = Some("WASM Document Converter".to_string());
    metadata
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_odt_content() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                                 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <office:body>
                <office:text>
                    <text:p>Hello World</text:p>
                    <text:p>こんにちは</text:p>
                </office:text>
            </office:body>
        </office:document-content>"#;

        let paragraphs = parse_odt_content(xml);
        assert!(paragraphs.iter().any(|p| p.contains("Hello World")));
        assert!(paragraphs.iter().any(|p| p.contains("こんにちは")));
    }

    #[test]
    fn test_format_name() {
        let converter = OdtConverter::new();
        assert_eq!(converter.format_name(), "ODT");
        assert_eq!(converter.supported_extensions(), &["odt"]);
    }
}
