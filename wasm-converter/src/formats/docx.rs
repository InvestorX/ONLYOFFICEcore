// formats/docx.rs - DOCX変換モジュール
//
// DOCX (Office Open XML) ファイルを解析し、テキストと基本的な構造を
// 抽出してドキュメントモデルに変換します。

use crate::converter::{
    ConvertError, Document, DocumentConverter, FontStyle, Metadata, Page, PageElement, TextAlign,
};

/// DOCXコンバーター
pub struct DocxConverter;

impl DocxConverter {
    pub fn new() -> Self {
        Self
    }
}

impl DocumentConverter for DocxConverter {
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError> {
        let cursor = std::io::Cursor::new(input);
        let mut archive = zip::ZipArchive::new(cursor)
            .map_err(|e| ConvertError::new("DOCX", &format!("ZIPアーカイブを開けません: {}", e)))?;

        // document.xmlを読み込む
        let doc_xml = read_zip_entry(&mut archive, "word/document.xml")?;

        // テキストを抽出
        let paragraphs = parse_docx_xml(&doc_xml)?;

        // メタデータの取得を試みる
        let metadata = read_docx_metadata(&mut archive);

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
        &["docx"]
    }

    fn format_name(&self) -> &str {
        "DOCX"
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
        .map_err(|e| ConvertError::new("DOCX", &format!("{}が見つかりません: {}", path, e)))?;

    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| ConvertError::new("DOCX", &format!("{}の読み込みエラー: {}", path, e)))?;

    Ok(content)
}

/// DOCX XMLからテキスト段落を抽出
fn parse_docx_xml(xml: &str) -> Result<Vec<String>, ConvertError> {
    let mut paragraphs = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut current_paragraph = String::new();
    let mut in_text = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"t" => in_text = true,
                    b"p" => current_paragraph.clear(),
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"t" => in_text = false,
                    b"p" => {
                        paragraphs.push(current_paragraph.clone());
                        current_paragraph.clear();
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) => {
                if in_text {
                    if let Ok(text) = e.unescape() {
                        current_paragraph.push_str(&text);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(e) => {
                return Err(ConvertError::new(
                    "DOCX",
                    &format!("XMLパースエラー: {}", e),
                ));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(paragraphs)
}

/// DOCXメタデータを読み取る
fn read_docx_metadata(archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> Metadata {
    let mut metadata = Metadata::default();

    if let Ok(core_xml) = read_zip_entry(archive, "docProps/core.xml") {
        let mut reader = quick_xml::Reader::from_str(&core_xml);
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
                                "creator" => metadata.author = Some(text.to_string()),
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
    fn test_parse_docx_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p><w:r><w:t>Hello World</w:t></w:r></w:p>
                <w:p><w:r><w:t>こんにちは</w:t></w:r></w:p>
            </w:body>
        </w:document>"#;

        let paragraphs = parse_docx_xml(xml).unwrap();
        assert!(paragraphs.iter().any(|p| p.contains("Hello World")));
        assert!(paragraphs.iter().any(|p| p.contains("こんにちは")));
    }
}
