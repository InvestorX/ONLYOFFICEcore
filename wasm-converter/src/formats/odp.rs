// formats/odp.rs - ODP変換モジュール
//
// ODP (OpenDocument Presentation) ファイルを解析し、
// 各スライドのテキストを抽出してドキュメントモデルに変換します。

use crate::converter::{
    Color, ConvertError, Document, DocumentConverter, FontStyle, Metadata, Page, PageElement,
    TextAlign,
};

/// ODPコンバーター
pub struct OdpConverter;

impl OdpConverter {
    pub fn new() -> Self {
        Self
    }
}

impl DocumentConverter for OdpConverter {
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError> {
        let cursor = std::io::Cursor::new(input);
        let mut archive = zip::ZipArchive::new(cursor)
            .map_err(|e| ConvertError::new("ODP", &format!("ZIPアーカイブを開けません: {}", e)))?;

        // content.xml を読み込む
        let content_xml = read_zip_entry(&mut archive, "content.xml")?;

        // スライドを抽出
        let slides = parse_odp_slides(&content_xml);

        // メタデータ
        let metadata = read_odp_metadata(&mut archive);

        let mut doc = Document::new();
        doc.metadata = metadata;

        if slides.is_empty() {
            doc.pages.push(Page::a4());
            return Ok(doc);
        }

        for (i, slide) in slides.iter().enumerate() {
            let page = render_slide_to_page(i + 1, slide);
            doc.pages.push(page);
        }

        Ok(doc)
    }

    fn supported_extensions(&self) -> &[&str] {
        &["odp"]
    }

    fn format_name(&self) -> &str {
        "ODP"
    }
}

/// スライド情報
struct OdpSlide {
    name: String,
    texts: Vec<SlideText>,
}

/// スライドテキスト
struct SlideText {
    text: String,
    is_title: bool,
}

/// ZIPアーカイブからエントリーを読み込む
fn read_zip_entry(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    path: &str,
) -> Result<String, ConvertError> {
    use std::io::Read;
    let mut file = archive
        .by_name(path)
        .map_err(|e| ConvertError::new("ODP", &format!("{}が見つかりません: {}", path, e)))?;

    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| ConvertError::new("ODP", &format!("{}の読み込みエラー: {}", path, e)))?;

    Ok(content)
}

/// ODP content.xml からスライドを解析
fn parse_odp_slides(xml: &str) -> Vec<OdpSlide> {
    let mut slides = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();

    let mut in_page = false;
    let mut in_frame = false;
    let mut in_text_box = false;
    let mut in_paragraph = false;
    let mut current_paragraph = String::new();
    let mut current_texts: Vec<SlideText> = Vec::new();
    let mut page_name = String::new();
    let mut is_title_frame = false;
    let mut frame_index = 0u32;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"page" => {
                        in_page = true;
                        current_texts.clear();
                        frame_index = 0;
                        page_name = String::new();
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"name" {
                                page_name =
                                    String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    b"frame" if in_page => {
                        in_frame = true;
                        // First frame is typically the title in ODP
                        is_title_frame = frame_index == 0;
                        // Check presentation:class attribute for title
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key.ends_with("class") {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val == "title" || val == "subtitle" {
                                    is_title_frame = true;
                                }
                            }
                        }
                    }
                    b"text-box" if in_frame => {
                        in_text_box = true;
                    }
                    b"p" | b"h" if in_text_box => {
                        in_paragraph = true;
                        current_paragraph.clear();
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"page" => {
                        if in_page {
                            slides.push(OdpSlide {
                                name: page_name.clone(),
                                texts: std::mem::take(&mut current_texts),
                            });
                        }
                        in_page = false;
                    }
                    b"frame" => {
                        if in_frame {
                            frame_index += 1;
                        }
                        in_frame = false;
                        in_text_box = false;
                    }
                    b"text-box" => {
                        in_text_box = false;
                    }
                    b"p" | b"h" if in_paragraph => {
                        if !current_paragraph.trim().is_empty() {
                            current_texts.push(SlideText {
                                text: current_paragraph.trim().to_string(),
                                is_title: is_title_frame,
                            });
                        }
                        current_paragraph.clear();
                        in_paragraph = false;
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
                let local = e.local_name();
                if local.as_ref() == b"tab" && in_paragraph {
                    current_paragraph.push('\t');
                } else if local.as_ref() == b"line-break" && in_paragraph {
                    current_paragraph.push('\n');
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    slides
}

/// スライドをページに変換
fn render_slide_to_page(slide_number: usize, slide: &OdpSlide) -> Page {
    let mut page = Page {
        width: 841.89,  // A4横
        height: 595.28,
        elements: Vec::new(),
    };

    let margin = 50.0;
    let usable_width = page.width - margin * 2.0;

    // スライド番号ヘッダー
    let slide_label = if slide.name.is_empty() {
        format!("スライド {}", slide_number)
    } else {
        format!("スライド {} - {}", slide_number, slide.name)
    };

    page.elements.push(PageElement::Text {
        x: margin,
        y: margin,
        width: usable_width,
        text: slide_label,
        style: FontStyle {
            font_size: 10.0,
            color: Color::rgb(128, 128, 128),
            ..FontStyle::default()
        },
        align: TextAlign::Right,
    });

    // 枠線
    page.elements.push(PageElement::Rect {
        x: margin - 5.0,
        y: margin + 15.0,
        width: usable_width + 10.0,
        height: page.height - margin * 2.0 - 20.0,
        fill: None,
        stroke: Some(Color::rgb(200, 200, 200)),
        stroke_width: 1.0,
        rotation_deg: 0.0,
    });

    let mut y = margin + 40.0;

    for slide_text in &slide.texts {
        if slide_text.is_title {
            page.elements.push(PageElement::Text {
                x: margin + 20.0,
                y,
                width: usable_width - 40.0,
                text: slide_text.text.clone(),
                style: FontStyle {
                    font_size: 24.0,
                    bold: true,
                    ..FontStyle::default()
                },
                align: TextAlign::Center,
            });
            y += 40.0;
        } else {
            page.elements.push(PageElement::Text {
                x: margin + 30.0,
                y,
                width: usable_width - 60.0,
                text: slide_text.text.clone(),
                style: FontStyle {
                    font_size: 14.0,
                    ..FontStyle::default()
                },
                align: TextAlign::Left,
            });
            y += 22.0;
        }

        if y > page.height - margin - 30.0 {
            break;
        }
    }

    // テキストが空の場合
    if slide.texts.is_empty() {
        page.elements.push(PageElement::Text {
            x: margin + 20.0,
            y: page.height / 2.0 - 20.0,
            width: usable_width - 40.0,
            text: format!("(スライド {} - テキストなし)", slide_number),
            style: FontStyle {
                font_size: 16.0,
                color: Color::rgb(160, 160, 160),
                italic: true,
                ..FontStyle::default()
            },
            align: TextAlign::Center,
        });
    }

    page
}

/// ODPメタデータを読み取る
fn read_odp_metadata(archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> Metadata {
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
    fn test_parse_odp_slides() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <office:document-content
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
            xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
            xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <office:body>
                <office:presentation>
                    <draw:page draw:name="Slide1">
                        <draw:frame presentation:class="title">
                            <draw:text-box>
                                <text:p>タイトル</text:p>
                            </draw:text-box>
                        </draw:frame>
                        <draw:frame>
                            <draw:text-box>
                                <text:p>本文テキスト</text:p>
                            </draw:text-box>
                        </draw:frame>
                    </draw:page>
                </office:presentation>
            </office:body>
        </office:document-content>"#;

        let slides = parse_odp_slides(xml);
        assert_eq!(slides.len(), 1);
        assert!(slides[0].texts.iter().any(|t| t.text.contains("タイトル")));
        assert!(slides[0].texts.iter().any(|t| t.text.contains("本文テキスト")));
    }

    #[test]
    fn test_render_slide_to_page() {
        let slide = OdpSlide {
            name: "Test".to_string(),
            texts: vec![
                SlideText {
                    text: "タイトル".to_string(),
                    is_title: true,
                },
                SlideText {
                    text: "内容".to_string(),
                    is_title: false,
                },
            ],
        };
        let page = render_slide_to_page(1, &slide);
        assert!(!page.elements.is_empty());
        assert!(page.width > page.height); // ランドスケープ
    }

    #[test]
    fn test_format_name() {
        let converter = OdpConverter::new();
        assert_eq!(converter.format_name(), "ODP");
        assert_eq!(converter.supported_extensions(), &["odp"]);
    }
}
