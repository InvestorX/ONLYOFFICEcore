// formats/pptx.rs - PPTX変換モジュール
//
// PPTX (Office Open XML Presentation) ファイルを解析し、
// 各スライドのテキストと基本的な構造を抽出してドキュメントモデルに変換します。

use crate::converter::{
    Color, ConvertError, Document, DocumentConverter, FontStyle, Metadata, Page, PageElement,
    TextAlign,
};

/// PPTXコンバーター
pub struct PptxConverter;

impl PptxConverter {
    pub fn new() -> Self {
        Self
    }
}

impl DocumentConverter for PptxConverter {
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError> {
        let cursor = std::io::Cursor::new(input);
        let mut archive = zip::ZipArchive::new(cursor)
            .map_err(|e| ConvertError::new("PPTX", &format!("ZIPアーカイブを開けません: {}", e)))?;

        // スライド数を特定
        let slide_paths = find_slide_paths(&mut archive);

        if slide_paths.is_empty() {
            return Err(ConvertError::new(
                "PPTX",
                "スライドが見つかりません",
            ));
        }

        // メタデータの取得を試みる
        let metadata = read_pptx_metadata(&mut archive);

        let mut doc = Document::new();
        doc.metadata = metadata;

        // 各スライドを処理
        for (slide_index, slide_path) in slide_paths.iter().enumerate() {
            let slide_xml = read_zip_entry(&mut archive, slide_path)?;
            let slide_texts = parse_slide_xml(&slide_xml);

            let page = render_slide_to_page(slide_index + 1, &slide_texts);
            doc.pages.push(page);
        }

        if doc.pages.is_empty() {
            doc.pages.push(Page::a4());
        }

        Ok(doc)
    }

    fn supported_extensions(&self) -> &[&str] {
        &["pptx"]
    }

    fn format_name(&self) -> &str {
        "PPTX"
    }
}

/// ZIPアーカイブからスライドパスを検出
fn find_slide_paths(archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> Vec<String> {
    let mut paths = Vec::new();
    for i in 0..archive.len() {
        if let Ok(file) = archive.by_index(i) {
            let name = file.name().to_string();
            if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
                paths.push(name);
            }
        }
    }
    // スライド番号順にソート
    paths.sort_by(|a, b| {
        let num_a = extract_slide_number(a);
        let num_b = extract_slide_number(b);
        num_a.cmp(&num_b)
    });
    paths
}

/// スライドファイル名から番号を抽出
fn extract_slide_number(path: &str) -> u32 {
    path.trim_start_matches("ppt/slides/slide")
        .trim_end_matches(".xml")
        .parse::<u32>()
        .unwrap_or(0)
}

/// ZIPアーカイブからエントリーを読み込む
fn read_zip_entry(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    path: &str,
) -> Result<String, ConvertError> {
    use std::io::Read;
    let mut file = archive
        .by_name(path)
        .map_err(|e| ConvertError::new("PPTX", &format!("{}が見つかりません: {}", path, e)))?;

    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| ConvertError::new("PPTX", &format!("{}の読み込みエラー: {}", path, e)))?;

    Ok(content)
}

/// スライドXMLからテキスト要素を抽出
/// PPTX XMLでは<a:t>タグにテキストが含まれ、<a:p>で段落を表す
fn parse_slide_xml(xml: &str) -> Vec<SlideText> {
    let mut texts = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut current_paragraph = String::new();
    let mut in_text = false;
    let mut is_title = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) | Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"t" => in_text = true,
                    b"p" => {
                        current_paragraph.clear();
                        is_title = false;
                    }
                    b"ph" => {
                        // プレースホルダー型をチェック（タイトル、サブタイトル等）
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"type" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val == "title" || val == "ctrTitle" || val == "subTitle" {
                                    is_title = true;
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"t" => in_text = false,
                    b"p" => {
                        if !current_paragraph.trim().is_empty() {
                            texts.push(SlideText {
                                text: current_paragraph.clone(),
                                is_title,
                            });
                        }
                        current_paragraph.clear();
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) => {
                if in_text {
                    if let Ok(text) = e.unescape() {
                        if !current_paragraph.is_empty() {
                            current_paragraph.push(' ');
                        }
                        current_paragraph.push_str(&text);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    texts
}

/// スライドテキスト情報
struct SlideText {
    text: String,
    is_title: bool,
}

/// スライドをページに変換
fn render_slide_to_page(slide_number: usize, texts: &[SlideText]) -> Page {
    // PPTXはランドスケープが一般的（16:9 → A4横向きに近い比率で表示）
    let mut page = Page {
        width: 841.89,  // A4横
        height: 595.28,
        elements: Vec::new(),
    };

    let margin = 50.0;
    let usable_width = page.width - margin * 2.0;

    // スライド番号ヘッダー
    page.elements.push(PageElement::Text {
        x: margin,
        y: margin,
        width: usable_width,
        text: format!("スライド {}", slide_number),
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
    });

    let mut y = margin + 40.0;

    for slide_text in texts {
        if slide_text.is_title {
            // タイトルテキスト
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
            // 通常テキスト
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

        // ページ内に収まるように制限
        if y > page.height - margin - 30.0 {
            break;
        }
    }

    // テキストが空の場合
    if texts.is_empty() {
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

/// PPTXメタデータを読み取る
fn read_pptx_metadata(archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> Metadata {
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
    fn test_parse_slide_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                        <p:txBody>
                            <a:p><a:r><a:t>テストタイトル</a:t></a:r></a:p>
                        </p:txBody>
                    </p:sp>
                    <p:sp>
                        <p:txBody>
                            <a:p><a:r><a:t>本文テキスト1</a:t></a:r></a:p>
                            <a:p><a:r><a:t>本文テキスト2</a:t></a:r></a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"#;

        let texts = parse_slide_xml(xml);
        assert!(!texts.is_empty());
        assert!(texts.iter().any(|t| t.text.contains("テストタイトル")));
        assert!(texts.iter().any(|t| t.text.contains("本文テキスト1")));
    }

    #[test]
    fn test_extract_slide_number() {
        assert_eq!(extract_slide_number("ppt/slides/slide1.xml"), 1);
        assert_eq!(extract_slide_number("ppt/slides/slide10.xml"), 10);
        assert_eq!(extract_slide_number("ppt/slides/slide2.xml"), 2);
    }

    #[test]
    fn test_render_slide_to_page() {
        let texts = vec![
            SlideText { text: "タイトル".to_string(), is_title: true },
            SlideText { text: "内容".to_string(), is_title: false },
        ];
        let page = render_slide_to_page(1, &texts);
        assert!(!page.elements.is_empty());
        // ランドスケープサイズ
        assert!(page.width > page.height);
    }
}
