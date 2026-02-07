// formats/docx_layout.rs - DOCX変換モジュール（レイアウト保持版）
//
// DOCX (Office Open XML) ファイルを解析し、
// 段落の書式・テーブル・画像・ページマージンを忠実に再現して
// ドキュメントモデルに変換します。

use crate::converter::{
    Color, ConvertError, Document, DocumentConverter, FontStyle, Metadata, Page, PageElement, TextAlign,
};

/// DOCXコンバーター（レイアウト保持版）
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

        // document.xml を読み込む
        let doc_xml = read_zip_entry_string(&mut archive, "word/document.xml")?;

        // リレーションシップを読み込む（画像解決用）
        let rels = read_zip_entry_string(&mut archive, "word/_rels/document.xml.rels").ok();

        // メタデータ
        let metadata = read_docx_metadata(&mut archive);

        // ページ設定を解析
        let page_setup = parse_section_properties(&doc_xml);

        // ドキュメント本文を解析
        let body_elements = parse_document_body(&doc_xml);

        // 画像を解決
        let resolved_elements = resolve_images(&body_elements, &rels, &mut archive);

        // ページにレイアウト
        let pages = layout_pages(&resolved_elements, &page_setup);

        let mut doc = Document::new();
        doc.metadata = metadata;
        doc.pages = pages;

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

// ── 型定義 ──

/// ページ設定
#[derive(Debug, Clone)]
struct PageSetup {
    width: f64,      // ポイント
    height: f64,
    margin_top: f64,
    margin_bottom: f64,
    margin_left: f64,
    margin_right: f64,
}

impl Default for PageSetup {
    fn default() -> Self {
        Self {
            width: 595.28,      // A4
            height: 841.89,
            margin_top: 72.0,   // 1インチ
            margin_bottom: 72.0,
            margin_left: 72.0,
            margin_right: 72.0,
        }
    }
}

/// TWIP → ポイント変換 (1 pt = 20 twip)
const TWIP_PER_PT: f64 = 20.0;

/// ドキュメント本文要素
#[derive(Debug, Clone)]
enum BodyElement {
    Paragraph(DocParagraph),
    Table(DocTable),
}

/// 段落
#[derive(Debug, Clone)]
#[allow(dead_code)]
struct DocParagraph {
    runs: Vec<DocRun>,
    align: TextAlign,
    spacing_before: f64,   // ポイント
    spacing_after: f64,
    line_spacing: f64,     // 倍率（1.0 = シングル）
    indent_left: f64,      // ポイント
    indent_first: f64,     // 最初の行のインデント
    is_heading: bool,
    heading_level: u32,
    numbering: Option<String>,
}

/// テキストラン
#[derive(Debug, Clone)]
#[allow(dead_code)]
struct DocRun {
    content: RunContent,
    font_size: f64,
    bold: bool,
    italic: bool,
    underline: bool,
    color: Color,
    font_name: Option<String>,
    highlight: Option<Color>,
}

#[derive(Debug, Clone)]
enum RunContent {
    Text(String),
    Image { r_id: String },
    ImageData { data: Vec<u8>, mime_type: String, width: f64, height: f64 },
    LineBreak,
    Tab,
}

/// テーブル
#[derive(Debug, Clone)]
#[allow(dead_code)]
struct DocTable {
    rows: Vec<DocTableRow>,
    column_widths: Vec<f64>,
    border_color: Color,
}

#[derive(Debug, Clone)]
struct DocTableRow {
    cells: Vec<DocTableCell>,
    height: Option<f64>,
}

#[derive(Debug, Clone)]
struct DocTableCell {
    paragraphs: Vec<DocParagraph>,
    width: f64,
    shading: Option<Color>,
}

// ── ZIP helpers ──

fn read_zip_entry_string(
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

fn read_zip_entry_bytes(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    path: &str,
) -> Result<Vec<u8>, ConvertError> {
    use std::io::Read;
    let mut file = archive
        .by_name(path)
        .map_err(|e| ConvertError::new("DOCX", &format!("{}が見つかりません: {}", path, e)))?;
    let mut data = Vec::new();
    file.read_to_end(&mut data)
        .map_err(|e| ConvertError::new("DOCX", &format!("{}の読み込みエラー: {}", path, e)))?;
    Ok(data)
}

// ── セクションプロパティ解析 ──

fn parse_section_properties(xml: &str) -> PageSetup {
    let mut setup = PageSetup::default();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e))
            | Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"pgSz" => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"w" => {
                                    setup.width = parse_twip(&attr.value);
                                }
                                b"h" => {
                                    setup.height = parse_twip(&attr.value);
                                }
                                _ => {}
                            }
                        }
                    }
                    b"pgMar" => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"top" => setup.margin_top = parse_twip(&attr.value),
                                b"bottom" => setup.margin_bottom = parse_twip(&attr.value),
                                b"left" => setup.margin_left = parse_twip(&attr.value),
                                b"right" => setup.margin_right = parse_twip(&attr.value),
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    setup
}

fn parse_twip(val: &[u8]) -> f64 {
    let s = String::from_utf8_lossy(val);
    s.parse::<f64>().unwrap_or(0.0) / TWIP_PER_PT
}

// ── ドキュメント本文解析 ──

fn parse_document_body(xml: &str) -> Vec<BodyElement> {
    let mut elements = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();

    let mut depth = 0u32;
    let mut in_body = false;
    let mut in_paragraph = false;
    let mut para_depth = 0u32;
    let mut in_table = false;
    let mut table_depth = 0u32;

    // Paragraph state
    let mut cur_runs: Vec<DocRun> = Vec::new();
    let mut cur_align = TextAlign::Left;
    let mut cur_spacing_before = 0.0f64;
    let mut cur_spacing_after = 8.0f64;  // Default Word spacing
    let mut cur_line_spacing = 1.15f64;
    let mut cur_indent_left = 0.0f64;
    let mut cur_indent_first = 0.0f64;
    let mut cur_is_heading = false;
    let mut cur_heading_level = 0u32;
    let mut cur_numbering: Option<String> = None;

    // Run state
    let mut cur_font_size = 11.0f64;
    let mut cur_bold = false;
    let mut cur_italic = false;
    let mut cur_underline = false;
    let mut cur_color = Color::BLACK;
    let mut cur_font_name: Option<String> = None;
    let mut cur_highlight: Option<Color> = None;
    let mut in_run = false;
    let mut in_rpr = false;
    let mut in_ppr = false;
    let mut in_text = false;
    let mut cur_text = String::new();

    // Table state
    let mut tbl_rows: Vec<DocTableRow> = Vec::new();
    let mut tbl_col_widths: Vec<f64> = Vec::new();
    let mut in_tbl_row = false;
    let mut cur_cells: Vec<DocTableCell> = Vec::new();
    let mut cur_row_height: Option<f64> = None;
    let mut in_tbl_cell = false;
    let mut cell_paragraphs: Vec<DocParagraph> = Vec::new();
    let mut cell_width = 0.0f64;
    let mut cell_shading: Option<Color> = None;

    // Image state
    let mut in_drawing = false;
    let mut drawing_r_id = String::new();
    let mut _drawing_cx = 0.0f64;
    let mut _drawing_cy = 0.0f64;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                depth += 1;
                let local = e.local_name();
                match local.as_ref() {
                    b"body" => in_body = true,
                    b"p" if in_body && !in_table => {
                        in_paragraph = true;
                        para_depth = depth;
                        cur_runs.clear();
                        cur_align = TextAlign::Left;
                        cur_spacing_before = 0.0;
                        cur_spacing_after = 8.0;
                        cur_line_spacing = 1.15;
                        cur_indent_left = 0.0;
                        cur_indent_first = 0.0;
                        cur_is_heading = false;
                        cur_heading_level = 0;
                        cur_numbering = None;
                    }
                    b"p" if in_tbl_cell => {
                        in_paragraph = true;
                        para_depth = depth;
                        cur_runs.clear();
                        cur_align = TextAlign::Left;
                        cur_spacing_before = 0.0;
                        cur_spacing_after = 0.0;
                        cur_line_spacing = 1.0;
                        cur_indent_left = 0.0;
                        cur_indent_first = 0.0;
                        cur_is_heading = false;
                        cur_heading_level = 0;
                        cur_numbering = None;
                    }
                    b"pPr" if in_paragraph => {
                        in_ppr = true;
                    }
                    b"r" if in_paragraph => {
                        in_run = true;
                        // Reset run state to paragraph defaults
                        cur_font_size = 11.0;
                        cur_bold = false;
                        cur_italic = false;
                        cur_underline = false;
                        cur_color = Color::BLACK;
                        cur_font_name = None;
                        cur_highlight = None;
                    }
                    b"rPr" if in_run => {
                        in_rpr = true;
                    }
                    b"t" if in_run => {
                        in_text = true;
                        cur_text.clear();
                    }
                    b"drawing" if in_run => {
                        in_drawing = true;
                        drawing_r_id.clear();
                        _drawing_cx = 72.0;
                        _drawing_cy = 72.0;
                    }
                    b"tbl" if in_body => {
                        in_table = true;
                        table_depth = depth;
                        tbl_rows.clear();
                        tbl_col_widths.clear();
                    }
                    b"tr" if in_table => {
                        in_tbl_row = true;
                        cur_cells.clear();
                        cur_row_height = None;
                    }
                    b"tc" if in_tbl_row => {
                        in_tbl_cell = true;
                        cell_paragraphs.clear();
                        cell_width = 0.0;
                        cell_shading = None;
                    }
                    _ => {}
                }
            }

            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    // Paragraph properties
                    b"jc" if in_ppr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                cur_align = match String::from_utf8_lossy(&attr.value).as_ref() {
                                    "center" => TextAlign::Center,
                                    "right" => TextAlign::Right,
                                    _ => TextAlign::Left,
                                };
                            }
                        }
                    }
                    b"spacing" if in_ppr => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"before" => {
                                    cur_spacing_before = parse_twip(&attr.value);
                                }
                                b"after" => {
                                    cur_spacing_after = parse_twip(&attr.value);
                                }
                                b"line" => {
                                    let val = String::from_utf8_lossy(&attr.value)
                                        .parse::<f64>()
                                        .unwrap_or(240.0);
                                    cur_line_spacing = val / 240.0;
                                }
                                _ => {}
                            }
                        }
                    }
                    b"ind" if in_ppr => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"left" => {
                                    cur_indent_left = parse_twip(&attr.value);
                                }
                                b"firstLine" => {
                                    cur_indent_first = parse_twip(&attr.value);
                                }
                                b"hanging" => {
                                    cur_indent_first = -parse_twip(&attr.value);
                                }
                                _ => {}
                            }
                        }
                    }
                    b"pStyle" if in_ppr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                let style = String::from_utf8_lossy(&attr.value).to_string();
                                if style.starts_with("Heading") || style.starts_with("heading") {
                                    cur_is_heading = true;
                                    cur_heading_level = style
                                        .chars()
                                        .filter(|c| c.is_ascii_digit())
                                        .collect::<String>()
                                        .parse()
                                        .unwrap_or(1);
                                }
                            }
                        }
                    }
                    // Run properties
                    b"sz" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                // Half-points → points
                                cur_font_size = String::from_utf8_lossy(&attr.value)
                                    .parse::<f64>()
                                    .unwrap_or(22.0)
                                    / 2.0;
                            }
                        }
                    }
                    b"b" if in_rpr => {
                        cur_bold = true;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                cur_bold = String::from_utf8_lossy(&attr.value).as_ref() != "0";
                            }
                        }
                    }
                    b"i" if in_rpr => {
                        cur_italic = true;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                cur_italic = String::from_utf8_lossy(&attr.value).as_ref() != "0";
                            }
                        }
                    }
                    b"u" if in_rpr => {
                        cur_underline = true;
                    }
                    b"color" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                let hex = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(c) = parse_hex_color(&hex) {
                                    cur_color = c;
                                }
                            }
                        }
                    }
                    b"rFonts" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.local_name();
                            if key.as_ref() == b"ascii" || key.as_ref() == b"eastAsia"
                                || key.as_ref() == b"hAnsi"
                            {
                                cur_font_name =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                                break;
                            }
                        }
                    }
                    b"highlight" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                cur_highlight = highlight_name_to_color(
                                    &String::from_utf8_lossy(&attr.value),
                                );
                            }
                        }
                    }
                    // Line break
                    b"br" if in_run => {
                        cur_runs.push(DocRun {
                            content: RunContent::LineBreak,
                            font_size: cur_font_size,
                            bold: cur_bold,
                            italic: cur_italic,
                            underline: cur_underline,
                            color: cur_color,
                            font_name: cur_font_name.clone(),
                            highlight: cur_highlight,
                        });
                    }
                    // Tab
                    b"tab" if in_run => {
                        cur_runs.push(DocRun {
                            content: RunContent::Tab,
                            font_size: cur_font_size,
                            bold: cur_bold,
                            italic: cur_italic,
                            underline: cur_underline,
                            color: cur_color,
                            font_name: cur_font_name.clone(),
                            highlight: cur_highlight,
                        });
                    }
                    // Table column widths
                    b"gridCol" if in_table && !in_tbl_row => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"w" {
                                tbl_col_widths.push(parse_twip(&attr.value));
                            }
                        }
                    }
                    // Table cell width
                    b"tcW" if in_tbl_cell => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"w" {
                                cell_width = parse_twip(&attr.value);
                            }
                        }
                    }
                    // Cell shading
                    b"shd" if in_tbl_cell => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"fill" {
                                let hex = String::from_utf8_lossy(&attr.value).to_string();
                                if hex != "auto" {
                                    cell_shading = parse_hex_color(&hex);
                                }
                            }
                        }
                    }
                    // Row height
                    b"trHeight" if in_tbl_row => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                cur_row_height = Some(parse_twip(&attr.value));
                            }
                        }
                    }
                    // Image blip in drawing
                    b"blip" if in_drawing => {
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key.ends_with("embed") {
                                drawing_r_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    // Image extent
                    b"ext" if in_drawing => {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"cx" => {
                                    _drawing_cx = String::from_utf8_lossy(&attr.value)
                                        .parse::<f64>()
                                        .unwrap_or(0.0)
                                        / 12700.0; // EMU to pt
                                }
                                b"cy" => {
                                    _drawing_cy = String::from_utf8_lossy(&attr.value)
                                        .parse::<f64>()
                                        .unwrap_or(0.0)
                                        / 12700.0;
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }

            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"body" => in_body = false,
                    b"t" => {
                        if in_text {
                            cur_runs.push(DocRun {
                                content: RunContent::Text(cur_text.clone()),
                                font_size: cur_font_size,
                                bold: cur_bold,
                                italic: cur_italic,
                                underline: cur_underline,
                                color: cur_color,
                                font_name: cur_font_name.clone(),
                                highlight: cur_highlight,
                            });
                            cur_text.clear();
                            in_text = false;
                        }
                    }
                    b"r" => {
                        in_run = false;
                    }
                    b"rPr" => {
                        in_rpr = false;
                    }
                    b"pPr" => {
                        in_ppr = false;
                    }
                    b"drawing" => {
                        if in_drawing && !drawing_r_id.is_empty() {
                            cur_runs.push(DocRun {
                                content: RunContent::Image { r_id: drawing_r_id.clone() },
                                font_size: cur_font_size,
                                bold: false,
                                italic: false,
                                underline: false,
                                color: Color::BLACK,
                                font_name: None,
                                highlight: None,
                            });
                        }
                        in_drawing = false;
                    }
                    b"p" if in_paragraph && depth == para_depth => {
                        let para = DocParagraph {
                            runs: cur_runs.clone(),
                            align: cur_align,
                            spacing_before: cur_spacing_before,
                            spacing_after: cur_spacing_after,
                            line_spacing: cur_line_spacing,
                            indent_left: cur_indent_left,
                            indent_first: cur_indent_first,
                            is_heading: cur_is_heading,
                            heading_level: cur_heading_level,
                            numbering: cur_numbering.clone(),
                        };

                        if in_tbl_cell {
                            cell_paragraphs.push(para);
                        } else {
                            elements.push(BodyElement::Paragraph(para));
                        }
                        in_paragraph = false;
                    }
                    b"tc" if in_tbl_cell => {
                        cur_cells.push(DocTableCell {
                            paragraphs: cell_paragraphs.clone(),
                            width: cell_width,
                            shading: cell_shading,
                        });
                        in_tbl_cell = false;
                        cell_paragraphs.clear();
                    }
                    b"tr" if in_tbl_row => {
                        tbl_rows.push(DocTableRow {
                            cells: cur_cells.clone(),
                            height: cur_row_height,
                        });
                        in_tbl_row = false;
                    }
                    b"tbl" if in_table && depth == table_depth => {
                        elements.push(BodyElement::Table(DocTable {
                            rows: tbl_rows.clone(),
                            column_widths: tbl_col_widths.clone(),
                            border_color: Color::rgb(0, 0, 0),
                        }));
                        in_table = false;
                    }
                    _ => {}
                }
                depth -= 1;
            }

            Ok(quick_xml::events::Event::Text(ref e)) => {
                if in_text {
                    if let Ok(text) = e.unescape() {
                        cur_text.push_str(&text);
                    }
                }
            }

            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    elements
}

/// 画像リレーションシップを解決
fn resolve_images(
    elements: &[BodyElement],
    rels: &Option<String>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> Vec<BodyElement> {
    elements
        .iter()
        .map(|elem| match elem {
            BodyElement::Paragraph(para) => {
                let resolved_runs: Vec<DocRun> = para
                    .runs
                    .iter()
                    .map(|run| {
                        if let RunContent::Image { r_id } = &run.content {
                            if let Some(ref rels_xml) = rels {
                                if let Some(target) = resolve_rel(rels_xml, r_id) {
                                    let img_path = format!("word/{}", target);
                                    if let Ok(data) = read_zip_entry_bytes(archive, &img_path) {
                                        let mime = if img_path.ends_with(".png") {
                                            "image/png"
                                        } else {
                                            "image/jpeg"
                                        };
                                        return DocRun {
                                            content: RunContent::ImageData {
                                                data,
                                                mime_type: mime.to_string(),
                                                width: 200.0, // Default; EMU parsed above
                                                height: 150.0,
                                            },
                                            ..run.clone()
                                        };
                                    }
                                }
                            }
                        }
                        run.clone()
                    })
                    .collect();
                BodyElement::Paragraph(DocParagraph {
                    runs: resolved_runs,
                    ..para.clone()
                })
            }
            other => other.clone(),
        })
        .collect()
}

fn resolve_rel(rels_xml: &str, r_id: &str) -> Option<String> {
    let mut reader = quick_xml::Reader::from_str(rels_xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e))
            | Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut id = String::new();
                    let mut target = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Id" => id = String::from_utf8_lossy(&attr.value).to_string(),
                            b"Target" => target = String::from_utf8_lossy(&attr.value).to_string(),
                            _ => {}
                        }
                    }
                    if id == r_id {
                        return Some(target);
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

// ── ページレイアウト ──

fn layout_pages(elements: &[BodyElement], setup: &PageSetup) -> Vec<Page> {
    let mut pages = Vec::new();
    let usable_width = setup.width - setup.margin_left - setup.margin_right;
    let usable_height = setup.height - setup.margin_top - setup.margin_bottom;

    let mut cur_y = 0.0f64; // Current Y position relative to content area

    let new_page = || -> Page {
        Page {
            width: setup.width,
            height: setup.height,
            elements: Vec::new(),
        }
    };

    let mut page = new_page();

    for element in elements {
        match element {
            BodyElement::Paragraph(para) => {
                // Apply heading styles
                let (base_font_size, is_bold) = if para.is_heading {
                    match para.heading_level {
                        1 => (20.0, true),
                        2 => (16.0, true),
                        3 => (14.0, true),
                        _ => (12.0, true),
                    }
                } else {
                    (11.0, false)
                };

                // Spacing before
                cur_y += para.spacing_before;

                // Calculate paragraph height
                let first_run_size = para.runs.iter().find_map(|r| {
                    if let RunContent::Text(_) = &r.content {
                        Some(r.font_size)
                    } else {
                        None
                    }
                }).unwrap_or(base_font_size);

                let effective_font_size = if para.is_heading { base_font_size } else { first_run_size };
                let line_height = effective_font_size * para.line_spacing;

                // Check if we need a new page
                if cur_y + line_height > usable_height && !page.elements.is_empty() {
                    pages.push(page);
                    page = new_page();
                    cur_y = 0.0;
                }

                // Render paragraph runs
                let _abs_y = setup.margin_top + cur_y;
                let abs_x = setup.margin_left + para.indent_left;

                // Concatenate all text runs for this line, respecting formatting
                let mut line_x = abs_x + para.indent_first.max(0.0);
                let mut _first_line = true;

                for run in &para.runs {
                    match &run.content {
                        RunContent::Text(text) => {
                            if text.is_empty() {
                                continue;
                            }

                            let font_size = if para.is_heading { base_font_size } else { run.font_size };
                            let bold = run.bold || is_bold;

                            // Word wrap within available width
                            let available = setup.margin_left + usable_width - line_x;
                            let lines = wrap_text_width(text, available, font_size);

                            for (li, line_text) in lines.iter().enumerate() {
                                if cur_y + font_size > usable_height {
                                    pages.push(page);
                                    page = new_page();
                                    cur_y = 0.0;
                                    line_x = abs_x;
                                }

                                let text_y = setup.margin_top + cur_y;

                                // Highlight background
                                if let Some(hl_color) = run.highlight {
                                    let text_width = estimate_text_width(line_text, font_size);
                                    page.elements.push(PageElement::Rect {
                                        x: line_x,
                                        y: text_y,
                                        width: text_width,
                                        height: font_size * 1.2,
                                        fill: Some(hl_color),
                                        stroke: None,
                                        stroke_width: 0.0,
                                    });
                                }

                                page.elements.push(PageElement::Text {
                                    x: line_x,
                                    y: text_y,
                                    width: available,
                                    text: line_text.clone(),
                                    style: FontStyle {
                                        font_size,
                                        bold,
                                        italic: run.italic,
                                        color: run.color,
                                        ..FontStyle::default()
                                    },
                                    align: para.align,
                                });

                                if li < lines.len() - 1 {
                                    cur_y += line_height;
                                    line_x = abs_x;
                                    _first_line = false;
                                } else {
                                    // Update x position for next inline run
                                    line_x += estimate_text_width(line_text, font_size);
                                }
                            }
                        }
                        RunContent::ImageData { data, mime_type, width, height } => {
                            // Constrain image to page width
                            let max_w = usable_width;
                            let (img_w, img_h) = if *width > max_w {
                                let scale = max_w / width;
                                (max_w, height * scale)
                            } else {
                                (*width, *height)
                            };

                            if cur_y + img_h > usable_height {
                                pages.push(page);
                                page = new_page();
                                cur_y = 0.0;
                            }

                            page.elements.push(PageElement::Image {
                                x: setup.margin_left,
                                y: setup.margin_top + cur_y,
                                width: img_w,
                                height: img_h,
                                data: data.clone(),
                                mime_type: mime_type.clone(),
                            });
                            cur_y += img_h + 4.0;
                            line_x = abs_x;
                        }
                        RunContent::LineBreak => {
                            cur_y += line_height;
                            line_x = abs_x;
                        }
                        RunContent::Tab => {
                            line_x += 36.0; // ~0.5 inch tab
                        }
                        RunContent::Image { .. } => {
                            // Unresolved - skip
                        }
                    }
                }

                // Move to next line after paragraph
                cur_y += line_height;

                // Spacing after
                cur_y += para.spacing_after;
            }

            BodyElement::Table(table) => {
                // Calculate table dimensions
                let total_width: f64 = if table.column_widths.is_empty() {
                    usable_width
                } else {
                    table.column_widths.iter().sum()
                };
                let col_count = table.rows.first().map_or(1, |r| r.cells.len()).max(1);
                let default_col_width = total_width / col_count as f64;

                let row_height = 20.0;

                for tbl_row in &table.rows {
                    let rh = tbl_row.height.unwrap_or(row_height);

                    // Check page break
                    if cur_y + rh > usable_height {
                        pages.push(page);
                        page = new_page();
                        cur_y = 0.0;
                    }

                    let abs_y = setup.margin_top + cur_y;
                    let mut cell_x = setup.margin_left;

                    for (ci, cell) in tbl_row.cells.iter().enumerate() {
                        let cw = if ci < table.column_widths.len() {
                            table.column_widths[ci]
                        } else if cell.width > 0.0 {
                            cell.width
                        } else {
                            default_col_width
                        };

                        // Cell background
                        if let Some(shading) = cell.shading {
                            page.elements.push(PageElement::Rect {
                                x: cell_x,
                                y: abs_y,
                                width: cw,
                                height: rh,
                                fill: Some(shading),
                                stroke: None,
                                stroke_width: 0.0,
                            });
                        }

                        // Cell border
                        page.elements.push(PageElement::Rect {
                            x: cell_x,
                            y: abs_y,
                            width: cw,
                            height: rh,
                            fill: None,
                            stroke: Some(Color::rgb(0, 0, 0)),
                            stroke_width: 0.5,
                        });

                        // Cell text
                        let mut text_y = abs_y + 3.0;
                        for cp in &cell.paragraphs {
                            for run in &cp.runs {
                                if let RunContent::Text(text) = &run.content {
                                    if !text.is_empty() {
                                        page.elements.push(PageElement::Text {
                                            x: cell_x + 3.0,
                                            y: text_y,
                                            width: cw - 6.0,
                                            text: text.clone(),
                                            style: FontStyle {
                                                font_size: run.font_size.min(rh - 4.0),
                                                bold: run.bold,
                                                italic: run.italic,
                                                color: run.color,
                                                ..FontStyle::default()
                                            },
                                            align: cp.align,
                                        });
                                        text_y += run.font_size * 1.2;
                                    }
                                }
                            }
                        }

                        cell_x += cw;
                    }

                    cur_y += rh;
                }

                cur_y += 8.0; // Table spacing
            }
        }
    }

    // Push final page
    if !page.elements.is_empty() {
        pages.push(page);
    }

    pages
}

/// テキスト幅の推定
fn estimate_text_width(text: &str, font_size: f64) -> f64 {
    text.chars()
        .map(|ch| {
            if ch.is_ascii() {
                font_size * 0.5
            } else {
                font_size * 1.0
            }
        })
        .sum()
}

/// テキストを利用可能幅で折り返す
fn wrap_text_width(text: &str, available_width: f64, font_size: f64) -> Vec<String> {
    if text.is_empty() || available_width <= 0.0 {
        return vec![text.to_string()];
    }

    let mut lines = Vec::new();
    let mut current_line = String::new();
    let mut current_width = 0.0;

    for ch in text.chars() {
        let char_width = if ch.is_ascii() {
            font_size * 0.5
        } else {
            font_size * 1.0
        };

        if current_width + char_width > available_width && !current_line.is_empty() {
            lines.push(current_line.clone());
            current_line.clear();
            current_width = 0.0;
        }

        current_line.push(ch);
        current_width += char_width;
    }

    if !current_line.is_empty() {
        lines.push(current_line);
    }

    if lines.is_empty() {
        lines.push(String::new());
    }

    lines
}

fn parse_hex_color(hex: &str) -> Option<Color> {
    if hex.len() == 6 {
        let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
        let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
        let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
        Some(Color::rgb(r, g, b))
    } else {
        None
    }
}

fn highlight_name_to_color(name: &str) -> Option<Color> {
    Some(match name {
        "yellow" => Color::rgb(255, 255, 0),
        "green" => Color::rgb(0, 255, 0),
        "cyan" => Color::rgb(0, 255, 255),
        "magenta" => Color::rgb(255, 0, 255),
        "blue" => Color::rgb(0, 0, 255),
        "red" => Color::rgb(255, 0, 0),
        "darkBlue" => Color::rgb(0, 0, 139),
        "darkCyan" => Color::rgb(0, 139, 139),
        "darkGreen" => Color::rgb(0, 100, 0),
        "darkMagenta" => Color::rgb(139, 0, 139),
        "darkRed" => Color::rgb(139, 0, 0),
        "darkYellow" => Color::rgb(139, 139, 0),
        "darkGray" => Color::rgb(169, 169, 169),
        "lightGray" => Color::rgb(211, 211, 211),
        "black" => Color::rgb(0, 0, 0),
        _ => return None,
    })
}

/// DOCXメタデータを読み取る
fn read_docx_metadata(archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> Metadata {
    let mut metadata = Metadata::default();

    if let Ok(core_xml) = read_zip_entry_string(archive, "docProps/core.xml") {
        let mut reader = quick_xml::Reader::from_str(&core_xml);
        let mut buf = Vec::new();
        let mut current_tag = String::new();
        let mut in_tag = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    current_tag =
                        String::from_utf8_lossy(e.local_name().as_ref()).to_string();
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
                Ok(quick_xml::events::Event::End(_)) => in_tag = false,
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
    fn test_parse_section_properties() {
        let xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:sectPr>
              <w:pgSz w:w="12240" w:h="15840"/>
              <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            </w:sectPr>
          </w:body>
        </w:document>"#;
        let setup = parse_section_properties(xml);
        assert!((setup.width - 612.0).abs() < 0.1); // Letter width
        assert!((setup.margin_top - 72.0).abs() < 0.1); // 1 inch
    }

    #[test]
    fn test_parse_formatted_paragraph() {
        let xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr><w:jc w:val="center"/></w:pPr>
              <w:r>
                <w:rPr>
                  <w:b/>
                  <w:sz w:val="28"/>
                  <w:color w:val="FF0000"/>
                </w:rPr>
                <w:t>Bold Red Title</w:t>
              </w:r>
            </w:p>
            <w:p>
              <w:r><w:t>Normal text</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>"#;
        let elements = parse_document_body(xml);
        assert_eq!(elements.len(), 2);

        if let BodyElement::Paragraph(para) = &elements[0] {
            assert!(matches!(para.align, TextAlign::Center));
            assert!(!para.runs.is_empty());
            let run = &para.runs[0];
            assert!(run.bold);
            assert!((run.font_size - 14.0).abs() < 0.1); // 28 half-points = 14pt
            assert_eq!(run.color.r, 255);
        } else {
            panic!("Expected paragraph");
        }
    }

    #[test]
    fn test_wrap_text() {
        let lines = wrap_text_width("Hello World Test", 50.0, 12.0);
        assert!(lines.len() >= 1);
    }
}
