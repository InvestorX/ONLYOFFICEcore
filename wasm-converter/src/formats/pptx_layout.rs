// formats/pptx.rs - PPTX変換モジュール（レイアウト保持版）
//
// PPTX (Office Open XML Presentation) ファイルを解析し、
// シェイプの位置・サイズ・書式・画像を忠実に再現してドキュメントモデルに変換します。
// Officeソフトで開いてPDF化するのと同等の出力を目指します。

use crate::converter::{
    Color, ConvertError, Document, DocumentConverter, FontStyle, Metadata, Page, PageElement, TextAlign,
};

/// EMU (English Metric Unit) → ポイント変換定数
/// 1インチ = 914400 EMU = 72 pt
const EMU_PER_PT: f64 = 914400.0 / 72.0;

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

        // スライドサイズをpresentation.xmlから取得
        let slide_size = read_slide_size(&mut archive);

        // スライドパスを検出
        let slide_paths = find_slide_paths(&mut archive);
        if slide_paths.is_empty() {
            return Err(ConvertError::new("PPTX", "スライドが見つかりません"));
        }

        // メタデータ
        let metadata = read_pptx_metadata(&mut archive);

        let mut doc = Document::new();
        doc.metadata = metadata;

        // 各スライドを処理
        for slide_path in &slide_paths {
            let slide_xml = read_zip_entry_string(&mut archive, slide_path)?;

            // スライドのリレーションシップを読み込む（画像参照解決用）
            let rels_path = slide_path
                .replace("ppt/slides/", "ppt/slides/_rels/")
                + ".rels";
            let rels = read_zip_entry_string(&mut archive, &rels_path).ok();

            // XMLからシェイプを解析
            let shapes = parse_slide_shapes(&slide_xml);

            // スライド背景色を検出
            let bg_color = parse_slide_background(&slide_xml);

            // 画像データを解決
            let mut resolved_shapes = Vec::new();
            for shape in shapes {
                if let ShapeContent::Image { r_id } = &shape.content {
                    if let Some(ref rels_xml) = rels {
                        if let Some(target) = resolve_relationship(rels_xml, r_id) {
                            let img_path = if target.starts_with('/') {
                                target[1..].to_string()
                            } else {
                                format!("ppt/slides/{}", target)
                            };
                            // Normalize path
                            let img_path = normalize_zip_path(&img_path);
                            if let Ok(data) = read_zip_entry_bytes(&mut archive, &img_path) {
                                let mime = if img_path.ends_with(".png") {
                                    "image/png"
                                } else if img_path.ends_with(".jpeg") || img_path.ends_with(".jpg") {
                                    "image/jpeg"
                                } else {
                                    "image/png"
                                };
                                let mut s = shape.clone();
                                s.content = ShapeContent::ImageData {
                                    data,
                                    mime_type: mime.to_string(),
                                };
                                resolved_shapes.push(s);
                                continue;
                            }
                        }
                    }
                }
                resolved_shapes.push(shape);
            }

            let page = render_slide_page(&resolved_shapes, &slide_size, bg_color);
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

/// スライドサイズ（ポイント単位）
#[derive(Debug, Clone, Copy)]
struct SlideSize {
    width: f64,
    height: f64,
}

impl Default for SlideSize {
    fn default() -> Self {
        // デフォルト: 10インチ × 7.5インチ (標準 4:3)
        Self {
            width: 720.0,
            height: 540.0,
        }
    }
}

/// シェイプの解析結果
#[derive(Debug, Clone)]
#[allow(dead_code)]
struct SlideShape {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
    content: ShapeContent,
    fill: Option<Color>,
    outline: Option<(Color, f64)>,
    rotation: f64,
}

/// シェイプの内容
#[derive(Debug, Clone)]
enum ShapeContent {
    TextBox { paragraphs: Vec<ShapeParagraph> },
    Image { r_id: String },
    ImageData { data: Vec<u8>, mime_type: String },
    Connector,
    Empty,
}

/// 段落
#[derive(Debug, Clone)]
struct ShapeParagraph {
    runs: Vec<TextRun>,
    align: TextAlign,
    bullet: Option<String>,
    level: u32,
}

/// テキストラン（書式付きテキスト断片）
#[derive(Debug, Clone)]
#[allow(dead_code)]
struct TextRun {
    text: String,
    font_size: f64,
    bold: bool,
    italic: bool,
    color: Option<Color>,
    font_name: Option<String>,
}

// ── ZIP helpers ──

fn read_zip_entry_string(
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

fn read_zip_entry_bytes(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    path: &str,
) -> Result<Vec<u8>, ConvertError> {
    use std::io::Read;
    let mut file = archive
        .by_name(path)
        .map_err(|e| ConvertError::new("PPTX", &format!("{}が見つかりません: {}", path, e)))?;
    let mut data = Vec::new();
    file.read_to_end(&mut data)
        .map_err(|e| ConvertError::new("PPTX", &format!("{}の読み込みエラー: {}", path, e)))?;
    Ok(data)
}

fn normalize_zip_path(path: &str) -> String {
    let parts: Vec<&str> = path.split('/').collect();
    let mut normalized = Vec::new();
    for part in parts {
        if part == ".." {
            normalized.pop();
        } else if part != "." && !part.is_empty() {
            normalized.push(part);
        }
    }
    normalized.join("/")
}

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
    paths.sort_by(|a, b| extract_slide_number(a).cmp(&extract_slide_number(b)));
    paths
}

fn extract_slide_number(path: &str) -> u32 {
    path.trim_start_matches("ppt/slides/slide")
        .trim_end_matches(".xml")
        .parse::<u32>()
        .unwrap_or(0)
}

// ── Presentation parsing ──

fn read_slide_size(archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> SlideSize {
    let pres_xml = match read_zip_entry_string(archive, "ppt/presentation.xml") {
        Ok(s) => s,
        Err(_) => return SlideSize::default(),
    };
    let mut reader = quick_xml::Reader::from_str(&pres_xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e))
            | Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"sldSz" {
                    let mut cx = 0u64;
                    let mut cy = 0u64;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"cx" => {
                                cx = String::from_utf8_lossy(&attr.value)
                                    .parse()
                                    .unwrap_or(0);
                            }
                            b"cy" => {
                                cy = String::from_utf8_lossy(&attr.value)
                                    .parse()
                                    .unwrap_or(0);
                            }
                            _ => {}
                        }
                    }
                    if cx > 0 && cy > 0 {
                        return SlideSize {
                            width: cx as f64 / EMU_PER_PT,
                            height: cy as f64 / EMU_PER_PT,
                        };
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    SlideSize::default()
}

// ── Slide XML parsing ──

/// スライドXMLからシェイプを完全解析
fn parse_slide_shapes(xml: &str) -> Vec<SlideShape> {
    let mut shapes = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();

    // State machine for parsing nested XML
    let mut depth = 0u32;
    let mut in_sp = false;       // <p:sp>
    let mut in_pic = false;      // <p:pic>
    let mut in_cxn = false;      // <p:cxnSp>
    let mut shape_depth = 0u32;

    // Current shape state
    let mut cur_x: f64 = 0.0;
    let mut cur_y: f64 = 0.0;
    let mut cur_w: f64 = 0.0;
    let mut cur_h: f64 = 0.0;
    let mut cur_fill: Option<Color> = None;
    let mut cur_outline: Option<(Color, f64)> = None;
    let mut cur_rotation: f64 = 0.0;
    let mut cur_paragraphs: Vec<ShapeParagraph> = Vec::new();
    let mut cur_runs: Vec<TextRun> = Vec::new();
    let mut cur_align = TextAlign::Left;
    let mut cur_bullet: Option<String> = None;
    let mut cur_level: u32 = 0;
    let mut cur_text = String::new();
    let mut cur_font_size: f64 = 18.0;
    let mut cur_bold = false;
    let mut cur_italic = false;
    let mut cur_color: Option<Color> = None;
    let mut cur_font_name: Option<String> = None;
    let mut in_text = false;
    let mut cur_r_id = String::new();  // image rId

    // For tracking sp offset/extent in xfrm
    let mut in_xfrm = false;
    let mut in_sp_pr = false;    // <p:spPr> or inner <a:...>
    let mut in_ln = false;       // <a:ln> (outline)
    let mut in_rpr = false;      // <a:rPr>
    let mut in_solid_fill = false;
    let mut solid_fill_ctx = 0u8; // 0=shape, 1=outline, 2=text

    macro_rules! reset_shape_state {
        () => {
            cur_x = 0.0;
            cur_y = 0.0;
            cur_w = 0.0;
            cur_h = 0.0;
            cur_fill = None;
            cur_outline = None;
            cur_rotation = 0.0;
            cur_paragraphs = Vec::new();
            cur_runs = Vec::new();
            cur_align = TextAlign::Left;
            cur_bullet = None;
            cur_level = 0;
            cur_text = String::new();
            cur_font_size = 18.0;
            cur_bold = false;
            cur_italic = false;
            cur_color = None;
            cur_font_name = None;
            in_text = false;
            cur_r_id = String::new();
            in_xfrm = false;
            in_sp_pr = false;
            in_ln = false;
            in_rpr = false;
            in_solid_fill = false;
            solid_fill_ctx = 0;
        };
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                depth += 1;
                let name = e.local_name();
                let local = name.as_ref();

                match local {
                    b"sp" if !in_sp && !in_pic && !in_cxn => {
                        in_sp = true;
                        shape_depth = depth;
                        reset_shape_state!();
                    }
                    b"pic" if !in_sp && !in_pic && !in_cxn => {
                        in_pic = true;
                        shape_depth = depth;
                        reset_shape_state!();
                    }
                    b"cxnSp" if !in_sp && !in_pic && !in_cxn => {
                        in_cxn = true;
                        shape_depth = depth;
                        reset_shape_state!();
                    }
                    b"spPr" if in_sp || in_pic || in_cxn => {
                        in_sp_pr = true;
                    }
                    b"xfrm" if in_sp_pr => {
                        in_xfrm = true;
                        // Check rotation attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"rot" {
                                cur_rotation = String::from_utf8_lossy(&attr.value)
                                    .parse::<f64>()
                                    .unwrap_or(0.0)
                                    / 60000.0; // 60000ths of a degree
                            }
                        }
                    }
                    b"ln" if in_sp_pr => {
                        in_ln = true;
                    }
                    b"solidFill" => {
                        in_solid_fill = true;
                        if in_rpr {
                            solid_fill_ctx = 2; // text color
                        } else if in_ln {
                            solid_fill_ctx = 1; // outline
                        } else if in_sp_pr {
                            solid_fill_ctx = 0; // shape fill
                        }
                    }
                    b"pPr" if (in_sp || in_pic) && !in_sp_pr => {
                        // Paragraph properties
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"algn" => {
                                    cur_align = match String::from_utf8_lossy(&attr.value).as_ref()
                                    {
                                        "ctr" => TextAlign::Center,
                                        "r" => TextAlign::Right,
                                        _ => TextAlign::Left,
                                    };
                                }
                                b"lvl" => {
                                    cur_level = String::from_utf8_lossy(&attr.value)
                                        .parse()
                                        .unwrap_or(0);
                                }
                                _ => {}
                            }
                        }
                    }
                    b"rPr" if (in_sp || in_pic) && !in_sp_pr => {
                        in_rpr = true;
                        // Run properties
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"sz" => {
                                    // Font size in hundredths of a point
                                    cur_font_size = String::from_utf8_lossy(&attr.value)
                                        .parse::<f64>()
                                        .unwrap_or(1800.0)
                                        / 100.0;
                                }
                                b"b" => {
                                    cur_bold =
                                        String::from_utf8_lossy(&attr.value).as_ref() == "1";
                                }
                                b"i" => {
                                    cur_italic =
                                        String::from_utf8_lossy(&attr.value).as_ref() == "1";
                                }
                                _ => {}
                            }
                        }
                    }
                    b"t" if (in_sp || in_pic) && !in_sp_pr => {
                        in_text = true;
                    }
                    b"p" if (in_sp || in_pic) && !in_sp_pr && depth > shape_depth + 1 => {
                        // New paragraph in text body
                        cur_runs.clear();
                        cur_align = TextAlign::Left;
                        cur_bullet = None;
                        cur_level = 0;
                        cur_font_size = 18.0;
                        cur_bold = false;
                        cur_italic = false;
                        cur_color = None;
                        cur_font_name = None;
                    }
                    b"blipFill" if in_pic => {
                        // Image fill - look for blip with r:embed
                    }
                    b"blip" if in_pic => {
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key.ends_with("embed") || key == "embed" {
                                cur_r_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    _ => {}
                }
            }

            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let name = e.local_name();
                let local = name.as_ref();

                if in_xfrm {
                    match local {
                        b"off" => {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"x" => {
                                        cur_x = String::from_utf8_lossy(&attr.value)
                                            .parse::<f64>()
                                            .unwrap_or(0.0)
                                            / EMU_PER_PT;
                                    }
                                    b"y" => {
                                        cur_y = String::from_utf8_lossy(&attr.value)
                                            .parse::<f64>()
                                            .unwrap_or(0.0)
                                            / EMU_PER_PT;
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"ext" => {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"cx" => {
                                        cur_w = String::from_utf8_lossy(&attr.value)
                                            .parse::<f64>()
                                            .unwrap_or(0.0)
                                            / EMU_PER_PT;
                                    }
                                    b"cy" => {
                                        cur_h = String::from_utf8_lossy(&attr.value)
                                            .parse::<f64>()
                                            .unwrap_or(0.0)
                                            / EMU_PER_PT;
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }

                // Color elements in solidFill
                if in_solid_fill {
                    let color = parse_color_element(e);
                    if let Some(c) = color {
                        match solid_fill_ctx {
                            0 => cur_fill = Some(c),
                            1 => {
                                cur_outline = Some((c, cur_outline.map_or(1.0, |o| o.1)));
                            }
                            2 => cur_color = Some(c),
                            _ => {}
                        }
                    }
                }

                // Run properties (empty element variant)
                if local == b"rPr" && (in_sp || in_pic) && !in_sp_pr {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"sz" => {
                                cur_font_size = String::from_utf8_lossy(&attr.value)
                                    .parse::<f64>()
                                    .unwrap_or(1800.0)
                                    / 100.0;
                            }
                            b"b" => {
                                cur_bold = String::from_utf8_lossy(&attr.value).as_ref() == "1";
                            }
                            b"i" => {
                                cur_italic = String::from_utf8_lossy(&attr.value).as_ref() == "1";
                            }
                            _ => {}
                        }
                    }
                }

                // Paragraph properties (empty variant)
                if local == b"pPr" && (in_sp || in_pic) && !in_sp_pr {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"algn" => {
                                cur_align =
                                    match String::from_utf8_lossy(&attr.value).as_ref() {
                                        "ctr" => TextAlign::Center,
                                        "r" => TextAlign::Right,
                                        _ => TextAlign::Left,
                                    };
                            }
                            b"lvl" => {
                                cur_level = String::from_utf8_lossy(&attr.value)
                                    .parse()
                                    .unwrap_or(0);
                            }
                            _ => {}
                        }
                    }
                }

                // Bullet markers
                if (local == b"buChar" || local == b"buAutoNum") && (in_sp || in_pic) {
                    if local == b"buChar" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"char" {
                                cur_bullet =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    } else {
                        cur_bullet = Some("•".to_string());
                    }
                }

                // Image blip (empty variant)
                if local == b"blip" && in_pic {
                    for attr in e.attributes().flatten() {
                        let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                        if key.ends_with("embed") || key == "embed" {
                            cur_r_id = String::from_utf8_lossy(&attr.value).to_string();
                        }
                    }
                }
            }

            Ok(quick_xml::events::Event::End(ref e)) => {
                let name = e.local_name();
                let local = name.as_ref();

                match local {
                    b"sp" if in_sp && depth == shape_depth => {
                        // Emit shape
                        let content = if cur_paragraphs.is_empty() && cur_runs.is_empty() {
                            ShapeContent::Empty
                        } else {
                            ShapeContent::TextBox {
                                paragraphs: cur_paragraphs.clone(),
                            }
                        };
                        shapes.push(SlideShape {
                            x: cur_x,
                            y: cur_y,
                            width: cur_w,
                            height: cur_h,
                            content,
                            fill: cur_fill,
                            outline: cur_outline,
                            rotation: cur_rotation,
                        });
                        in_sp = false;
                    }
                    b"pic" if in_pic && depth == shape_depth => {
                        let content = if !cur_r_id.is_empty() {
                            ShapeContent::Image {
                                r_id: cur_r_id.clone(),
                            }
                        } else {
                            ShapeContent::Empty
                        };
                        shapes.push(SlideShape {
                            x: cur_x,
                            y: cur_y,
                            width: cur_w,
                            height: cur_h,
                            content,
                            fill: cur_fill,
                            outline: cur_outline,
                            rotation: cur_rotation,
                        });
                        in_pic = false;
                    }
                    b"cxnSp" if in_cxn && depth == shape_depth => {
                        shapes.push(SlideShape {
                            x: cur_x,
                            y: cur_y,
                            width: cur_w,
                            height: cur_h,
                            content: ShapeContent::Connector,
                            fill: None,
                            outline: cur_outline.or(Some((Color::BLACK, 1.0))),
                            rotation: cur_rotation,
                        });
                        in_cxn = false;
                    }
                    b"spPr" => {
                        in_sp_pr = false;
                    }
                    b"xfrm" => {
                        in_xfrm = false;
                    }
                    b"ln" => {
                        in_ln = false;
                    }
                    b"solidFill" => {
                        in_solid_fill = false;
                    }
                    b"rPr" => {
                        in_rpr = false;
                    }
                    b"t" => {
                        if in_text {
                            // Finish text run
                            cur_runs.push(TextRun {
                                text: cur_text.clone(),
                                font_size: cur_font_size,
                                bold: cur_bold,
                                italic: cur_italic,
                                color: cur_color,
                                font_name: cur_font_name.clone(),
                            });
                            cur_text.clear();
                            in_text = false;
                        }
                    }
                    b"p" if (in_sp || in_pic) && !in_sp_pr && depth > shape_depth => {
                        // End paragraph
                        if !cur_runs.is_empty()
                            || cur_runs.iter().any(|r| !r.text.trim().is_empty())
                        {
                            cur_paragraphs.push(ShapeParagraph {
                                runs: cur_runs.clone(),
                                align: cur_align,
                                bullet: cur_bullet.clone(),
                                level: cur_level,
                            });
                        }
                        cur_runs.clear();
                        cur_bullet = None;
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

    shapes
}

/// XML要素から色を解析
fn parse_color_element(e: &quick_xml::events::BytesStart) -> Option<Color> {
    let local = e.local_name();
    match local.as_ref() {
        b"srgbClr" => {
            for attr in e.attributes().flatten() {
                if attr.key.as_ref() == b"val" {
                    return parse_hex_color(&String::from_utf8_lossy(&attr.value));
                }
            }
            None
        }
        b"schemeClr" => {
            // Scheme colors - return a reasonable default based on common themes
            for attr in e.attributes().flatten() {
                if attr.key.as_ref() == b"val" {
                    let scheme = String::from_utf8_lossy(&attr.value).to_string();
                    return Some(match scheme.as_str() {
                        "tx1" | "dk1" => Color::rgb(0, 0, 0),
                        "bg1" | "lt1" => Color::rgb(255, 255, 255),
                        "dk2" | "tx2" => Color::rgb(68, 84, 106),
                        "lt2" | "bg2" => Color::rgb(237, 237, 237),
                        "accent1" => Color::rgb(68, 114, 196),
                        "accent2" => Color::rgb(237, 125, 49),
                        "accent3" => Color::rgb(165, 165, 165),
                        "accent4" => Color::rgb(255, 192, 0),
                        "accent5" => Color::rgb(91, 155, 213),
                        "accent6" => Color::rgb(112, 173, 71),
                        _ => Color::rgb(0, 0, 0),
                    });
                }
            }
            None
        }
        _ => None,
    }
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

/// スライド背景色を解析
fn parse_slide_background(xml: &str) -> Option<Color> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut in_bg = false;
    let mut in_solid_fill = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"bg" {
                    in_bg = true;
                } else if local.as_ref() == b"solidFill" && in_bg {
                    in_solid_fill = true;
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                if in_solid_fill {
                    if let Some(c) = parse_color_element(e) {
                        return Some(c);
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"bg" {
                    in_bg = false;
                } else if local.as_ref() == b"solidFill" {
                    in_solid_fill = false;
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

/// リレーションシップXMLからrIdを解決
fn resolve_relationship(rels_xml: &str, r_id: &str) -> Option<String> {
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
                            b"Target" => {
                                target = String::from_utf8_lossy(&attr.value).to_string()
                            }
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

// ── Page rendering ──

/// 解析済みシェイプからページを構築
fn render_slide_page(
    shapes: &[SlideShape],
    slide_size: &SlideSize,
    bg_color: Option<Color>,
) -> Page {
    let mut page = Page {
        width: slide_size.width,
        height: slide_size.height,
        elements: Vec::new(),
    };

    // 背景
    if let Some(bg) = bg_color {
        page.elements.push(PageElement::Rect {
            x: 0.0,
            y: 0.0,
            width: slide_size.width,
            height: slide_size.height,
            fill: Some(bg),
            stroke: None,
            stroke_width: 0.0,
        });
    }

    for shape in shapes {
        match &shape.content {
            ShapeContent::TextBox { paragraphs } => {
                // Shape fill rectangle
                if let Some(fill) = shape.fill {
                    page.elements.push(PageElement::Rect {
                        x: shape.x,
                        y: shape.y,
                        width: shape.width,
                        height: shape.height,
                        fill: Some(fill),
                        stroke: None,
                        stroke_width: 0.0,
                    });
                }

                // Shape outline
                if let Some((color, width)) = shape.outline {
                    page.elements.push(PageElement::Rect {
                        x: shape.x,
                        y: shape.y,
                        width: shape.width,
                        height: shape.height,
                        fill: None,
                        stroke: Some(color),
                        stroke_width: width,
                    });
                }

                // Render text paragraphs positioned within the shape
                let margin = 4.0;
                let mut text_y = shape.y + margin;

                for para in paragraphs {
                    let indent = para.level as f64 * 18.0;

                    // Concatenate all runs to get full paragraph text
                    let mut full_text = String::new();
                    if let Some(ref bullet) = para.bullet {
                        full_text.push_str(bullet);
                        full_text.push(' ');
                    }

                    // Use first run's style for paragraph
                    let first_run = para.runs.first();
                    let font_size = first_run.map_or(18.0, |r| r.font_size);
                    let bold = first_run.map_or(false, |r| r.bold);
                    let italic = first_run.map_or(false, |r| r.italic);
                    let color = first_run.and_then(|r| r.color).unwrap_or(Color::BLACK);

                    for run in &para.runs {
                        full_text.push_str(&run.text);
                    }

                    if !full_text.trim().is_empty() {
                        let line_height = font_size * 1.3;

                        // Word-wrap the text within the shape width
                        let available_width = shape.width - margin * 2.0 - indent;
                        let lines = wrap_text(&full_text, available_width, font_size);

                        for line in &lines {
                            if text_y + font_size > shape.y + shape.height {
                                break; // Clip to shape bounds
                            }

                            page.elements.push(PageElement::Text {
                                x: shape.x + margin + indent,
                                y: text_y,
                                width: available_width,
                                text: line.clone(),
                                style: FontStyle {
                                    font_size,
                                    bold,
                                    italic,
                                    color,
                                    ..FontStyle::default()
                                },
                                align: para.align,
                            });
                            text_y += line_height;
                        }
                    } else {
                        // Empty paragraph - add line spacing
                        text_y += font_size * 0.8;
                    }
                }
            }

            ShapeContent::ImageData { data, mime_type } => {
                page.elements.push(PageElement::Image {
                    x: shape.x,
                    y: shape.y,
                    width: shape.width,
                    height: shape.height,
                    data: data.clone(),
                    mime_type: mime_type.clone(),
                });
            }

            ShapeContent::Image { .. } => {
                // Unresolved image - draw placeholder
                page.elements.push(PageElement::Rect {
                    x: shape.x,
                    y: shape.y,
                    width: shape.width,
                    height: shape.height,
                    fill: Some(Color::rgb(230, 230, 230)),
                    stroke: Some(Color::rgb(180, 180, 180)),
                    stroke_width: 0.5,
                });
                page.elements.push(PageElement::Text {
                    x: shape.x + 4.0,
                    y: shape.y + shape.height / 2.0 - 6.0,
                    width: shape.width - 8.0,
                    text: "[Image]".to_string(),
                    style: FontStyle {
                        font_size: 10.0,
                        color: Color::rgb(150, 150, 150),
                        italic: true,
                        ..FontStyle::default()
                    },
                    align: TextAlign::Center,
                });
            }

            ShapeContent::Connector => {
                // Draw line from top-left to bottom-right
                let color = shape
                    .outline
                    .map(|(c, _)| c)
                    .unwrap_or(Color::rgb(0, 0, 0));
                let width = shape.outline.map(|(_, w)| w).unwrap_or(1.0);
                page.elements.push(PageElement::Line {
                    x1: shape.x,
                    y1: shape.y,
                    x2: shape.x + shape.width,
                    y2: shape.y + shape.height,
                    width,
                    color,
                });
            }

            ShapeContent::Empty => {
                // Shape with fill but no content
                if let Some(fill) = shape.fill {
                    page.elements.push(PageElement::Rect {
                        x: shape.x,
                        y: shape.y,
                        width: shape.width,
                        height: shape.height,
                        fill: Some(fill),
                        stroke: shape.outline.map(|(c, _)| c),
                        stroke_width: shape.outline.map(|(_, w)| w).unwrap_or(0.0),
                    });
                } else if let Some((color, width)) = shape.outline {
                    page.elements.push(PageElement::Rect {
                        x: shape.x,
                        y: shape.y,
                        width: shape.width,
                        height: shape.height,
                        fill: None,
                        stroke: Some(color),
                        stroke_width: width,
                    });
                }
            }
        }
    }

    page
}

/// テキストをシェイプ幅に合わせて折り返す
fn wrap_text(text: &str, available_width: f64, font_size: f64) -> Vec<String> {
    if text.is_empty() {
        return vec![];
    }

    // Approximate character width: CJK ≈ font_size, Latin ≈ 0.5 * font_size
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

/// PPTXメタデータを読み取る
fn read_pptx_metadata(archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> Metadata {
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
    fn test_emu_to_pt_conversion() {
        // 1 inch = 914400 EMU = 72 pt
        let emu = 914400.0;
        let pt = emu / EMU_PER_PT;
        assert!((pt - 72.0).abs() < 0.01);
    }

    #[test]
    fn test_parse_hex_color() {
        assert_eq!(
            parse_hex_color("FF0000"),
            Some(Color::rgb(255, 0, 0))
        );
        assert_eq!(
            parse_hex_color("0000FF"),
            Some(Color::rgb(0, 0, 255))
        );
    }

    #[test]
    fn test_wrap_text() {
        let lines = wrap_text("Hello World", 100.0, 12.0);
        assert!(!lines.is_empty());
    }

    #[test]
    fn test_extract_slide_number() {
        assert_eq!(extract_slide_number("ppt/slides/slide1.xml"), 1);
        assert_eq!(extract_slide_number("ppt/slides/slide10.xml"), 10);
    }

    #[test]
    fn test_default_slide_size() {
        let ss = SlideSize::default();
        assert!((ss.width - 720.0).abs() < 0.01);
        assert!((ss.height - 540.0).abs() < 0.01);
    }
}
