// formats/pptx.rs - PPTX変換モジュール（レイアウト保持版）
//
// PPTX (Office Open XML Presentation) ファイルを解析し、
// シェイプの位置・サイズ・書式・画像を忠実に再現してドキュメントモデルに変換します。
// Officeソフトで開いてPDF化するのと同等の出力を目指します。

use crate::converter::{
    Color, ConvertError, Document, DocumentConverter, FontStyle, GradientStop, GradientType,
    Metadata, Page, PageElement, PathCommand, TextAlign,
};

/// EMU (English Metric Unit) → ポイント変換定数
/// 1インチ = 914400 EMU = 72 pt
const EMU_PER_PT: f64 = 914400.0 / 72.0;

/// 3D効果の押し出し深度（ポイント単位）
const SHAPE_3D_EXTRUSION_DEPTH: f64 = 6.0;

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

        // テーマカラーを読み込む
        let theme_colors = read_theme_colors(&mut archive);

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

            // XMLからシェイプを解析（グループシェイプも含む）
            let shapes = parse_slide_shapes(&slide_xml, &theme_colors);

            // スライド背景を解析（画像・グラデーション含む）
            let bg = parse_slide_background_full(&slide_xml, &rels, &mut archive, &theme_colors);

            // 画像データを解決
            let mut resolved_shapes = Vec::new();
            for shape in shapes {
                let resolved = resolve_shape_images(shape, &rels, &mut archive);
                resolved_shapes.push(resolved);
            }

            // チャート参照を検出して描画要素を収集
            let chart_elements = detect_and_render_charts(
                &slide_xml, &rels, &mut archive, &slide_size,
            );

            // SmartArt/ダイアグラム参照を検出して描画要素を収集
            let smartart_elements = detect_and_render_smartart(
                &slide_xml, &rels, &mut archive, &slide_size,
            );

            let mut page = render_slide_page(&resolved_shapes, &slide_size, bg.as_ref());

            // チャート要素を追加
            page.elements.extend(chart_elements);

            // SmartArt要素を追加
            page.elements.extend(smartart_elements);

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

/// スライド背景
#[derive(Debug, Clone)]
enum SlideBg {
    Solid(Color),
    Gradient {
        stops: Vec<GradientStop>,
        angle: f64, // radians
    },
    Image {
        data: Vec<u8>,
        mime_type: String,
    },
}

/// テーマカラーマップ
#[derive(Debug, Clone)]
struct ThemeColors {
    dk1: Color,
    lt1: Color,
    dk2: Color,
    lt2: Color,
    accent1: Color,
    accent2: Color,
    accent3: Color,
    accent4: Color,
    accent5: Color,
    accent6: Color,
    hlink: Color,
    fol_hlink: Color,
}

impl Default for ThemeColors {
    fn default() -> Self {
        Self {
            dk1: Color::rgb(0, 0, 0),
            lt1: Color::rgb(255, 255, 255),
            dk2: Color::rgb(68, 84, 106),
            lt2: Color::rgb(231, 230, 230),
            accent1: Color::rgb(91, 155, 213),
            accent2: Color::rgb(237, 125, 49),
            accent3: Color::rgb(165, 165, 165),
            accent4: Color::rgb(255, 192, 0),
            accent5: Color::rgb(68, 114, 196),
            accent6: Color::rgb(112, 173, 71),
            hlink: Color::rgb(5, 99, 193),
            fol_hlink: Color::rgb(149, 79, 114),
        }
    }
}

impl ThemeColors {
    fn resolve(&self, scheme_name: &str) -> Color {
        match scheme_name {
            "tx1" | "dk1" => self.dk1,
            "bg1" | "lt1" => self.lt1,
            "dk2" | "tx2" => self.dk2,
            "lt2" | "bg2" => self.lt2,
            "accent1" => self.accent1,
            "accent2" => self.accent2,
            "accent3" => self.accent3,
            "accent4" => self.accent4,
            "accent5" => self.accent5,
            "accent6" => self.accent6,
            "hlink" => self.hlink,
            "folHlink" => self.fol_hlink,
            _ => Color::rgb(0, 0, 0),
        }
    }
}

/// シェイプ塗りつぶし
#[derive(Debug, Clone)]
#[allow(dead_code)]
enum ShapeFill {
    Solid(Color),
    Gradient {
        stops: Vec<GradientStop>,
        angle: f64,
    },
    Image {
        data: Vec<u8>,
        mime_type: String,
    },
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
    fill: Option<ShapeFill>,
    outline: Option<(Color, f64)>,
    rotation: f64,
    shadow: Option<ShadowEffect>,
    has_3d: bool,
    preset_geometry: Option<String>,
    custom_path: Option<Vec<crate::converter::PathCommand>>,
    custom_path_viewport: Option<(f64, f64)>,
    /// blipFill r:embed on the shape (resolved later)
    fill_image_r_id: Option<String>,
    /// Text body margins in points (from bodyPr lIns, tIns, rIns, bIns)
    text_margin_left: f64,
    text_margin_top: f64,
    text_margin_right: f64,
    text_margin_bottom: f64,
}

/// シャドウ効果
#[derive(Debug, Clone)]
#[allow(dead_code)]
struct ShadowEffect {
    color: Color,
    blur_radius: f64,
    offset_x: f64,
    offset_y: f64,
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

/// テーマカラーをtheme.xmlから読み込む
fn read_theme_colors(archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> ThemeColors {
    let theme_xml = match read_zip_entry_string(archive, "ppt/theme/theme1.xml") {
        Ok(s) => s,
        Err(_) => return ThemeColors::default(),
    };
    let mut reader = quick_xml::Reader::from_str(&theme_xml);
    let mut buf = Vec::new();
    let mut colors = ThemeColors::default();
    let mut current_scheme_entry = String::new();
    let mut in_clr_scheme = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"clrScheme" => in_clr_scheme = true,
                    b"dk1" | b"lt1" | b"dk2" | b"lt2" | b"accent1" | b"accent2" | b"accent3"
                    | b"accent4" | b"accent5" | b"accent6" | b"hlink" | b"folHlink"
                        if in_clr_scheme =>
                    {
                        current_scheme_entry =
                            String::from_utf8_lossy(local.as_ref()).to_string();
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                if in_clr_scheme && !current_scheme_entry.is_empty() {
                    let local = e.local_name();
                    let color = match local.as_ref() {
                        b"srgbClr" => {
                            e.attributes().flatten().find(|a| a.key.as_ref() == b"val")
                                .and_then(|a| parse_hex_color(&String::from_utf8_lossy(&a.value)))
                        }
                        b"sysClr" => {
                            e.attributes().flatten().find(|a| a.key.as_ref() == b"lastClr")
                                .and_then(|a| parse_hex_color(&String::from_utf8_lossy(&a.value)))
                        }
                        _ => None,
                    };
                    if let Some(c) = color {
                        match current_scheme_entry.as_str() {
                            "dk1" => colors.dk1 = c,
                            "lt1" => colors.lt1 = c,
                            "dk2" => colors.dk2 = c,
                            "lt2" => colors.lt2 = c,
                            "accent1" => colors.accent1 = c,
                            "accent2" => colors.accent2 = c,
                            "accent3" => colors.accent3 = c,
                            "accent4" => colors.accent4 = c,
                            "accent5" => colors.accent5 = c,
                            "accent6" => colors.accent6 = c,
                            "hlink" => colors.hlink = c,
                            "folHlink" => colors.fol_hlink = c,
                            _ => {}
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"clrScheme" {
                    in_clr_scheme = false;
                }
                if in_clr_scheme {
                    match local.as_ref() {
                        b"dk1" | b"lt1" | b"dk2" | b"lt2" | b"accent1" | b"accent2"
                        | b"accent3" | b"accent4" | b"accent5" | b"accent6" | b"hlink"
                        | b"folHlink" => {
                            current_scheme_entry.clear();
                        }
                        _ => {}
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    colors
}

/// シェイプの画像参照を解決
fn resolve_shape_images(
    shape: SlideShape,
    rels: &Option<String>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> SlideShape {
    let mut s = shape;

    // Resolve content image (pic element)
    if let ShapeContent::Image { ref r_id } = s.content {
        if let Some(ref rels_xml) = rels {
            if let Some(target) = resolve_relationship(rels_xml, r_id) {
                let img_path = if target.starts_with('/') {
                    target[1..].to_string()
                } else {
                    format!("ppt/slides/{}", target)
                };
                let img_path = normalize_zip_path(&img_path);
                if let Ok(data) = read_zip_entry_bytes(archive, &img_path) {
                    let mime = guess_mime(&img_path);
                    s.content = ShapeContent::ImageData {
                        data,
                        mime_type: mime.to_string(),
                    };
                }
            }
        }
    }

    // Resolve fill image (blipFill on shape)
    if let Some(r_id) = s.fill_image_r_id.as_ref() {
        if let Some(ref rels_xml) = rels {
            if let Some(target) = resolve_relationship(rels_xml, r_id) {
                let img_path = if target.starts_with('/') {
                    target[1..].to_string()
                } else {
                    format!("ppt/slides/{}", target)
                };
                let img_path = normalize_zip_path(&img_path);
                if let Ok(data) = read_zip_entry_bytes(archive, &img_path) {
                    let mime = guess_mime(&img_path);
                    s.fill = Some(ShapeFill::Image {
                        data,
                        mime_type: mime.to_string(),
                    });
                    s.fill_image_r_id = None;
                }
            }
        }
    }

    s
}

/// ファイルパスからMIMEタイプを推測
fn guess_mime(path: &str) -> &'static str {
    let lower = path.to_lowercase();
    if lower.ends_with(".png") {
        "image/png"
    } else if lower.ends_with(".jpeg") || lower.ends_with(".jpg") {
        "image/jpeg"
    } else if lower.ends_with(".gif") {
        "image/gif"
    } else if lower.ends_with(".svg") {
        "image/svg+xml"
    } else if lower.ends_with(".emf") {
        "image/emf"
    } else if lower.ends_with(".wmf") {
        "image/wmf"
    } else {
        "image/png"
    }
}

/// スライド背景を完全解析（画像・グラデーション・ソリッド）
fn parse_slide_background_full(
    xml: &str,
    rels: &Option<String>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    theme_colors: &ThemeColors,
) -> Option<SlideBg> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut in_bg = false;
    let mut in_bg_pr = false;
    let mut in_solid_fill = false;
    let mut in_grad_fill = false;
    let mut in_blip_fill = false;
    let mut grad_stops: Vec<GradientStop> = Vec::new();
    let mut grad_angle: f64 = 0.0;
    let mut cur_grad_pos: f64 = 0.0;
    let mut blip_r_id = String::new();
    let mut in_gs = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"bg" => in_bg = true,
                    b"bgPr" if in_bg => in_bg_pr = true,
                    b"solidFill" if in_bg_pr => in_solid_fill = true,
                    b"gradFill" if in_bg_pr => {
                        in_grad_fill = true;
                        grad_stops.clear();
                    }
                    b"blipFill" if in_bg_pr => in_blip_fill = true,
                    b"blip" if in_blip_fill => {
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key.ends_with("embed") || key == "embed" {
                                blip_r_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    b"gs" if in_grad_fill => {
                        in_gs = true;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"pos" {
                                cur_grad_pos = String::from_utf8_lossy(&attr.value)
                                    .parse::<f64>()
                                    .unwrap_or(0.0)
                                    / 100000.0;
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                if (in_solid_fill || in_gs) && (local.as_ref() == b"srgbClr" || local.as_ref() == b"schemeClr") {
                    let color = parse_color_element_themed(e, theme_colors);
                    if let Some(c) = color {
                        if in_gs {
                            grad_stops.push(GradientStop {
                                position: cur_grad_pos,
                                color: c,
                            });
                        } else if in_solid_fill {
                            return Some(SlideBg::Solid(c));
                        }
                    }
                }
                // blip in empty form
                if local.as_ref() == b"blip" && in_blip_fill {
                    for attr in e.attributes().flatten() {
                        let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                        if key.ends_with("embed") || key == "embed" {
                            blip_r_id = String::from_utf8_lossy(&attr.value).to_string();
                        }
                    }
                }
                // Linear gradient angle
                if local.as_ref() == b"lin" && in_grad_fill {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"ang" {
                            let ang_60k = String::from_utf8_lossy(&attr.value)
                                .parse::<f64>()
                                .unwrap_or(0.0);
                            grad_angle = ang_60k / 60000.0 * std::f64::consts::PI / 180.0;
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"bg" => {
                        // End of background - return what we found
                        if !blip_r_id.is_empty() {
                            // Resolve background image
                            if let Some(ref rels_xml) = rels {
                                if let Some(target) = resolve_relationship(rels_xml, &blip_r_id) {
                                    let img_path = if target.starts_with('/') {
                                        target[1..].to_string()
                                    } else {
                                        format!("ppt/slides/{}", target)
                                    };
                                    let img_path = normalize_zip_path(&img_path);
                                    if let Ok(data) = read_zip_entry_bytes(archive, &img_path) {
                                        let mime = guess_mime(&img_path);
                                        return Some(SlideBg::Image {
                                            data,
                                            mime_type: mime.to_string(),
                                        });
                                    }
                                }
                            }
                        }
                        if !grad_stops.is_empty() {
                            return Some(SlideBg::Gradient {
                                stops: grad_stops,
                                angle: grad_angle,
                            });
                        }
                        in_bg = false;
                    }
                    b"bgPr" => in_bg_pr = false,
                    b"solidFill" => in_solid_fill = false,
                    b"gradFill" if in_bg_pr => {
                        if !grad_stops.is_empty() {
                            return Some(SlideBg::Gradient {
                                stops: grad_stops,
                                angle: grad_angle,
                            });
                        }
                        in_grad_fill = false;
                    }
                    b"blipFill" => in_blip_fill = false,
                    b"gs" => in_gs = false,
                    _ => {}
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

/// スライドXMLからシェイプを完全解析（グループシェイプ含む）
fn parse_slide_shapes(xml: &str, theme_colors: &ThemeColors) -> Vec<SlideShape> {
    let mut shapes = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();

    // State machine for parsing nested XML
    let mut depth = 0u32;
    let mut in_sp = false;       // <p:sp>
    let mut in_pic = false;      // <p:pic>
    let mut in_cxn = false;      // <p:cxnSp>
    let mut in_grp = false;      // <p:grpSp>
    let mut grp_depth = 0u32;
    let mut shape_depth = 0u32;

    // Current shape state
    let mut cur_x: f64 = 0.0;
    let mut cur_y: f64 = 0.0;
    let mut cur_w: f64 = 0.0;
    let mut cur_h: f64 = 0.0;
    let mut cur_fill: Option<ShapeFill> = None;
    let mut cur_outline: Option<(Color, f64)> = None;
    let mut cur_rotation: f64 = 0.0;
    let mut cur_shadow: Option<ShadowEffect> = None;
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
    let mut in_grad_fill = false;
    let mut _grad_fill_ctx = 0u8; // 0=shape (reserved for future per-context gradients)
    let mut grad_stops: Vec<GradientStop> = Vec::new();
    let mut grad_angle: f64 = 0.0;
    let mut cur_grad_pos: f64 = 0.0;
    let mut in_gs = false;
    let mut in_effect_lst = false;
    let mut in_outer_shdw = false;
    let mut shdw_color: Option<Color> = None;
    let mut shdw_blur: f64 = 0.0;
    let mut shdw_dist: f64 = 0.0;
    let mut shdw_dir: f64 = 0.0;

    // Group shape offset for coordinate transform
    let mut grp_off_x: f64 = 0.0;
    let mut grp_off_y: f64 = 0.0;

    // 3D effects and geometry
    let mut cur_has_3d = false;
    let mut cur_preset_geom: Option<String> = None;

    // Custom geometry state
    let mut in_cust_geom = false;
    let mut in_path_lst = false;
    let mut cust_path_cmds: Vec<PathCommand> = Vec::new();
    let mut cust_path_w: f64 = 21600.0;
    let mut cust_path_h: f64 = 21600.0;
    let mut cust_geom_pts: Vec<(f64, f64)> = Vec::new();

    // blipFill on shape (image texture fill)
    let mut in_sp_blip_fill = false;
    let mut cur_fill_blip_r_id = String::new();

    // Style references (p:style > a:fillRef / a:lnRef)
    let mut in_style = false;
    let mut in_fill_ref = false;
    let mut in_ln_ref = false;
    let mut style_fill_color: Option<Color> = None;
    let mut style_ln_color: Option<Color> = None;

    // Text body margins (from bodyPr)
    let mut text_margin_left: f64 = 4.0;   // default 4pt
    let mut text_margin_top: f64 = 4.0;
    let mut text_margin_right: f64 = 4.0;
    let mut text_margin_bottom: f64 = 4.0;

    macro_rules! reset_shape_state {
        () => {
            cur_x = 0.0;
            cur_y = 0.0;
            cur_w = 0.0;
            cur_h = 0.0;
            cur_fill = None;
            cur_outline = None;
            cur_rotation = 0.0;
            cur_shadow = None;
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
            in_grad_fill = false;
            _grad_fill_ctx = 0;
            grad_stops = Vec::new();
            grad_angle = 0.0;
            cur_grad_pos = 0.0;
            in_gs = false;
            in_effect_lst = false;
            in_outer_shdw = false;
            shdw_color = None;
            shdw_blur = 0.0;
            shdw_dist = 0.0;
            shdw_dir = 0.0;
            cur_has_3d = false;
            cur_preset_geom = None;
            in_cust_geom = false;
            in_path_lst = false;
            cust_path_cmds.clear();
            cust_path_w = 21600.0;
            cust_path_h = 21600.0;
            cust_geom_pts.clear();
            in_sp_blip_fill = false;
            cur_fill_blip_r_id = String::new();
            in_style = false;
            in_fill_ref = false;
            in_ln_ref = false;
            style_fill_color = None;
            style_ln_color = None;
            text_margin_left = 4.0;
            text_margin_top = 4.0;
            text_margin_right = 4.0;
            text_margin_bottom = 4.0;
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
                    b"grpSp" if !in_sp && !in_pic && !in_cxn && !in_grp => {
                        in_grp = true;
                        grp_depth = depth;
                        grp_off_x = 0.0;
                        grp_off_y = 0.0;
                    }
                    b"spPr" if in_sp || in_pic || in_cxn => {
                        in_sp_pr = true;
                    }
                    b"grpSpPr" if in_grp && !in_sp && !in_pic && !in_cxn => {
                        // Group shape properties - get offset
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
                        } else if in_outer_shdw {
                            solid_fill_ctx = 3; // shadow color
                        } else if in_sp_pr {
                            solid_fill_ctx = 0; // shape fill
                        }
                    }
                    b"gradFill" if in_sp_pr && !in_ln => {
                        in_grad_fill = true;
                        _grad_fill_ctx = 0;
                        grad_stops.clear();
                        grad_angle = 0.0;
                    }
                    b"gs" if in_grad_fill => {
                        in_gs = true;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"pos" {
                                cur_grad_pos = String::from_utf8_lossy(&attr.value)
                                    .parse::<f64>()
                                    .unwrap_or(0.0)
                                    / 100000.0;
                            }
                        }
                    }
                    b"effectLst" if in_sp_pr || (in_sp || in_pic) => {
                        in_effect_lst = true;
                    }
                    b"outerShdw" if in_effect_lst => {
                        in_outer_shdw = true;
                        shdw_color = None;
                        shdw_blur = 0.0;
                        shdw_dist = 0.0;
                        shdw_dir = 0.0;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"blurRad" => {
                                    shdw_blur = String::from_utf8_lossy(&attr.value)
                                        .parse::<f64>()
                                        .unwrap_or(0.0)
                                        / EMU_PER_PT;
                                }
                                b"dist" => {
                                    shdw_dist = String::from_utf8_lossy(&attr.value)
                                        .parse::<f64>()
                                        .unwrap_or(0.0)
                                        / EMU_PER_PT;
                                }
                                b"dir" => {
                                    shdw_dir = String::from_utf8_lossy(&attr.value)
                                        .parse::<f64>()
                                        .unwrap_or(0.0)
                                        / 60000.0
                                        * std::f64::consts::PI
                                        / 180.0;
                                }
                                _ => {}
                            }
                        }
                    }
                    // 3D effects detection
                    b"scene3d" | b"sp3d" if in_sp_pr || in_sp || in_pic => {
                        cur_has_3d = true;
                    }
                    // Preset geometry
                    b"prstGeom" if in_sp_pr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"prst" {
                                cur_preset_geom = Some(
                                    String::from_utf8_lossy(&attr.value).to_string()
                                );
                            }
                        }
                    }
                    // Custom geometry
                    b"custGeom" if in_sp_pr => {
                        in_cust_geom = true;
                        cust_path_cmds.clear();
                    }
                    b"path" if in_cust_geom => {
                        in_path_lst = true;
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            match attr.key.local_name().as_ref() {
                                b"w" => {
                                    if let Ok(v) = String::from_utf8_lossy(&attr.value).parse::<f64>() {
                                        cust_path_w = v;
                                    }
                                }
                                b"h" => {
                                    if let Ok(v) = String::from_utf8_lossy(&attr.value).parse::<f64>() {
                                        cust_path_h = v;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    b"moveTo" if in_path_lst => {
                        cust_geom_pts.clear();
                    }
                    b"lnTo" if in_path_lst => {
                        cust_geom_pts.clear();
                    }
                    b"cubicBezTo" if in_path_lst => {
                        cust_geom_pts.clear();
                    }
                    b"quadBezTo" if in_path_lst => {
                        cust_geom_pts.clear();
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
                    b"blipFill" if in_sp && in_sp_pr => {
                        // Image texture fill on shape
                        in_sp_blip_fill = true;
                    }
                    b"blip" if in_pic => {
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key.ends_with("embed") || key == "embed" {
                                cur_r_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    b"blip" if in_sp_blip_fill => {
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key.ends_with("embed") || key == "embed" {
                                cur_fill_blip_r_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    // Style references (p:style)
                    b"style" if (in_sp || in_pic) && !in_sp_pr => {
                        in_style = true;
                    }
                    b"fillRef" if in_style => {
                        in_fill_ref = true;
                    }
                    b"lnRef" if in_style => {
                        in_ln_ref = true;
                    }
                    b"schemeClr" if in_fill_ref => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                style_fill_color = resolve_scheme_color(&val, theme_colors);
                            }
                        }
                    }
                    b"schemeClr" if in_ln_ref => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                style_ln_color = resolve_scheme_color(&val, theme_colors);
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
                    let color = parse_color_element_themed(e, theme_colors);
                    if let Some(c) = color {
                        match solid_fill_ctx {
                            0 => cur_fill = Some(ShapeFill::Solid(c)),
                            1 => {
                                cur_outline = Some((c, cur_outline.map_or(1.0, |o| o.1)));
                            }
                            2 => cur_color = Some(c),
                            3 => shdw_color = Some(c), // shadow
                            _ => {}
                        }
                    }
                }

                // Gradient stop colors
                if in_gs {
                    let color = parse_color_element_themed(e, theme_colors);
                    if let Some(c) = color {
                        grad_stops.push(GradientStop {
                            position: cur_grad_pos,
                            color: c,
                        });
                    }
                }

                // Shadow colors (in outerShdw directly)
                if in_outer_shdw && !in_solid_fill {
                    let color = parse_color_element_themed(e, theme_colors);
                    if let Some(c) = color {
                        shdw_color = Some(c);
                    }
                }

                // Linear gradient angle
                if local == b"lin" && in_grad_fill {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"ang" {
                            let ang_60k = String::from_utf8_lossy(&attr.value)
                                .parse::<f64>()
                                .unwrap_or(0.0);
                            grad_angle = ang_60k / 60000.0 * std::f64::consts::PI / 180.0;
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

                // Text body properties - parse margins
                if local == b"bodyPr" && (in_sp || in_pic) && !in_sp_pr {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"lIns" => {
                                text_margin_left = String::from_utf8_lossy(&attr.value)
                                    .parse::<f64>()
                                    .unwrap_or(91440.0)
                                    / EMU_PER_PT;
                            }
                            b"tIns" => {
                                text_margin_top = String::from_utf8_lossy(&attr.value)
                                    .parse::<f64>()
                                    .unwrap_or(45720.0)
                                    / EMU_PER_PT;
                            }
                            b"rIns" => {
                                text_margin_right = String::from_utf8_lossy(&attr.value)
                                    .parse::<f64>()
                                    .unwrap_or(91440.0)
                                    / EMU_PER_PT;
                            }
                            b"bIns" => {
                                text_margin_bottom = String::from_utf8_lossy(&attr.value)
                                    .parse::<f64>()
                                    .unwrap_or(45720.0)
                                    / EMU_PER_PT;
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

                // Shape fill blip (empty variant)
                if local == b"blip" && in_sp_blip_fill {
                    for attr in e.attributes().flatten() {
                        let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                        if key.ends_with("embed") || key == "embed" {
                            cur_fill_blip_r_id = String::from_utf8_lossy(&attr.value).to_string();
                        }
                    }
                }

                // Style schemeClr (empty variant)
                if local == b"schemeClr" && (in_fill_ref || in_ln_ref) {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"val" {
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if in_fill_ref {
                                style_fill_color = resolve_scheme_color(&val, theme_colors);
                            } else if in_ln_ref {
                                style_ln_color = resolve_scheme_color(&val, theme_colors);
                            }
                        }
                    }
                }

                // 3D effects (empty variants)
                if (local == b"scene3d" || local == b"sp3d") && (in_sp_pr || in_sp || in_pic) {
                    cur_has_3d = true;
                }

                // Preset geometry (empty variant)
                if local == b"prstGeom" && in_sp_pr {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"prst" {
                            cur_preset_geom = Some(
                                String::from_utf8_lossy(&attr.value).to_string()
                            );
                        }
                    }
                }

                // Custom geometry: collect points
                if local == b"pt" && in_path_lst {
                    let mut px = 0.0f64;
                    let mut py = 0.0f64;
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        match attr.key.local_name().as_ref() {
                            b"x" => {
                                if let Ok(v) = String::from_utf8_lossy(&attr.value).parse::<f64>() {
                                    px = v;
                                }
                            }
                            b"y" => {
                                if let Ok(v) = String::from_utf8_lossy(&attr.value).parse::<f64>() {
                                    py = v;
                                }
                            }
                            _ => {}
                        }
                    }
                    cust_geom_pts.push((px, py));
                }

                // Custom geometry: <a:close/> is an empty element
                if local == b"close" && in_path_lst {
                    cust_path_cmds.push(PathCommand::Close);
                }
            }

            Ok(quick_xml::events::Event::End(ref e)) => {
                let name = e.local_name();
                let local = name.as_ref();

                match local {
                    b"sp" if in_sp && depth == shape_depth => {
                        // Build shadow from collected shadow data
                        if in_outer_shdw || shdw_color.is_some() {
                            let offset_x = shdw_dist * shdw_dir.cos();
                            let offset_y = shdw_dist * shdw_dir.sin();
                            cur_shadow = Some(ShadowEffect {
                                color: shdw_color.unwrap_or(Color::rgb(0, 0, 0)),
                                blur_radius: shdw_blur,
                                offset_x,
                                offset_y,
                            });
                        }
                        // Finalize gradient fill if pending
                        if cur_fill.is_none() && !grad_stops.is_empty() {
                            cur_fill = Some(ShapeFill::Gradient {
                                stops: grad_stops.clone(),
                                angle: grad_angle,
                            });
                        }
                        // Apply style-based fill/outline as fallback
                        if cur_fill.is_none() && cur_fill_blip_r_id.is_empty() {
                            if let Some(c) = style_fill_color {
                                cur_fill = Some(ShapeFill::Solid(c));
                            }
                        }
                        if cur_outline.is_none() {
                            if let Some(c) = style_ln_color {
                                cur_outline = Some((c, 1.0));
                            }
                        }
                        // Emit shape
                        let content = if cur_paragraphs.is_empty() && cur_runs.is_empty() {
                            ShapeContent::Empty
                        } else {
                            ShapeContent::TextBox {
                                paragraphs: cur_paragraphs.clone(),
                            }
                        };
                        shapes.push(SlideShape {
                            x: cur_x + grp_off_x,
                            y: cur_y + grp_off_y,
                            width: cur_w,
                            height: cur_h,
                            content,
                            fill: cur_fill.clone(),
                            outline: cur_outline,
                            rotation: cur_rotation,
                            shadow: cur_shadow.clone(),
                            has_3d: cur_has_3d,
                            preset_geometry: cur_preset_geom.clone(),
                            custom_path: if cust_path_cmds.is_empty() { None } else { Some(cust_path_cmds.clone()) },
                            custom_path_viewport: if cust_path_cmds.is_empty() { None } else { Some((cust_path_w, cust_path_h)) },
                            fill_image_r_id: if cur_fill_blip_r_id.is_empty() { None } else { Some(cur_fill_blip_r_id.clone()) },
                            text_margin_left,
                            text_margin_top,
                            text_margin_right,
                            text_margin_bottom,
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
                            x: cur_x + grp_off_x,
                            y: cur_y + grp_off_y,
                            width: cur_w,
                            height: cur_h,
                            content,
                            fill: cur_fill.clone(),
                            outline: cur_outline,
                            rotation: cur_rotation,
                            shadow: cur_shadow.clone(),
                            has_3d: cur_has_3d,
                            preset_geometry: cur_preset_geom.clone(),
                            custom_path: None,
                            custom_path_viewport: None,
                            fill_image_r_id: None,
                            text_margin_left,
                            text_margin_top,
                            text_margin_right,
                            text_margin_bottom,
                        });
                        in_pic = false;
                    }
                    b"cxnSp" if in_cxn && depth == shape_depth => {
                        shapes.push(SlideShape {
                            x: cur_x + grp_off_x,
                            y: cur_y + grp_off_y,
                            width: cur_w,
                            height: cur_h,
                            content: ShapeContent::Connector,
                            fill: None,
                            outline: cur_outline.or(Some((Color::BLACK, 1.0))),
                            rotation: cur_rotation,
                            shadow: None,
                            has_3d: false,
                            preset_geometry: None,
                            custom_path: None,
                            custom_path_viewport: None,
                            fill_image_r_id: None,
                            text_margin_left: 4.0,
                            text_margin_top: 4.0,
                            text_margin_right: 4.0,
                            text_margin_bottom: 4.0,
                        });
                        in_cxn = false;
                    }
                    b"grpSp" if in_grp && depth == grp_depth => {
                        in_grp = false;
                        grp_off_x = 0.0;
                        grp_off_y = 0.0;
                    }
                    b"spPr" => {
                        in_sp_pr = false;
                    }
                    // Custom geometry path command end elements
                    b"moveTo" if in_path_lst => {
                        if let Some(&(px, py)) = cust_geom_pts.first() {
                            cust_path_cmds.push(PathCommand::MoveTo(px, py));
                        }
                    }
                    b"lnTo" if in_path_lst => {
                        if let Some(&(px, py)) = cust_geom_pts.first() {
                            cust_path_cmds.push(PathCommand::LineTo(px, py));
                        }
                    }
                    b"cubicBezTo" if in_path_lst => {
                        if cust_geom_pts.len() >= 3 {
                            cust_path_cmds.push(PathCommand::CubicTo(
                                cust_geom_pts[0].0, cust_geom_pts[0].1,
                                cust_geom_pts[1].0, cust_geom_pts[1].1,
                                cust_geom_pts[2].0, cust_geom_pts[2].1,
                            ));
                        }
                    }
                    b"quadBezTo" if in_path_lst => {
                        if cust_geom_pts.len() >= 2 {
                            cust_path_cmds.push(PathCommand::QuadTo(
                                cust_geom_pts[0].0, cust_geom_pts[0].1,
                                cust_geom_pts[1].0, cust_geom_pts[1].1,
                            ));
                        }
                    }
                    b"path" if in_cust_geom => {
                        in_path_lst = false;
                    }
                    b"custGeom" => {
                        in_cust_geom = false;
                        in_path_lst = false;
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
                    b"gradFill" => {
                        if in_sp_pr && !in_ln && !grad_stops.is_empty() {
                            cur_fill = Some(ShapeFill::Gradient {
                                stops: grad_stops.clone(),
                                angle: grad_angle,
                            });
                        }
                        in_grad_fill = false;
                    }
                    b"gs" => {
                        in_gs = false;
                    }
                    b"effectLst" => {
                        in_effect_lst = false;
                    }
                    b"outerShdw" => {
                        if shdw_color.is_some() {
                            let offset_x = shdw_dist * shdw_dir.cos();
                            let offset_y = shdw_dist * shdw_dir.sin();
                            cur_shadow = Some(ShadowEffect {
                                color: shdw_color.unwrap_or(Color::rgb(0, 0, 0)),
                                blur_radius: shdw_blur,
                                offset_x,
                                offset_y,
                            });
                        }
                        in_outer_shdw = false;
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
                    b"blipFill" if in_sp_blip_fill => {
                        in_sp_blip_fill = false;
                    }
                    b"style" if in_style => {
                        in_style = false;
                    }
                    b"fillRef" if in_fill_ref => {
                        in_fill_ref = false;
                    }
                    b"lnRef" if in_ln_ref => {
                        in_ln_ref = false;
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

/// XML要素から色を解析（テーマカラー対応版）
fn parse_color_element_themed(e: &quick_xml::events::BytesStart, theme: &ThemeColors) -> Option<Color> {
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
            for attr in e.attributes().flatten() {
                if attr.key.as_ref() == b"val" {
                    let scheme = String::from_utf8_lossy(&attr.value).to_string();
                    return Some(theme.resolve(&scheme));
                }
            }
            None
        }
        b"sysClr" => {
            // System color - use lastClr attribute
            for attr in e.attributes().flatten() {
                if attr.key.as_ref() == b"lastClr" {
                    return parse_hex_color(&String::from_utf8_lossy(&attr.value));
                }
            }
            // Fallback to val
            for attr in e.attributes().flatten() {
                if attr.key.as_ref() == b"val" {
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    return Some(match val.as_str() {
                        "windowText" => Color::rgb(0, 0, 0),
                        "window" => Color::rgb(255, 255, 255),
                        _ => Color::rgb(0, 0, 0),
                    });
                }
            }
            None
        }
        b"prstClr" => {
            for attr in e.attributes().flatten() {
                if attr.key.as_ref() == b"val" {
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    return Some(match val.as_str() {
                        "black" => Color::rgb(0, 0, 0),
                        "white" => Color::rgb(255, 255, 255),
                        "red" => Color::rgb(255, 0, 0),
                        "green" => Color::rgb(0, 128, 0),
                        "blue" => Color::rgb(0, 0, 255),
                        "yellow" => Color::rgb(255, 255, 0),
                        _ => Color::rgb(0, 0, 0),
                    });
                }
            }
            None
        }
        _ => None,
    }
}

/// XML要素から色を解析（デフォルトテーマ用互換関数）
#[allow(dead_code)]
fn parse_color_element(e: &quick_xml::events::BytesStart) -> Option<Color> {
    parse_color_element_themed(e, &ThemeColors::default())
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

/// Resolve a schemeClr value to a Color via theme
fn resolve_scheme_color(val: &str, theme: &ThemeColors) -> Option<Color> {
    const VALID: &[&str] = &[
        "tx1", "dk1", "bg1", "lt1", "dk2", "tx2", "lt2", "bg2",
        "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
        "hlink", "folHlink",
    ];
    if VALID.contains(&val) {
        Some(theme.resolve(val))
    } else {
        None
    }
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

// ── Chart and SmartArt detection ──

/// スライドXMLからチャート参照を検出し、チャートを描画
fn detect_and_render_charts(
    slide_xml: &str,
    rels: &Option<String>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    slide_size: &SlideSize,
) -> Vec<PageElement> {
    let rels_xml = match rels {
        Some(r) => r,
        None => return Vec::new(),
    };

    let mut elements = Vec::new();

    // チャートフレームの位置とrIdを検出
    let chart_refs = find_chart_frames(slide_xml);

    for (r_id, cx, cy, cw, ch) in chart_refs {
        // リレーションシップからチャートXMLパスを解決
        if let Some(target) = resolve_relationship(rels_xml, &r_id) {
            let chart_path = if target.starts_with("../") {
                format!("ppt/{}", target.trim_start_matches("../"))
            } else if target.starts_with('/') {
                target.trim_start_matches('/').to_string()
            } else {
                format!("ppt/slides/{}", target)
            };

            if let Ok(chart_xml) = read_zip_entry_string(archive, &chart_path) {
                let chart_x = if cx > 0.0 { cx } else { slide_size.width * 0.1 };
                let chart_y = if cy > 0.0 { cy } else { slide_size.height * 0.15 };
                let chart_w = if cw > 0.0 { cw } else { slide_size.width * 0.8 };
                let chart_h = if ch > 0.0 { ch } else { slide_size.height * 0.7 };

                let chart_elems = crate::formats::chart::render_chart(
                    &chart_xml,
                    chart_x,
                    chart_y,
                    chart_w,
                    chart_h,
                );
                elements.extend(chart_elems);
            }
        }
    }

    elements
}

/// スライドXMLからSmartArt/ダイアグラム参照を検出し、描画
fn detect_and_render_smartart(
    slide_xml: &str,
    rels: &Option<String>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    slide_size: &SlideSize,
) -> Vec<PageElement> {
    let rels_xml = match rels {
        Some(r) => r,
        None => return Vec::new(),
    };

    let mut elements = Vec::new();

    // ダイアグラム描画のリレーションシップを検出
    let dgm_refs = find_diagram_frames(slide_xml, rels_xml);

    for (drawing_path, dx, dy, dw, dh) in dgm_refs {
        let full_path = if drawing_path.starts_with("../") {
            format!("ppt/{}", drawing_path.trim_start_matches("../"))
        } else if drawing_path.starts_with('/') {
            drawing_path.trim_start_matches('/').to_string()
        } else {
            drawing_path
        };

        if let Ok(drawing_xml) = read_zip_entry_string(archive, &full_path) {
            let sa_x = if dx > 0.0 { dx } else { slide_size.width * 0.1 };
            let sa_y = if dy > 0.0 { dy } else { slide_size.height * 0.15 };
            let sa_w = if dw > 0.0 { dw } else { slide_size.width * 0.8 };
            let sa_h = if dh > 0.0 { dh } else { slide_size.height * 0.7 };

            let sa_elems = crate::formats::smartart::render_smartart(
                &drawing_xml,
                sa_x,
                sa_y,
                sa_w,
                sa_h,
            );
            elements.extend(sa_elems);
        }
    }

    elements
}

/// スライドXMLからチャートフレームの位置とrIdを検出
/// Returns: Vec<(rId, x, y, width, height)>
fn find_chart_frames(xml: &str) -> Vec<(String, f64, f64, f64, f64)> {
    let mut results = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();

    let mut in_graphic_frame = false;
    let mut frame_x = 0.0f64;
    let mut frame_y = 0.0f64;
    let mut frame_w = 0.0f64;
    let mut frame_h = 0.0f64;
    let mut in_xfrm = false;
    let mut chart_r_id: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();

                match name {
                    b"graphicFrame" => {
                        in_graphic_frame = true;
                        frame_x = 0.0;
                        frame_y = 0.0;
                        frame_w = 0.0;
                        frame_h = 0.0;
                        chart_r_id = None;
                    }
                    b"xfrm" if in_graphic_frame => { in_xfrm = true; }
                    b"chart" if in_graphic_frame => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "id" || key.ends_with(":id") {
                                chart_r_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();

                match name {
                    b"off" if in_xfrm && in_graphic_frame => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val_str = std::str::from_utf8(&attr.value).unwrap_or("0");
                            let val = val_str.parse::<f64>().unwrap_or(0.0) / EMU_PER_PT;
                            match key {
                                "x" => frame_x = val,
                                "y" => frame_y = val,
                                _ => {}
                            }
                        }
                    }
                    b"ext" if in_xfrm && in_graphic_frame => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val_str = std::str::from_utf8(&attr.value).unwrap_or("0");
                            let val = val_str.parse::<f64>().unwrap_or(0.0) / EMU_PER_PT;
                            match key {
                                "cx" => frame_w = val,
                                "cy" => frame_h = val,
                                _ => {}
                            }
                        }
                    }
                    b"chart" if in_graphic_frame => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "id" || key.ends_with(":id") {
                                chart_r_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"graphicFrame" => {
                        if let Some(r_id) = chart_r_id.take() {
                            results.push((r_id, frame_x, frame_y, frame_w, frame_h));
                        }
                        in_graphic_frame = false;
                    }
                    b"xfrm" => { in_xfrm = false; }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    results
}

/// スライドXMLとリレーションシップからダイアグラム描画パスと位置を検出
/// Returns: Vec<(drawing_path, x, y, width, height)>
fn find_diagram_frames(xml: &str, rels_xml: &str) -> Vec<(String, f64, f64, f64, f64)> {
    let mut results = Vec::new();

    // ダイアグラム関連のリレーションシップを検出
    let mut dgm_drawing_targets = Vec::new();
    {
        let mut reader = quick_xml::Reader::from_str(rels_xml);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Empty(ref e))
                | Ok(quick_xml::events::Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"Relationship" {
                        let mut rel_type = String::new();
                        let mut target = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"Type" => rel_type = String::from_utf8_lossy(&attr.value).to_string(),
                                b"Target" => target = String::from_utf8_lossy(&attr.value).to_string(),
                                _ => {}
                            }
                        }
                        // dgm名前空間のdrawingタイプ
                        if rel_type.contains("diagramDrawing") || rel_type.contains("dgmDrawing") {
                            dgm_drawing_targets.push(target);
                        }
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    if dgm_drawing_targets.is_empty() {
        return results;
    }

    // graphicFrame内のダイアグラムの位置を取得
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut in_graphic_frame = false;
    let mut frame_x = 0.0f64;
    let mut frame_y = 0.0f64;
    let mut frame_w = 0.0f64;
    let mut frame_h = 0.0f64;
    let mut in_xfrm = false;
    let mut has_dgm = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();
                match name {
                    b"graphicFrame" => {
                        in_graphic_frame = true;
                        frame_x = 0.0;
                        frame_y = 0.0;
                        frame_w = 0.0;
                        frame_h = 0.0;
                        has_dgm = false;
                    }
                    b"xfrm" if in_graphic_frame => { in_xfrm = true; }
                    _ => {
                        if in_graphic_frame {
                            // Check for dgm namespace presence
                            let qname = e.name();
                            let name_bytes = qname.as_ref();
                            let full_name = std::str::from_utf8(name_bytes).unwrap_or("");
                            if full_name.contains("dgm") || full_name.contains("diagram") {
                                has_dgm = true;
                            }
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();
                match name {
                    b"off" if in_xfrm && in_graphic_frame => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val_str = std::str::from_utf8(&attr.value).unwrap_or("0");
                            let val = val_str.parse::<f64>().unwrap_or(0.0) / EMU_PER_PT;
                            match key {
                                "x" => frame_x = val,
                                "y" => frame_y = val,
                                _ => {}
                            }
                        }
                    }
                    b"ext" if in_xfrm && in_graphic_frame => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val_str = std::str::from_utf8(&attr.value).unwrap_or("0");
                            let val = val_str.parse::<f64>().unwrap_or(0.0) / EMU_PER_PT;
                            match key {
                                "cx" => frame_w = val,
                                "cy" => frame_h = val,
                                _ => {}
                            }
                        }
                    }
                    _ => {
                        if in_graphic_frame {
                            let qname = e.name();
                            let name_bytes = qname.as_ref();
                            let full_name = std::str::from_utf8(name_bytes).unwrap_or("");
                            if full_name.contains("dgm") || full_name.contains("diagram") {
                                has_dgm = true;
                            }
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"graphicFrame" => {
                        if in_graphic_frame && has_dgm {
                            // Use first available diagram drawing target
                            if let Some(target) = dgm_drawing_targets.first() {
                                results.push((target.clone(), frame_x, frame_y, frame_w, frame_h));
                            }
                        }
                        in_graphic_frame = false;
                    }
                    b"xfrm" => { in_xfrm = false; }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    results
}

// ── Page rendering ──

/// 解析済みシェイプからページを構築
fn render_slide_page(
    shapes: &[SlideShape],
    slide_size: &SlideSize,
    bg: Option<&SlideBg>,
) -> Page {
    let mut page = Page {
        width: slide_size.width,
        height: slide_size.height,
        elements: Vec::new(),
    };

    // 背景
    match bg {
        Some(SlideBg::Solid(color)) => {
            page.elements.push(PageElement::Rect {
                x: 0.0,
                y: 0.0,
                width: slide_size.width,
                height: slide_size.height,
                fill: Some(*color),
                stroke: None,
                stroke_width: 0.0,
            });
        }
        Some(SlideBg::Gradient { stops, angle }) => {
            page.elements.push(PageElement::GradientRect {
                x: 0.0,
                y: 0.0,
                width: slide_size.width,
                height: slide_size.height,
                stops: stops.clone(),
                gradient_type: GradientType::Linear(*angle),
            });
        }
        Some(SlideBg::Image { data, mime_type }) => {
            page.elements.push(PageElement::Image {
                x: 0.0,
                y: 0.0,
                width: slide_size.width,
                height: slide_size.height,
                data: data.clone(),
                mime_type: mime_type.clone(),
            });
        }
        None => {}
    }

    for shape in shapes {
        // Render shadow first (behind the shape)
        if let Some(ref shadow) = shape.shadow {
            let shadow_alpha = (shadow.color.a as f64 * 0.5) as u8;
            page.elements.push(PageElement::Rect {
                x: shape.x + shadow.offset_x,
                y: shape.y + shadow.offset_y,
                width: shape.width,
                height: shape.height,
                fill: Some(Color {
                    r: shadow.color.r,
                    g: shadow.color.g,
                    b: shadow.color.b,
                    a: shadow_alpha,
                }),
                stroke: None,
                stroke_width: 0.0,
            });
        }

        match &shape.content {
            ShapeContent::TextBox { paragraphs } => {
                // Check for ellipse/rounded geometry
                let is_ellipse = shape.preset_geometry.as_deref() == Some("ellipse");

                // 3D effect: draw depth extrusion behind the shape
                if shape.has_3d && shape.width > 0.0 && shape.height > 0.0 {
                    let depth = SHAPE_3D_EXTRUSION_DEPTH;
                    let base_color = match &shape.fill {
                        Some(ShapeFill::Solid(c)) => *c,
                        _ => Color::rgb(150, 150, 150),
                    };
                    // Darker version for 3D sides
                    let dark_color = Color::rgb(
                        (base_color.r as f64 * 0.6) as u8,
                        (base_color.g as f64 * 0.6) as u8,
                        (base_color.b as f64 * 0.6) as u8,
                    );
                    // Bottom depth strip
                    page.elements.push(PageElement::Rect {
                        x: shape.x + 2.0,
                        y: shape.y + shape.height,
                        width: shape.width,
                        height: depth,
                        fill: Some(dark_color),
                        stroke: None,
                        stroke_width: 0.0,
                    });
                    // Right depth strip
                    page.elements.push(PageElement::Rect {
                        x: shape.x + shape.width,
                        y: shape.y + 2.0,
                        width: depth,
                        height: shape.height,
                        fill: Some(dark_color),
                        stroke: None,
                        stroke_width: 0.0,
                    });
                }

                // Shape fill - use preset geometry for rendering
                let mut shape_rendered = false;

                // Try custom geometry path rendering first
                if let (Some(ref cmds), Some((vp_w, vp_h))) = (&shape.custom_path, shape.custom_path_viewport) {
                    if !cmds.is_empty() && vp_w > 0.0 && vp_h > 0.0 {
                        let scale_x = shape.width / vp_w;
                        let scale_y = shape.height / vp_h;
                        let scaled_cmds: Vec<PathCommand> = cmds.iter().map(|cmd| {
                            match cmd {
                                PathCommand::MoveTo(x, y) => PathCommand::MoveTo(shape.x + x * scale_x, shape.y + y * scale_y),
                                PathCommand::LineTo(x, y) => PathCommand::LineTo(shape.x + x * scale_x, shape.y + y * scale_y),
                                PathCommand::QuadTo(cx, cy, x, y) => PathCommand::QuadTo(
                                    shape.x + cx * scale_x, shape.y + cy * scale_y,
                                    shape.x + x * scale_x, shape.y + y * scale_y,
                                ),
                                PathCommand::CubicTo(cx1, cy1, cx2, cy2, x, y) => PathCommand::CubicTo(
                                    shape.x + cx1 * scale_x, shape.y + cy1 * scale_y,
                                    shape.x + cx2 * scale_x, shape.y + cy2 * scale_y,
                                    shape.x + x * scale_x, shape.y + y * scale_y,
                                ),
                                PathCommand::ArcTo(rx, ry, rot, large, sweep, x, y) => PathCommand::ArcTo(
                                    rx * scale_x, ry * scale_y, *rot, *large, *sweep,
                                    shape.x + x * scale_x, shape.y + y * scale_y,
                                ),
                                PathCommand::Close => PathCommand::Close,
                            }
                        }).collect();

                        // Handle image fill for custom geometries
                        if let Some(ShapeFill::Image { data, mime_type }) = &shape.fill {
                            let (stroke_color, stroke_w) = shape.outline.map_or((None, 0.0), |(c, w)| (Some(c), w));
                            page.elements.push(PageElement::PathImage {
                                commands: scaled_cmds,
                                data: data.clone(),
                                mime_type: mime_type.clone(),
                                stroke: stroke_color,
                                stroke_width: stroke_w,
                            });
                        } else {
                            let fill_color = match &shape.fill {
                                Some(ShapeFill::Solid(c)) => Some(*c),
                                _ => None,
                            };
                            let (stroke_color, stroke_w) = shape.outline.map_or((None, 0.0), |(c, w)| (Some(c), w));
                            page.elements.push(PageElement::Path {
                                commands: scaled_cmds,
                                fill: fill_color,
                                stroke: stroke_color,
                                stroke_width: stroke_w,
                            });
                        }
                        shape_rendered = true;
                    }
                }

                // Try path-based rendering for non-trivial geometries
                if let Some(ref geom_name) = shape.preset_geometry {
                    if geom_name != "rect" && geom_name != "ellipse" {
                        if let Some(path_cmds) = generate_preset_path(geom_name, shape.x, shape.y, shape.width, shape.height) {
                            // Handle image fill for preset geometries
                            if let Some(ShapeFill::Image { data, mime_type }) = &shape.fill {
                                let (stroke_color, stroke_w) = shape.outline.map_or((None, 0.0), |(c, w)| (Some(c), w));
                                page.elements.push(PageElement::PathImage {
                                    commands: path_cmds,
                                    data: data.clone(),
                                    mime_type: mime_type.clone(),
                                    stroke: stroke_color,
                                    stroke_width: stroke_w,
                                });
                            } else {
                                let fill_color = match &shape.fill {
                                    Some(ShapeFill::Solid(c)) => Some(*c),
                                    _ => None,
                                };
                                let (stroke_color, stroke_w) = shape.outline.map_or((None, 0.0), |(c, w)| (Some(c), w));
                                page.elements.push(PageElement::Path {
                                    commands: path_cmds,
                                    fill: fill_color,
                                    stroke: stroke_color,
                                    stroke_width: stroke_w,
                                });
                            }
                            shape_rendered = true;
                        }
                    }
                }
                if !shape_rendered {
                    if is_ellipse {
                        // Render ellipse with image fill using elliptical clipping
                        if let Some(ShapeFill::Image { data, mime_type }) = &shape.fill {
                            let stroke_info = shape.outline;
                            page.elements.push(PageElement::EllipseImage {
                                cx: shape.x + shape.width / 2.0,
                                cy: shape.y + shape.height / 2.0,
                                rx: shape.width / 2.0,
                                ry: shape.height / 2.0,
                                data: data.clone(),
                                mime_type: mime_type.clone(),
                                stroke: stroke_info.map(|(c, _)| c),
                                stroke_width: stroke_info.map_or(0.0, |(_, w)| w),
                            });
                            shape_rendered = true;
                        } else {
                            // Solid or gradient fill
                            let fill_color = match &shape.fill {
                                Some(ShapeFill::Solid(c)) => Some(*c),
                                _ => None,
                            };
                            let stroke_info = shape.outline;
                            page.elements.push(PageElement::Ellipse {
                                cx: shape.x + shape.width / 2.0,
                                cy: shape.y + shape.height / 2.0,
                                rx: shape.width / 2.0,
                                ry: shape.height / 2.0,
                                fill: fill_color,
                                stroke: stroke_info.map(|(c, _)| c),
                                stroke_width: stroke_info.map_or(0.0, |(_, w)| w),
                            });
                            shape_rendered = true;
                        }
                    } else {
                        // Standard rectangle or rounded rect rendering
                        match &shape.fill {
                            Some(ShapeFill::Solid(color)) => {
                                page.elements.push(PageElement::Rect {
                                    x: shape.x,
                                    y: shape.y,
                                    width: shape.width,
                                    height: shape.height,
                                    fill: Some(*color),
                                    stroke: None,
                                    stroke_width: 0.0,
                                });
                            }
                            Some(ShapeFill::Gradient { stops, angle }) => {
                                page.elements.push(PageElement::GradientRect {
                                    x: shape.x,
                                    y: shape.y,
                                    width: shape.width,
                                    height: shape.height,
                                    stops: stops.clone(),
                                    gradient_type: GradientType::Linear(*angle),
                                });
                            }
                            Some(ShapeFill::Image { data, mime_type }) => {
                                page.elements.push(PageElement::Image {
                                    x: shape.x,
                                    y: shape.y,
                                    width: shape.width,
                                    height: shape.height,
                                    data: data.clone(),
                                    mime_type: mime_type.clone(),
                                });
                            }
                            None => {}
                        }
                    }
                } // end shape fill

                // Shape outline (skip if already rendered as a path with stroke)
                if !shape_rendered {
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
                }

                // Render text paragraphs positioned within the shape
                let margin_left = shape.text_margin_left;
                let margin_top = shape.text_margin_top;
                let margin_right = shape.text_margin_right;
                let _margin_bottom = shape.text_margin_bottom;
                let mut text_y = shape.y + margin_top;

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
                        let available_width = shape.width - margin_left - margin_right - indent;
                        let lines = wrap_text(&full_text, available_width, font_size);

                        for line in &lines {
                            if text_y + font_size > shape.y + shape.height {
                                break; // Clip to shape bounds
                            }

                            page.elements.push(PageElement::Text {
                                x: shape.x + margin_left + indent,
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
                match &shape.fill {
                    Some(ShapeFill::Solid(color)) => {
                        page.elements.push(PageElement::Rect {
                            x: shape.x,
                            y: shape.y,
                            width: shape.width,
                            height: shape.height,
                            fill: Some(*color),
                            stroke: shape.outline.map(|(c, _)| c),
                            stroke_width: shape.outline.map(|(_, w)| w).unwrap_or(0.0),
                        });
                    }
                    Some(ShapeFill::Gradient { stops, angle }) => {
                        page.elements.push(PageElement::GradientRect {
                            x: shape.x,
                            y: shape.y,
                            width: shape.width,
                            height: shape.height,
                            stops: stops.clone(),
                            gradient_type: GradientType::Linear(*angle),
                        });
                    }
                    Some(ShapeFill::Image { data, mime_type }) => {
                        page.elements.push(PageElement::Image {
                            x: shape.x,
                            y: shape.y,
                            width: shape.width,
                            height: shape.height,
                            data: data.clone(),
                            mime_type: mime_type.clone(),
                        });
                    }
                    None => {
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
                    }
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

/// プリセットジオメトリ名からパスコマンドを生成
/// シェイプのバウンディングボックス (x, y, width, height) を基にパスを計算
///
/// DOCX, XLSX, PPTX で共通利用可能な86種類のプリセットジオメトリをサポート
pub fn generate_preset_path(name: &str, x: f64, y: f64, w: f64, h: f64) -> Option<Vec<crate::converter::PathCommand>> {
    use crate::converter::PathCommand;
    use std::f64::consts::PI;

    match name {
        "triangle" | "isosTriangle" => {
            Some(vec![
                PathCommand::MoveTo(x + w / 2.0, y),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
            ])
        }
        "rtTriangle" => {
            Some(vec![
                PathCommand::MoveTo(x, y),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
            ])
        }
        "diamond" => {
            Some(vec![
                PathCommand::MoveTo(x + w / 2.0, y),
                PathCommand::LineTo(x + w, y + h / 2.0),
                PathCommand::LineTo(x + w / 2.0, y + h),
                PathCommand::LineTo(x, y + h / 2.0),
                PathCommand::Close,
            ])
        }
        "parallelogram" => {
            let off = w * 0.25;
            Some(vec![
                PathCommand::MoveTo(x + off, y),
                PathCommand::LineTo(x + w, y),
                PathCommand::LineTo(x + w - off, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
            ])
        }
        "trapezoid" => {
            let off = w * 0.2;
            Some(vec![
                PathCommand::MoveTo(x + off, y),
                PathCommand::LineTo(x + w - off, y),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
            ])
        }
        "pentagon" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let mut cmds = Vec::new();
            for i in 0..5 {
                let angle = -PI / 2.0 + 2.0 * PI * i as f64 / 5.0;
                let px = cx + rx * angle.cos();
                let py = cy + ry * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        "hexagon" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let mut cmds = Vec::new();
            for i in 0..6 {
                let angle = 2.0 * PI * i as f64 / 6.0;
                let px = cx + rx * angle.cos();
                let py = cy + ry * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        "octagon" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let mut cmds = Vec::new();
            for i in 0..8 {
                let angle = -PI / 8.0 + 2.0 * PI * i as f64 / 8.0;
                let px = cx + rx * angle.cos();
                let py = cy + ry * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        "star4" | "star5" | "star6" | "star8" | "star10" | "star12" | "star16" | "star24" | "star32" => {
            let points: usize = match name {
                "star4" => 4,
                "star5" => 5,
                "star6" => 6,
                "star8" => 8,
                "star10" => 10,
                "star12" => 12,
                "star16" => 16,
                "star24" => 24,
                "star32" => 32,
                _ => 5,
            };
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let inner_ratio = 0.4;
            let mut cmds = Vec::new();
            for i in 0..(points * 2) {
                let angle = -PI / 2.0 + PI * i as f64 / points as f64;
                let (r_x, r_y) = if i % 2 == 0 { (rx, ry) } else { (rx * inner_ratio, ry * inner_ratio) };
                let px = cx + r_x * angle.cos();
                let py = cy + r_y * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        "arc" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let steps = 32;
            let mut cmds = Vec::new();
            for i in 0..=steps {
                let angle = PI + PI * i as f64 / steps as f64;
                let px = cx + rx * angle.cos();
                let py = cy + ry * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            Some(cmds)
        }
        "pie" | "pieWedge" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let start_angle = -PI / 4.0;
            let end_angle = PI / 4.0;
            let steps = 24;
            let mut cmds = vec![PathCommand::MoveTo(cx, cy)];
            for i in 0..=steps {
                let angle = start_angle + (end_angle - start_angle) * i as f64 / steps as f64;
                let px = cx + rx * angle.cos();
                let py = cy + ry * angle.sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        "donut" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let steps = 48;
            let mut cmds = Vec::new();
            for i in 0..=steps {
                let angle = 2.0 * PI * i as f64 / steps as f64;
                let px = cx + rx * angle.cos();
                let py = cy + ry * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        "blockArc" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let steps = 32;
            let mut cmds = Vec::new();
            for i in 0..=steps {
                let angle = PI * 0.75 + PI * 1.5 * i as f64 / steps as f64;
                let px = cx + rx * angle.cos();
                let py = cy + ry * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            let ir = 0.6;
            for i in (0..=steps).rev() {
                let angle = PI * 0.75 + PI * 1.5 * i as f64 / steps as f64;
                let px = cx + rx * ir * angle.cos();
                let py = cy + ry * ir * angle.sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        "heart" => {
            let cx = x + w / 2.0;
            Some(vec![
                PathCommand::MoveTo(cx, y + h),
                PathCommand::CubicTo(x, y + h * 0.6, x, y, cx - w * 0.02, y + h * 0.35),
                PathCommand::CubicTo(cx, y, cx, y, cx, y + h * 0.35),
                PathCommand::CubicTo(cx, y, cx, y, cx + w * 0.02, y + h * 0.35),
                PathCommand::CubicTo(x + w, y, x + w, y + h * 0.6, cx, y + h),
                PathCommand::Close,
            ])
        }
        "lightningBolt" => {
            Some(vec![
                PathCommand::MoveTo(x + w * 0.4, y),
                PathCommand::LineTo(x + w * 0.65, y + h * 0.35),
                PathCommand::LineTo(x + w * 0.5, y + h * 0.35),
                PathCommand::LineTo(x + w * 0.75, y + h * 0.65),
                PathCommand::LineTo(x + w * 0.55, y + h * 0.65),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x + w * 0.35, y + h * 0.55),
                PathCommand::LineTo(x + w * 0.5, y + h * 0.55),
                PathCommand::LineTo(x + w * 0.25, y + h * 0.25),
                PathCommand::LineTo(x + w * 0.4, y + h * 0.25),
                PathCommand::Close,
            ])
        }
        "rightArrow" | "arrow" => {
            let ah = h * 0.2;
            Some(vec![
                PathCommand::MoveTo(x, y + ah),
                PathCommand::LineTo(x + w * 0.6, y + ah),
                PathCommand::LineTo(x + w * 0.6, y),
                PathCommand::LineTo(x + w, y + h / 2.0),
                PathCommand::LineTo(x + w * 0.6, y + h),
                PathCommand::LineTo(x + w * 0.6, y + h - ah),
                PathCommand::LineTo(x, y + h - ah),
                PathCommand::Close,
            ])
        }
        "leftArrow" => {
            let ah = h * 0.2;
            Some(vec![
                PathCommand::MoveTo(x + w, y + ah),
                PathCommand::LineTo(x + w * 0.4, y + ah),
                PathCommand::LineTo(x + w * 0.4, y),
                PathCommand::LineTo(x, y + h / 2.0),
                PathCommand::LineTo(x + w * 0.4, y + h),
                PathCommand::LineTo(x + w * 0.4, y + h - ah),
                PathCommand::LineTo(x + w, y + h - ah),
                PathCommand::Close,
            ])
        }
        "upArrow" => {
            let aw = w * 0.2;
            Some(vec![
                PathCommand::MoveTo(x + aw, y + h),
                PathCommand::LineTo(x + aw, y + h * 0.4),
                PathCommand::LineTo(x, y + h * 0.4),
                PathCommand::LineTo(x + w / 2.0, y),
                PathCommand::LineTo(x + w, y + h * 0.4),
                PathCommand::LineTo(x + w - aw, y + h * 0.4),
                PathCommand::LineTo(x + w - aw, y + h),
                PathCommand::Close,
            ])
        }
        "downArrow" => {
            let aw = w * 0.2;
            Some(vec![
                PathCommand::MoveTo(x + aw, y),
                PathCommand::LineTo(x + aw, y + h * 0.6),
                PathCommand::LineTo(x, y + h * 0.6),
                PathCommand::LineTo(x + w / 2.0, y + h),
                PathCommand::LineTo(x + w, y + h * 0.6),
                PathCommand::LineTo(x + w - aw, y + h * 0.6),
                PathCommand::LineTo(x + w - aw, y),
                PathCommand::Close,
            ])
        }
        "leftRightArrow" | "notchedRightArrow" => {
            let ah = h * 0.2;
            Some(vec![
                PathCommand::MoveTo(x, y + h / 2.0),
                PathCommand::LineTo(x + w * 0.2, y),
                PathCommand::LineTo(x + w * 0.2, y + ah),
                PathCommand::LineTo(x + w * 0.8, y + ah),
                PathCommand::LineTo(x + w * 0.8, y),
                PathCommand::LineTo(x + w, y + h / 2.0),
                PathCommand::LineTo(x + w * 0.8, y + h),
                PathCommand::LineTo(x + w * 0.8, y + h - ah),
                PathCommand::LineTo(x + w * 0.2, y + h - ah),
                PathCommand::LineTo(x + w * 0.2, y + h),
                PathCommand::Close,
            ])
        }
        "upDownArrow" => {
            let aw = w * 0.2;
            Some(vec![
                PathCommand::MoveTo(x + w / 2.0, y),
                PathCommand::LineTo(x + w, y + h * 0.2),
                PathCommand::LineTo(x + w - aw, y + h * 0.2),
                PathCommand::LineTo(x + w - aw, y + h * 0.8),
                PathCommand::LineTo(x + w, y + h * 0.8),
                PathCommand::LineTo(x + w / 2.0, y + h),
                PathCommand::LineTo(x, y + h * 0.8),
                PathCommand::LineTo(x + aw, y + h * 0.8),
                PathCommand::LineTo(x + aw, y + h * 0.2),
                PathCommand::LineTo(x, y + h * 0.2),
                PathCommand::Close,
            ])
        }
        "chevron" | "homePlate" => {
            Some(vec![
                PathCommand::MoveTo(x, y),
                PathCommand::LineTo(x + w * 0.8, y),
                PathCommand::LineTo(x + w, y + h / 2.0),
                PathCommand::LineTo(x + w * 0.8, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::LineTo(x + w * 0.2, y + h / 2.0),
                PathCommand::Close,
            ])
        }
        "plus" | "cross" => {
            let t = 0.33;
            Some(vec![
                PathCommand::MoveTo(x + w * t, y),
                PathCommand::LineTo(x + w * (1.0 - t), y),
                PathCommand::LineTo(x + w * (1.0 - t), y + h * t),
                PathCommand::LineTo(x + w, y + h * t),
                PathCommand::LineTo(x + w, y + h * (1.0 - t)),
                PathCommand::LineTo(x + w * (1.0 - t), y + h * (1.0 - t)),
                PathCommand::LineTo(x + w * (1.0 - t), y + h),
                PathCommand::LineTo(x + w * t, y + h),
                PathCommand::LineTo(x + w * t, y + h * (1.0 - t)),
                PathCommand::LineTo(x, y + h * (1.0 - t)),
                PathCommand::LineTo(x, y + h * t),
                PathCommand::LineTo(x + w * t, y + h * t),
                PathCommand::Close,
            ])
        }
        "wave" | "doubleWave" => {
            let steps = 32;
            let mut cmds = Vec::new();
            let amp = h * 0.15;
            for i in 0..=steps {
                let t = i as f64 / steps as f64;
                let px = x + t * w;
                let py = y + amp * (2.0 * PI * t).sin() + amp;
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            for i in (0..=steps).rev() {
                let t = i as f64 / steps as f64;
                let px = x + t * w;
                let py = y + h + amp * (2.0 * PI * t).sin() - amp;
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        "ribbon" | "ribbon2" | "ellipseRibbon" | "ellipseRibbon2" => {
            Some(vec![
                PathCommand::MoveTo(x, y + h * 0.3),
                PathCommand::LineTo(x + w * 0.15, y),
                PathCommand::LineTo(x + w * 0.15, y + h * 0.2),
                PathCommand::LineTo(x + w * 0.85, y + h * 0.2),
                PathCommand::LineTo(x + w * 0.85, y),
                PathCommand::LineTo(x + w, y + h * 0.3),
                PathCommand::LineTo(x + w * 0.85, y + h * 0.5),
                PathCommand::LineTo(x + w * 0.85, y + h),
                PathCommand::LineTo(x + w * 0.15, y + h),
                PathCommand::LineTo(x + w * 0.15, y + h * 0.5),
                PathCommand::Close,
            ])
        }
        "irregularSeal1" | "irregularSeal2" | "explosion1" | "explosion2" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let points = 12;
            let mut cmds = Vec::new();
            for i in 0..(points * 2) {
                let angle = 2.0 * PI * i as f64 / (points * 2) as f64;
                let jitter = if i % 2 == 0 { 1.0 } else { 0.55 + (i as f64 * 0.1).sin() * 0.15 };
                let px = cx + rx * jitter * angle.cos();
                let py = cy + ry * jitter * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        "cloud" | "cloudCallout" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let bumps = 16;
            let mut cmds = Vec::new();
            for i in 0..=bumps {
                let angle = 2.0 * PI * i as f64 / bumps as f64;
                let bump = 1.0 + 0.12 * (angle * 4.0).sin();
                let px = cx + rx * bump * angle.cos();
                let py = cy + ry * bump * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        "roundRect" | "flowChartAlternateProcess" => {
            let r = (w.min(h) * 0.15).min(10.0);
            Some(vec![
                PathCommand::MoveTo(x + r, y),
                PathCommand::LineTo(x + w - r, y),
                PathCommand::QuadTo(x + w, y, x + w, y + r),
                PathCommand::LineTo(x + w, y + h - r),
                PathCommand::QuadTo(x + w, y + h, x + w - r, y + h),
                PathCommand::LineTo(x + r, y + h),
                PathCommand::QuadTo(x, y + h, x, y + h - r),
                PathCommand::LineTo(x, y + r),
                PathCommand::QuadTo(x, y, x + r, y),
                PathCommand::Close,
            ])
        }
        "snip1Rect" | "snip2SameRect" | "snipRoundRect" => {
            let c = w.min(h) * 0.15;
            Some(vec![
                PathCommand::MoveTo(x, y),
                PathCommand::LineTo(x + w - c, y),
                PathCommand::LineTo(x + w, y + c),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
            ])
        }
        "can" | "flowChartMagneticDisk" => {
            let ry_top = h * 0.12;
            let steps = 24;
            let cx = x + w / 2.0;
            let mut cmds = Vec::new();
            for i in 0..=steps {
                let angle = PI + PI * i as f64 / steps as f64;
                let px = cx + (w / 2.0) * angle.cos();
                let py = y + ry_top + ry_top * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::LineTo(x + w, y + h - ry_top));
            for i in 0..=steps {
                let angle = PI * i as f64 / steps as f64;
                let px = cx + (w / 2.0) * angle.cos();
                let py = y + h - ry_top + ry_top * angle.sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::LineTo(x, y + ry_top));
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Flowchart: Process (simple rectangle)
        "flowChartProcess" => {
            Some(vec![
                PathCommand::MoveTo(x, y),
                PathCommand::LineTo(x + w, y),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
            ])
        }
        // Flowchart: Decision (diamond)
        "flowChartDecision" => {
            Some(vec![
                PathCommand::MoveTo(x + w / 2.0, y),
                PathCommand::LineTo(x + w, y + h / 2.0),
                PathCommand::LineTo(x + w / 2.0, y + h),
                PathCommand::LineTo(x, y + h / 2.0),
                PathCommand::Close,
            ])
        }
        // Flowchart: Terminator (stadium/rounded rectangle)
        "flowChartTerminator" => {
            let r = h / 2.0;
            Some(vec![
                PathCommand::MoveTo(x + r, y),
                PathCommand::LineTo(x + w - r, y),
                PathCommand::ArcTo(r, r, 0.0, false, true, x + w - r, y + h),
                PathCommand::LineTo(x + r, y + h),
                PathCommand::ArcTo(r, r, 0.0, false, true, x + r, y),
                PathCommand::Close,
            ])
        }
        // Flowchart: Document (rectangle with wavy bottom)
        "flowChartDocument" => {
            let wave_h = h * 0.08;
            let steps = 16;
            let mut cmds = vec![PathCommand::MoveTo(x, y)];
            cmds.push(PathCommand::LineTo(x + w, y));
            cmds.push(PathCommand::LineTo(x + w, y + h - wave_h));
            for i in 0..=steps {
                let t = i as f64 / steps as f64;
                let px = x + w - t * w;
                let py = y + h - wave_h + wave_h * (2.0 * PI * t * 3.0).sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::LineTo(x, y));
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Flowchart: Predefined Process (rectangle with double vertical lines)
        "flowChartPredefinedProcess" => {
            let margin = w * 0.1;
            Some(vec![
                PathCommand::MoveTo(x, y),
                PathCommand::LineTo(x + w, y),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
                // Left vertical line
                PathCommand::MoveTo(x + margin, y),
                PathCommand::LineTo(x + margin, y + h),
                // Right vertical line
                PathCommand::MoveTo(x + w - margin, y),
                PathCommand::LineTo(x + w - margin, y + h),
            ])
        }
        // Flowchart: Input/Output (parallelogram)
        "flowChartInputOutput" => {
            let offset = w * 0.2;
            Some(vec![
                PathCommand::MoveTo(x + offset, y),
                PathCommand::LineTo(x + w, y),
                PathCommand::LineTo(x + w - offset, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
            ])
        }
        // Flowchart: Preparation (hexagon)
        "flowChartPreparation" => {
            let offset = w * 0.2;
            Some(vec![
                PathCommand::MoveTo(x + offset, y),
                PathCommand::LineTo(x + w - offset, y),
                PathCommand::LineTo(x + w, y + h / 2.0),
                PathCommand::LineTo(x + w - offset, y + h),
                PathCommand::LineTo(x + offset, y + h),
                PathCommand::LineTo(x, y + h / 2.0),
                PathCommand::Close,
            ])
        }
        // Flowchart: Manual Input (trapezoid with slanted top)
        "flowChartManualInput" => {
            let offset = h * 0.15;
            Some(vec![
                PathCommand::MoveTo(x, y + offset),
                PathCommand::LineTo(x + w, y),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
            ])
        }
        // Flowchart: Manual Operation (trapezoid with slanted sides)
        "flowChartManualOperation" => {
            let offset = w * 0.15;
            Some(vec![
                PathCommand::MoveTo(x + offset, y),
                PathCommand::LineTo(x + w - offset, y),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
            ])
        }
        // Flowchart: Connector (circle)
        "flowChartConnector" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let r = w.min(h) / 2.0;
            let steps = 48;
            let mut cmds = Vec::new();
            for i in 0..=steps {
                let angle = 2.0 * PI * i as f64 / steps as f64;
                let px = cx + r * angle.cos();
                let py = cy + r * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Flowchart: Off-page Connector (home plate pentagon)
        "flowChartOffpageConnector" => {
            Some(vec![
                PathCommand::MoveTo(x, y),
                PathCommand::LineTo(x + w, y),
                PathCommand::LineTo(x + w, y + h * 0.75),
                PathCommand::LineTo(x + w / 2.0, y + h),
                PathCommand::LineTo(x, y + h * 0.75),
                PathCommand::Close,
            ])
        }
        // Flowchart: Sort (diamond with crossing lines)
        "flowChartSort" => {
            Some(vec![
                // Outer diamond
                PathCommand::MoveTo(x + w / 2.0, y),
                PathCommand::LineTo(x + w, y + h / 2.0),
                PathCommand::LineTo(x + w / 2.0, y + h),
                PathCommand::LineTo(x, y + h / 2.0),
                PathCommand::Close,
                // Horizontal line
                PathCommand::MoveTo(x, y + h / 2.0),
                PathCommand::LineTo(x + w, y + h / 2.0),
            ])
        }
        // Flowchart: Extract (triangle pointing down)
        "flowChartExtract" => {
            Some(vec![
                PathCommand::MoveTo(x, y),
                PathCommand::LineTo(x + w, y),
                PathCommand::LineTo(x + w / 2.0, y + h),
                PathCommand::Close,
            ])
        }
        // Flowchart: Merge (triangle pointing up)
        "flowChartMerge" => {
            Some(vec![
                PathCommand::MoveTo(x, y + h),
                PathCommand::LineTo(x + w / 2.0, y),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::Close,
            ])
        }
        // Flowchart: Delay (semi-circle on right side)
        "flowChartDelay" => {
            let steps = 24;
            let mut cmds = vec![PathCommand::MoveTo(x, y)];
            for i in 0..=steps {
                let angle = -PI / 2.0 + PI * i as f64 / steps as f64;
                let px = x + w + (w * 0.2) * angle.cos();
                let py = y + h / 2.0 + (h / 2.0) * angle.sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::LineTo(x, y + h));
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Flowchart: Display (curved sides)
        "flowChartDisplay" => {
            let steps = 16;
            let mut cmds = vec![PathCommand::MoveTo(x, y + h / 2.0)];
            // Top left curve
            for i in 0..=steps {
                let t = i as f64 / steps as f64;
                let px = x + w * 0.15 * t;
                let py = y + (h / 2.0) * (1.0 - t);
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::LineTo(x + w * 0.85, y));
            // Top right curve
            for i in 0..=steps {
                let t = i as f64 / steps as f64;
                let px = x + w * 0.85 + w * 0.15 * t;
                let py = y + (h / 2.0) * t;
                cmds.push(PathCommand::LineTo(px, py));
            }
            // Bottom right curve
            for i in 0..=steps {
                let t = i as f64 / steps as f64;
                let px = x + w - w * 0.15 * t;
                let py = y + h / 2.0 + (h / 2.0) * t;
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::LineTo(x + w * 0.15, y + h));
            // Bottom left curve
            for i in 0..=steps {
                let t = i as f64 / steps as f64;
                let px = x + w * 0.15 - w * 0.15 * t;
                let py = y + h - (h / 2.0) * t;
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Flowchart: Multidocument (stacked documents with wavy bottoms)
        "flowChartMultidocument" => {
            let wave_h = h * 0.06;
            let offset = w * 0.08;
            let steps = 12;
            let mut cmds = Vec::new();
            // Back document
            cmds.push(PathCommand::MoveTo(x + offset * 2.0, y));
            cmds.push(PathCommand::LineTo(x + w, y));
            cmds.push(PathCommand::LineTo(x + w, y + h * 0.3 - wave_h));
            for i in 0..=steps {
                let t = i as f64 / steps as f64;
                let px = x + w - t * (w - offset * 2.0);
                let py = y + h * 0.3 - wave_h + wave_h * (2.0 * PI * t * 2.0).sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            // Middle document
            cmds.push(PathCommand::MoveTo(x + offset, y + h * 0.15));
            cmds.push(PathCommand::LineTo(x + w - offset, y + h * 0.15));
            cmds.push(PathCommand::LineTo(x + w - offset, y + h * 0.5 - wave_h));
            for i in 0..=steps {
                let t = i as f64 / steps as f64;
                let px = x + w - offset - t * (w - offset * 2.0);
                let py = y + h * 0.5 - wave_h + wave_h * (2.0 * PI * t * 2.0).sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            // Front document
            cmds.push(PathCommand::MoveTo(x, y + h * 0.3));
            cmds.push(PathCommand::LineTo(x + w - offset * 2.0, y + h * 0.3));
            cmds.push(PathCommand::LineTo(x + w - offset * 2.0, y + h - wave_h));
            for i in 0..=steps {
                let t = i as f64 / steps as f64;
                let px = x + w - offset * 2.0 - t * (w - offset * 2.0);
                let py = y + h - wave_h + wave_h * (2.0 * PI * t * 2.0).sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Flowchart: Online Storage (curved bottom triangle)
        "flowChartOnlineStorage" => {
            let steps = 20;
            let mut cmds = vec![PathCommand::MoveTo(x, y)];
            cmds.push(PathCommand::LineTo(x + w, y));
            cmds.push(PathCommand::LineTo(x + w, y + h * 0.7));
            // Curved bottom
            for i in 0..=steps {
                let t = i as f64 / steps as f64;
                let px = x + w - t * w;
                let py = y + h * 0.7 + h * 0.3 * (PI * t).sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::LineTo(x, y));
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Special Shapes: Moon (crescent)
        "moon" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let rx = w / 2.0;
            let ry = h / 2.0;
            let steps = 48;
            let mut cmds = Vec::new();
            // Outer circle
            for i in 0..=steps {
                let angle = 2.0 * PI * i as f64 / steps as f64;
                let px = cx + rx * angle.cos();
                let py = cy + ry * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            // Inner crescent cutout (reverse direction for hole)
            let offset_x = w * 0.2;
            for i in (0..=steps).rev() {
                let angle = 2.0 * PI * i as f64 / steps as f64;
                let px = cx + offset_x + rx * 0.6 * angle.cos();
                let py = cy + ry * 0.6 * angle.sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Special Shapes: Smiley Face
        "smileyFace" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let r = w.min(h) / 2.0;
            let steps = 48;
            let mut cmds = Vec::new();
            // Face circle
            for i in 0..=steps {
                let angle = 2.0 * PI * i as f64 / steps as f64;
                let px = cx + r * angle.cos();
                let py = cy + r * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            // Left eye
            let eye_r = r * 0.08;
            let eye_y = cy - r * 0.25;
            for i in 0..=steps / 4 {
                let angle = 2.0 * PI * i as f64 / (steps / 4) as f64;
                let px = cx - r * 0.3 + eye_r * angle.cos();
                let py = eye_y + eye_r * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            // Right eye
            for i in 0..=steps / 4 {
                let angle = 2.0 * PI * i as f64 / (steps / 4) as f64;
                let px = cx + r * 0.3 + eye_r * angle.cos();
                let py = eye_y + eye_r * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            // Smile
            cmds.push(PathCommand::MoveTo(cx - r * 0.4, cy + r * 0.1));
            for i in 0..=steps / 4 {
                let t = i as f64 / (steps / 4) as f64;
                let angle = PI * 0.2 + PI * 0.6 * t;
                let px = cx + r * 0.4 * angle.cos();
                let py = cy + r * 0.4 * angle.sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            Some(cmds)
        }
        // Special Shapes: Sun (circle with rays)
        "sun" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let r_inner = w.min(h) * 0.35;
            let r_outer = w.min(h) * 0.5;
            let rays = 16;
            let steps_per_ray = 3;
            let mut cmds = Vec::new();
            for i in 0..(rays * steps_per_ray) {
                let angle = 2.0 * PI * i as f64 / (rays * steps_per_ray) as f64;
                let r = if i % steps_per_ray == 1 { r_outer } else { r_inner };
                let px = cx + r * angle.cos();
                let py = cy + r * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Special Shapes: No Smoking (circle with diagonal line)
        "noSmoking" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let r_outer = w.min(h) / 2.0;
            let r_inner = r_outer * 0.85;
            let line_w = r_outer * 0.15;
            let steps = 48;
            let mut cmds = Vec::new();
            // Outer circle
            for i in 0..=steps {
                let angle = 2.0 * PI * i as f64 / steps as f64;
                let px = cx + r_outer * angle.cos();
                let py = cy + r_outer * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            // Inner circle (reverse)
            for i in (0..=steps).rev() {
                let angle = 2.0 * PI * i as f64 / steps as f64;
                let px = cx + r_inner * angle.cos();
                let py = cy + r_inner * angle.sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            // Diagonal line (top-left to bottom-right)
            let offset = r_outer * 0.707;
            cmds.push(PathCommand::MoveTo(cx - offset, cy - offset - line_w));
            cmds.push(PathCommand::LineTo(cx - offset + line_w, cy - offset));
            cmds.push(PathCommand::LineTo(cx + offset, cy + offset + line_w));
            cmds.push(PathCommand::LineTo(cx + offset - line_w, cy + offset));
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Special Shapes: Folded Corner
        "foldedCorner" => {
            let fold = w.min(h) * 0.15;
            Some(vec![
                PathCommand::MoveTo(x, y),
                PathCommand::LineTo(x + w - fold, y),
                PathCommand::LineTo(x + w, y + fold),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
                // Fold triangle
                PathCommand::MoveTo(x + w - fold, y),
                PathCommand::LineTo(x + w - fold, y + fold),
                PathCommand::LineTo(x + w, y + fold),
                PathCommand::Close,
            ])
        }
        // Special Shapes: Frame (hollow rectangle)
        "frame" => {
            let thickness = w.min(h) * 0.15;
            Some(vec![
                // Outer rectangle
                PathCommand::MoveTo(x, y),
                PathCommand::LineTo(x + w, y),
                PathCommand::LineTo(x + w, y + h),
                PathCommand::LineTo(x, y + h),
                PathCommand::Close,
                // Inner rectangle (reverse direction for hole)
                PathCommand::MoveTo(x + thickness, y + thickness),
                PathCommand::LineTo(x + thickness, y + h - thickness),
                PathCommand::LineTo(x + w - thickness, y + h - thickness),
                PathCommand::LineTo(x + w - thickness, y + thickness),
                PathCommand::Close,
            ])
        }
        // Special Shapes: Bevel (3D beveled rectangle)
        "bevel" => {
            let bevel = w.min(h) * 0.12;
            Some(vec![
                // Outer shape
                PathCommand::MoveTo(x + bevel, y),
                PathCommand::LineTo(x + w - bevel, y),
                PathCommand::LineTo(x + w, y + bevel),
                PathCommand::LineTo(x + w, y + h - bevel),
                PathCommand::LineTo(x + w - bevel, y + h),
                PathCommand::LineTo(x + bevel, y + h),
                PathCommand::LineTo(x, y + h - bevel),
                PathCommand::LineTo(x, y + bevel),
                PathCommand::Close,
                // Inner bevel lines
                PathCommand::MoveTo(x + bevel, y),
                PathCommand::LineTo(x + bevel, y + bevel),
                PathCommand::LineTo(x + bevel, y + h - bevel),
                PathCommand::MoveTo(x + w - bevel, y),
                PathCommand::LineTo(x + w - bevel, y + bevel),
                PathCommand::LineTo(x + w - bevel, y + h - bevel),
                PathCommand::MoveTo(x + bevel, y + bevel),
                PathCommand::LineTo(x + w - bevel, y + bevel),
                PathCommand::MoveTo(x + bevel, y + h - bevel),
                PathCommand::LineTo(x + w - bevel, y + h - bevel),
            ])
        }
        // Special Shapes: Gear 6
        "gear6" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let r_outer = w.min(h) / 2.0;
            let r_inner = r_outer * 0.7;
            let r_tooth = r_outer * 0.85;
            let teeth = 6;
            let steps_per_tooth = 4;
            let mut cmds = Vec::new();
            for i in 0..(teeth * steps_per_tooth) {
                let angle = 2.0 * PI * i as f64 / (teeth * steps_per_tooth) as f64;
                let tooth_phase = i % steps_per_tooth;
                let r = match tooth_phase {
                    0 | 3 => r_tooth,
                    1 | 2 => r_outer,
                    _ => r_inner,
                };
                let px = cx + r * angle.cos();
                let py = cy + r * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            // Center hole
            let r_hole = r_inner * 0.5;
            let steps = 24;
            for i in (0..=steps).rev() {
                let angle = 2.0 * PI * i as f64 / steps as f64;
                let px = cx + r_hole * angle.cos();
                let py = cy + r_hole * angle.sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Special Shapes: Gear 9
        "gear9" => {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let r_outer = w.min(h) / 2.0;
            let r_inner = r_outer * 0.7;
            let r_tooth = r_outer * 0.85;
            let teeth = 9;
            let steps_per_tooth = 4;
            let mut cmds = Vec::new();
            for i in 0..(teeth * steps_per_tooth) {
                let angle = 2.0 * PI * i as f64 / (teeth * steps_per_tooth) as f64;
                let tooth_phase = i % steps_per_tooth;
                let r = match tooth_phase {
                    0 | 3 => r_tooth,
                    1 | 2 => r_outer,
                    _ => r_inner,
                };
                let px = cx + r * angle.cos();
                let py = cy + r * angle.sin();
                if i == 0 {
                    cmds.push(PathCommand::MoveTo(px, py));
                } else {
                    cmds.push(PathCommand::LineTo(px, py));
                }
            }
            cmds.push(PathCommand::Close);
            // Center hole
            let r_hole = r_inner * 0.5;
            let steps = 24;
            for i in (0..=steps).rev() {
                let angle = 2.0 * PI * i as f64 / steps as f64;
                let px = cx + r_hole * angle.cos();
                let py = cy + r_hole * angle.sin();
                cmds.push(PathCommand::LineTo(px, py));
            }
            cmds.push(PathCommand::Close);
            Some(cmds)
        }
        // Action Buttons: Blank (rounded rectangle)
        "actionButtonBlank" | "actionButtonHome" | "actionButtonHelp" => {
            let r = (w.min(h) * 0.08).min(6.0);
            Some(vec![
                PathCommand::MoveTo(x + r, y),
                PathCommand::LineTo(x + w - r, y),
                PathCommand::QuadTo(x + w, y, x + w, y + r),
                PathCommand::LineTo(x + w, y + h - r),
                PathCommand::QuadTo(x + w, y + h, x + w - r, y + h),
                PathCommand::LineTo(x + r, y + h),
                PathCommand::QuadTo(x, y + h, x, y + h - r),
                PathCommand::LineTo(x, y + r),
                PathCommand::QuadTo(x, y, x + r, y),
                PathCommand::Close,
            ])
        }
        _ => None,
    }
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
