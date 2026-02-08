// formats/pptx.rs - PPTX変換モジュール（レイアウト保持版）
//
// PPTX (Office Open XML Presentation) ファイルを解析し、
// シェイプの位置・サイズ・書式・画像を忠実に再現してドキュメントモデルに変換します。
// Officeソフトで開いてPDF化するのと同等の出力を目指します。

use crate::converter::{
    Color, ConvertError, Document, DocumentConverter, FontStyle, GradientStop, GradientType,
    Metadata, Page, PageElement, TextAlign,
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
    if let ShapeContent::Image { ref r_id } = shape.content {
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
                    let mut s = shape;
                    s.content = ShapeContent::ImageData {
                        data,
                        mime_type: mime.to_string(),
                    };
                    return s;
                }
            }
        }
    }
    shape
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
                let is_rounded = matches!(
                    shape.preset_geometry.as_deref(),
                    Some("roundRect") | Some("flowChartAlternateProcess")
                );

                // 3D effect: draw depth extrusion behind the shape
                if shape.has_3d && shape.width > 0.0 && shape.height > 0.0 {
                    let depth = 6.0; // 3D extrusion depth in points
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
                if is_ellipse {
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
                } else {
                    // Standard rectangle or rounded rect rendering
                    match &shape.fill {
                        Some(ShapeFill::Solid(color)) => {
                            if is_rounded {
                                // Approximate rounded rect: slightly inset ellipse corners
                                page.elements.push(PageElement::Rect {
                                    x: shape.x,
                                    y: shape.y,
                                    width: shape.width,
                                    height: shape.height,
                                    fill: Some(*color),
                                    stroke: None,
                                    stroke_width: 0.0,
                                });
                            } else {
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
                } // end else (non-ellipse)

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
