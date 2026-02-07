// image_renderer.rs - ページ画像化 + ZIP出力モジュール
//
// ドキュメントの各ページをPNG画像にレンダリングし、
// ZIPファイルにまとめて出力します。

use crate::converter::{Color, Document, FontStyle, Page, PageElement};
use crate::font_manager::FontManager;

/// 画像レンダリングの設定
pub struct ImageRenderConfig {
    /// DPI（デフォルト: 150）
    pub dpi: f64,
    /// 背景色
    pub background: Color,
    /// 画像フォーマット
    pub format: ImageFormat,
}

impl Default for ImageRenderConfig {
    fn default() -> Self {
        Self {
            dpi: 150.0,
            background: Color::WHITE,
            format: ImageFormat::Png,
        }
    }
}

/// 画像フォーマット
#[derive(Debug, Clone, Copy)]
pub enum ImageFormat {
    Png,
}

/// ページを画像バイト列にレンダリング
pub fn render_page_to_image(
    page: &Page,
    config: &ImageRenderConfig,
    _font_manager: &FontManager,
) -> Vec<u8> {
    let scale = config.dpi / 72.0;
    let width = (page.width * scale) as u32;
    let height = (page.height * scale) as u32;

    // RGBAピクセルバッファを作成（白背景）
    let mut pixels = vec![0u8; (width * height * 4) as usize];
    for i in (0..pixels.len()).step_by(4) {
        pixels[i] = config.background.r;
        pixels[i + 1] = config.background.g;
        pixels[i + 2] = config.background.b;
        pixels[i + 3] = config.background.a;
    }

    // 各要素をレンダリング
    for element in &page.elements {
        match element {
            PageElement::Text {
                x,
                y,
                width: _,
                text,
                style,
                align: _,
            } => {
                render_text_to_pixels(
                    &mut pixels, width, height, *x, *y, text, style, scale,
                );
            }
            PageElement::Rect {
                x,
                y,
                width: w,
                height: h,
                fill,
                stroke,
                stroke_width,
            } => {
                render_rect_to_pixels(
                    &mut pixels,
                    width,
                    height,
                    *x * scale,
                    *y * scale,
                    *w * scale,
                    *h * scale,
                    fill.as_ref(),
                    stroke.as_ref(),
                    *stroke_width * scale,
                );
            }
            PageElement::Line {
                x1,
                y1,
                x2,
                y2,
                width: w,
                color,
            } => {
                render_line_to_pixels(
                    &mut pixels,
                    width,
                    height,
                    *x1 * scale,
                    *y1 * scale,
                    *x2 * scale,
                    *y2 * scale,
                    *w * scale,
                    color,
                );
            }
            _ => {}
        }
    }

    // PNGにエンコード
    encode_png(&pixels, width, height)
}

/// テキストをピクセルバッファに描画
/// 注意: これは簡易実装であり、実際のグリフラスタライズではなく
/// 各文字を矩形ブロックとして描画します。完全な実装では
/// ab_glyph等のフォントラスタライザーを使用して
/// 正確なグリフ形状をレンダリングする必要があります。
fn render_text_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    x: f64,
    y: f64,
    text: &str,
    style: &FontStyle,
    scale: f64,
) {
    // 簡易テキスト描画: 各文字を小さなドットパターンで表現
    let px = (x * scale) as i32;
    let py = (y * scale) as i32;
    let font_px = (style.font_size * scale) as i32;
    let char_width = font_px * 6 / 10; // 半角文字幅

    for (i, ch) in text.chars().enumerate() {
        let cw = if ch.is_ascii() { char_width } else { font_px };
        let cx = px + i as i32 * char_width.max(1);
        let cy = py;

        // 文字領域をフォント色で塗りつぶし（簡易表現）
        if !ch.is_whitespace() {
            for dy in 2..font_px.saturating_sub(2) {
                for dx in 1..cw.saturating_sub(1) {
                    let pixel_x = (cx + dx) as u32;
                    let pixel_y = (cy + dy) as u32;
                    if pixel_x < img_width && pixel_y < img_height {
                        let idx = ((pixel_y * img_width + pixel_x) * 4) as usize;
                        if idx + 3 < pixels.len() {
                            pixels[idx] = style.color.r;
                            pixels[idx + 1] = style.color.g;
                            pixels[idx + 2] = style.color.b;
                            pixels[idx + 3] = 255;
                        }
                    }
                }
            }
        }
    }
}

/// 矩形をピクセルバッファに描画
fn render_rect_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    fill: Option<&Color>,
    stroke: Option<&Color>,
    _stroke_width: f64,
) {
    if let Some(fill_color) = fill {
        let x0 = x.max(0.0) as u32;
        let y0 = y.max(0.0) as u32;
        let x1 = ((x + w) as u32).min(img_width);
        let y1 = ((y + h) as u32).min(img_height);

        for py in y0..y1 {
            for px in x0..x1 {
                let idx = ((py * img_width + px) * 4) as usize;
                if idx + 3 < pixels.len() {
                    pixels[idx] = fill_color.r;
                    pixels[idx + 1] = fill_color.g;
                    pixels[idx + 2] = fill_color.b;
                    pixels[idx + 3] = fill_color.a;
                }
            }
        }
    }
    if let Some(stroke_color) = stroke {
        let x0 = x.max(0.0) as u32;
        let y0 = y.max(0.0) as u32;
        let x1 = ((x + w) as u32).min(img_width.saturating_sub(1));
        let y1 = ((y + h) as u32).min(img_height.saturating_sub(1));

        // 上辺と下辺
        for px in x0..=x1 {
            set_pixel(pixels, img_width, px, y0, stroke_color);
            set_pixel(pixels, img_width, px, y1, stroke_color);
        }
        // 左辺と右辺
        for py in y0..=y1 {
            set_pixel(pixels, img_width, x0, py, stroke_color);
            set_pixel(pixels, img_width, x1, py, stroke_color);
        }
    }
}

/// 直線をピクセルバッファに描画（ブレゼンハムアルゴリズム）
fn render_line_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    x1: f64,
    y1: f64,
    x2: f64,
    y2: f64,
    _width: f64,
    color: &Color,
) {
    let mut x = x1 as i32;
    let mut y = y1 as i32;
    let dx = ((x2 - x1) as i32).abs();
    let dy = -((y2 - y1) as i32).abs();
    let sx = if x1 < x2 { 1 } else { -1 };
    let sy = if y1 < y2 { 1 } else { -1 };
    let mut err = dx + dy;

    let end_x = x2 as i32;
    let end_y = y2 as i32;

    loop {
        if x >= 0 && y >= 0 && (x as u32) < img_width && (y as u32) < img_height {
            set_pixel(pixels, img_width, x as u32, y as u32, color);
        }
        if x == end_x && y == end_y {
            break;
        }
        let e2 = 2 * err;
        if e2 >= dy {
            err += dy;
            x += sx;
        }
        if e2 <= dx {
            err += dx;
            y += sy;
        }
    }
}

/// ピクセルを設定
fn set_pixel(pixels: &mut [u8], width: u32, x: u32, y: u32, color: &Color) {
    let idx = ((y * width + x) * 4) as usize;
    if idx + 3 < pixels.len() {
        pixels[idx] = color.r;
        pixels[idx + 1] = color.g;
        pixels[idx + 2] = color.b;
        pixels[idx + 3] = color.a;
    }
}

/// RGBAピクセルデータをPNGにエンコード
fn encode_png(pixels: &[u8], width: u32, height: u32) -> Vec<u8> {
    let mut output = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut output, width, height);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder.set_compression(png::Compression::Fast);
        if let Ok(mut writer) = encoder.write_header() {
            let _ = writer.write_image_data(pixels);
        }
    }
    output
}

/// ドキュメント全ページを画像化してZIPにまとめる
pub fn render_to_images_zip(doc: &Document, font_manager: &FontManager) -> Vec<u8> {
    render_to_images_zip_with_config(doc, font_manager, &ImageRenderConfig::default())
}

/// 設定指定でドキュメント全ページを画像化してZIPにまとめる
pub fn render_to_images_zip_with_config(
    doc: &Document,
    font_manager: &FontManager,
    config: &ImageRenderConfig,
) -> Vec<u8> {
    use std::io::Write;

    let mut zip_buffer = Vec::new();
    {
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(&mut zip_buffer));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for (i, page) in doc.pages.iter().enumerate() {
            let image_data = render_page_to_image(page, config, font_manager);
            let filename = format!("page_{:04}.png", i + 1);
            if zip.start_file(&filename, options).is_ok() {
                let _ = zip.write_all(&image_data);
            }
        }

        let _ = zip.finish();
    }
    zip_buffer
}
