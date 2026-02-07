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
            PageElement::GradientRect {
                x,
                y,
                width: w,
                height: h,
                stops,
                gradient_type,
            } => {
                render_gradient_rect_to_pixels(
                    &mut pixels,
                    width,
                    height,
                    *x * scale,
                    *y * scale,
                    *w * scale,
                    *h * scale,
                    stops,
                    gradient_type,
                );
            }
            PageElement::Ellipse {
                cx,
                cy,
                rx,
                ry,
                fill,
                stroke: _,
                stroke_width: _,
            } => {
                if let Some(fill_color) = fill {
                    render_ellipse_to_pixels(
                        &mut pixels,
                        width,
                        height,
                        *cx * scale,
                        *cy * scale,
                        *rx * scale,
                        *ry * scale,
                        fill_color,
                    );
                }
            }
            PageElement::Image {
                x: img_x,
                y: img_y,
                width: img_w,
                height: img_h,
                data,
                mime_type,
            } => {
                render_image_to_pixels(
                    &mut pixels,
                    width,
                    height,
                    *img_x * scale,
                    *img_y * scale,
                    *img_w * scale,
                    *img_h * scale,
                    data,
                    mime_type,
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
    let mut x = x1 as i64;
    let mut y = y1 as i64;
    let dx = ((x2 - x1) as i64).abs();
    let dy = -((y2 - y1) as i64).abs();
    let sx: i64 = if x1 < x2 { 1 } else { -1 };
    let sy: i64 = if y1 < y2 { 1 } else { -1 };
    let mut err = dx + dy;

    let end_x = x2 as i64;
    let end_y = y2 as i64;

    // Limit iterations to avoid infinite loops on huge coordinates
    const MAX_LINE_ITERATIONS: usize = 100_000;
    let max_iter = (dx.unsigned_abs() + dy.unsigned_abs() + 2) as usize;
    for _ in 0..max_iter.min(MAX_LINE_ITERATIONS) {
        if x >= 0 && y >= 0 && (x as u32) < img_width && (y as u32) < img_height {
            set_pixel(pixels, img_width, x as u32, y as u32, color);
        }
        if x == end_x && y == end_y {
            break;
        }
        let e2 = 2i64.saturating_mul(err);
        if e2 >= dy {
            err = err.saturating_add(dy);
            x = x.saturating_add(sx);
        }
        if e2 <= dx {
            err = err.saturating_add(dx);
            y = y.saturating_add(sy);
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

/// グラデーション矩形をピクセルバッファに描画
fn render_gradient_rect_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    stops: &[crate::converter::GradientStop],
    gradient_type: &crate::converter::GradientType,
) {
    if stops.is_empty() || w <= 0.0 || h <= 0.0 {
        return;
    }

    let x0 = x.max(0.0) as u32;
    let y0 = y.max(0.0) as u32;
    let x1 = ((x + w) as u32).min(img_width);
    let y1 = ((y + h) as u32).min(img_height);

    for py in y0..y1 {
        for px in x0..x1 {
            // Calculate gradient position (0.0 to 1.0)
            let t = match gradient_type {
                crate::converter::GradientType::Linear(angle) => {
                    let local_x = (px as f64 - x) / w;
                    let local_y = (py as f64 - y) / h;
                    // Project onto gradient direction
                    let cos_a = angle.cos();
                    let sin_a = angle.sin();
                    let proj = local_x * sin_a + local_y * cos_a;
                    proj.clamp(0.0, 1.0)
                }
                crate::converter::GradientType::Radial => {
                    let cx = x + w / 2.0;
                    let cy = y + h / 2.0;
                    let dx = (px as f64 - cx) / (w / 2.0);
                    let dy = (py as f64 - cy) / (h / 2.0);
                    (dx * dx + dy * dy).sqrt().min(1.0)
                }
            };

            let color = interpolate_gradient(stops, t);
            let idx = ((py * img_width + px) * 4) as usize;
            if idx + 3 < pixels.len() {
                // Alpha blending
                let alpha = color.a as f64 / 255.0;
                pixels[idx] = (pixels[idx] as f64 * (1.0 - alpha) + color.r as f64 * alpha) as u8;
                pixels[idx + 1] =
                    (pixels[idx + 1] as f64 * (1.0 - alpha) + color.g as f64 * alpha) as u8;
                pixels[idx + 2] =
                    (pixels[idx + 2] as f64 * (1.0 - alpha) + color.b as f64 * alpha) as u8;
                pixels[idx + 3] = 255;
            }
        }
    }
}

/// グラデーション停止点間の色を補間
fn interpolate_gradient(stops: &[crate::converter::GradientStop], t: f64) -> Color {
    if stops.is_empty() {
        return Color::WHITE;
    }
    if stops.len() == 1 {
        return stops[0].color;
    }

    // Find the two stops to interpolate between
    if t <= stops[0].position {
        return stops[0].color;
    }
    if t >= stops[stops.len() - 1].position {
        return stops[stops.len() - 1].color;
    }

    for i in 0..stops.len() - 1 {
        if t >= stops[i].position && t <= stops[i + 1].position {
            let range = stops[i + 1].position - stops[i].position;
            if range <= 0.0 {
                return stops[i].color;
            }
            let local_t = (t - stops[i].position) / range;
            let c1 = &stops[i].color;
            let c2 = &stops[i + 1].color;
            return Color {
                r: (c1.r as f64 + (c2.r as f64 - c1.r as f64) * local_t) as u8,
                g: (c1.g as f64 + (c2.g as f64 - c1.g as f64) * local_t) as u8,
                b: (c1.b as f64 + (c2.b as f64 - c1.b as f64) * local_t) as u8,
                a: 255,
            };
        }
    }
    stops[0].color
}

/// 楕円をピクセルバッファに描画
fn render_ellipse_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    cx: f64,
    cy: f64,
    rx: f64,
    ry: f64,
    color: &Color,
) {
    let x0 = (cx - rx).max(0.0) as u32;
    let y0 = (cy - ry).max(0.0) as u32;
    let x1 = ((cx + rx) as u32).min(img_width);
    let y1 = ((cy + ry) as u32).min(img_height);

    for py in y0..y1 {
        for px in x0..x1 {
            let dx = (px as f64 - cx) / rx;
            let dy = (py as f64 - cy) / ry;
            if dx * dx + dy * dy <= 1.0 {
                set_pixel(pixels, img_width, px, py, color);
            }
        }
    }
}

/// 画像をピクセルバッファに描画（JPEG/PNGデコード）
fn render_image_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    data: &[u8],
    _mime_type: &str,
) {
    // Try PNG decoding first
    if let Some(decoded) = decode_png_image(data) {
        blit_decoded_image(pixels, img_width, img_height, x, y, w, h, &decoded);
        return;
    }
    // Try JPEG decoding
    if let Some(decoded) = decode_jpeg_image(data) {
        blit_decoded_image(pixels, img_width, img_height, x, y, w, h, &decoded);
        return;
    }
    // Fallback: render placeholder rect
    render_rect_to_pixels(
        pixels, img_width, img_height, x, y, w, h,
        Some(&Color::rgb(220, 220, 220)),
        Some(&Color::rgb(180, 180, 180)),
        1.0,
    );
}

/// デコードされた画像データ
struct DecodedImage {
    width: u32,
    height: u32,
    pixels: Vec<u8>, // RGBA
}

/// PNG画像をデコード
fn decode_png_image(data: &[u8]) -> Option<DecodedImage> {
    let decoder = png::Decoder::new(std::io::Cursor::new(data));
    let mut reader = decoder.read_info().ok()?;
    let mut buf = vec![0u8; reader.output_buffer_size()];
    let info = reader.next_frame(&mut buf).ok()?;
    let width = info.width;
    let height = info.height;

    // Convert to RGBA
    let rgba = match info.color_type {
        png::ColorType::Rgba => buf[..info.buffer_size()].to_vec(),
        png::ColorType::Rgb => {
            let mut rgba = Vec::with_capacity((width * height * 4) as usize);
            for chunk in buf[..info.buffer_size()].chunks(3) {
                rgba.push(chunk[0]);
                rgba.push(chunk[1]);
                rgba.push(chunk[2]);
                rgba.push(255);
            }
            rgba
        }
        png::ColorType::Grayscale => {
            let mut rgba = Vec::with_capacity((width * height * 4) as usize);
            for &g in &buf[..info.buffer_size()] {
                rgba.push(g);
                rgba.push(g);
                rgba.push(g);
                rgba.push(255);
            }
            rgba
        }
        png::ColorType::GrayscaleAlpha => {
            let mut rgba = Vec::with_capacity((width * height * 4) as usize);
            for chunk in buf[..info.buffer_size()].chunks(2) {
                rgba.push(chunk[0]);
                rgba.push(chunk[0]);
                rgba.push(chunk[0]);
                rgba.push(chunk[1]);
            }
            rgba
        }
        _ => return None,
    };

    Some(DecodedImage { width, height, pixels: rgba })
}

/// JPEG画像をデコード（簡易実装）
fn decode_jpeg_image(data: &[u8]) -> Option<DecodedImage> {
    // Check JPEG magic bytes
    if data.len() < 4 || data[0] != 0xFF || data[1] != 0xD8 {
        return None;
    }

    // Parse JPEG markers to get dimensions
    let mut i = 2;
    let mut width = 0u32;
    let mut height = 0u32;
    while i + 4 < data.len() {
        if data[i] != 0xFF {
            i += 1;
            continue;
        }
        let marker = data[i + 1];
        if marker == 0xD9 {
            break; // EOI
        }
        if i + 3 >= data.len() {
            break;
        }
        let length = ((data[i + 2] as usize) << 8) | (data[i + 3] as usize);

        // SOF markers (Start of Frame)
        if (0xC0..=0xC3).contains(&marker) || (0xC5..=0xC7).contains(&marker) {
            if i + 8 < data.len() {
                height = ((data[i + 5] as u32) << 8) | (data[i + 6] as u32);
                width = ((data[i + 7] as u32) << 8) | (data[i + 8] as u32);
            }
            break;
        }
        i += 2 + length;
    }

    if width == 0 || height == 0 {
        return None;
    }

    // For JPEG, we produce a placeholder with estimated dimensions
    // Full JPEG decode requires complex DCT/Huffman - placeholder with solid color
    let mut pixels = vec![0u8; (width * height * 4) as usize];
    // Fill with a mid-gray as placeholder (indicates image position)
    for chunk in pixels.chunks_exact_mut(4) {
        chunk[0] = 200;
        chunk[1] = 200;
        chunk[2] = 200;
        chunk[3] = 255;
    }
    Some(DecodedImage { width, height, pixels })
}

/// デコードされた画像をターゲットバッファにブリット（スケーリング付き）
fn blit_decoded_image(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    src: &DecodedImage,
) {
    if src.width == 0 || src.height == 0 || w <= 0.0 || h <= 0.0 {
        return;
    }

    let x0 = x.max(0.0) as u32;
    let y0 = y.max(0.0) as u32;
    let x1 = ((x + w) as u32).min(img_width);
    let y1 = ((y + h) as u32).min(img_height);

    for py in y0..y1 {
        for px in x0..x1 {
            // Map to source pixel (nearest neighbor)
            let src_x = ((px as f64 - x) / w * src.width as f64) as u32;
            let src_y = ((py as f64 - y) / h * src.height as f64) as u32;
            if src_x < src.width && src_y < src.height {
                let src_idx = ((src_y * src.width + src_x) * 4) as usize;
                let dst_idx = ((py * img_width + px) * 4) as usize;
                if src_idx + 3 < src.pixels.len() && dst_idx + 3 < pixels.len() {
                    let src_a = src.pixels[src_idx + 3] as f64 / 255.0;
                    // Thresholds for fast-path (fully opaque) and skip (fully transparent)
                    const ALPHA_OPAQUE_THRESHOLD: f64 = 0.99;
                    const ALPHA_TRANSPARENT_THRESHOLD: f64 = 0.01;
                    if src_a > ALPHA_OPAQUE_THRESHOLD {
                        pixels[dst_idx] = src.pixels[src_idx];
                        pixels[dst_idx + 1] = src.pixels[src_idx + 1];
                        pixels[dst_idx + 2] = src.pixels[src_idx + 2];
                        pixels[dst_idx + 3] = 255;
                    } else if src_a > ALPHA_TRANSPARENT_THRESHOLD {
                        pixels[dst_idx] = (pixels[dst_idx] as f64 * (1.0 - src_a)
                            + src.pixels[src_idx] as f64 * src_a)
                            as u8;
                        pixels[dst_idx + 1] = (pixels[dst_idx + 1] as f64 * (1.0 - src_a)
                            + src.pixels[src_idx + 1] as f64 * src_a)
                            as u8;
                        pixels[dst_idx + 2] = (pixels[dst_idx + 2] as f64 * (1.0 - src_a)
                            + src.pixels[src_idx + 2] as f64 * src_a)
                            as u8;
                        pixels[dst_idx + 3] = 255;
                    }
                }
            }
        }
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
