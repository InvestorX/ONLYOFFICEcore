// image_renderer.rs - ページ画像化 + ZIP出力モジュール
//
// ドキュメントの各ページをPNG画像にレンダリングし、
// ZIPファイルにまとめて出力します。

use crate::converter::{Color, Document, FontStyle, Page, PageElement};
use crate::font_manager::FontManager;
use ab_glyph::{Font, FontRef, PxScale, ScaleFont};

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
    font_manager: &FontManager,
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
                    font_manager,
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
            PageElement::EllipseImage {
                cx,
                cy,
                rx,
                ry,
                data,
                mime_type,
                stroke,
                stroke_width,
            } => {
                render_ellipse_image_to_pixels(
                    &mut pixels,
                    width,
                    height,
                    *cx * scale,
                    *cy * scale,
                    *rx * scale,
                    *ry * scale,
                    data,
                    mime_type,
                );
                // Render stroke if specified
                if let Some(stroke_color) = stroke {
                    if *stroke_width > 0.0 {
                        render_ellipse_stroke_to_pixels(
                            &mut pixels,
                            width,
                            height,
                            *cx * scale,
                            *cy * scale,
                            *rx * scale,
                            *ry * scale,
                            stroke_color,
                            *stroke_width * scale,
                        );
                    }
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
            PageElement::Path {
                commands,
                fill,
                stroke,
                stroke_width,
            } => {
                render_path_to_pixels(
                    &mut pixels, width, height,
                    commands, fill.as_ref(), stroke.as_ref(), *stroke_width, scale,
                );
            }
            PageElement::PathImage {
                commands,
                data,
                mime_type,
                stroke,
                stroke_width,
            } => {
                // First render the clipped image
                render_path_image_to_pixels(
                    &mut pixels, width, height,
                    commands, data, mime_type, scale,
                );
                // Then render stroke if specified
                if let Some(stroke_color) = stroke.as_ref() {
                    if *stroke_width > 0.0 {
                        render_path_to_pixels(
                            &mut pixels, width, height,
                            commands,
                            None,
                            Some(stroke_color),
                            *stroke_width,
                            scale,
                        );
                    }
                }
            }
            PageElement::TableBlock {
                x: tbl_x,
                y: tbl_y,
                width: tbl_w,
                table,
            } => {
                // Render table: draw grid lines and cell text
                let row_height = 20.0; // Default row height in points
                let padding = 4.0;
                let mut cy = *tbl_y;

                // Calculate column widths
                let col_widths = if table.column_widths.is_empty() {
                    let ncols = table.rows.first().map_or(1, |r| r.len().max(1));
                    vec![*tbl_w / ncols as f64; ncols]
                } else {
                    table.column_widths.clone()
                };

                for row in &table.rows {
                    let mut cx = *tbl_x;
                    for (ci, cell) in row.iter().enumerate() {
                        let cw = col_widths.get(ci).copied().unwrap_or(60.0);

                        // Draw cell border
                        render_rect_to_pixels(
                            &mut pixels, width, height,
                            cx * scale, cy * scale, cw * scale, row_height * scale,
                            None,
                            Some(&Color::rgb(128, 128, 128)),
                            1.0,
                        );

                        // Draw cell text
                        if !cell.text.is_empty() {
                            render_text_to_pixels(
                                &mut pixels, width, height,
                                cx + padding, cy + padding,
                                &cell.text, &cell.style, scale, font_manager,
                            );
                        }

                        cx += cw;
                    }
                    cy += row_height;
                }
            }
        }
    }

    // PNGにエンコード
    encode_png(&pixels, width, height)
}

/// テキストをピクセルバッファに描画
/// ab_glyphフォントラスタライザーを使用して正確なグリフ形状をレンダリングします。
/// フォントが利用できない場合は簡易矩形フォールバックを使用します。
fn render_text_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    x: f64,
    y: f64,
    text: &str,
    style: &FontStyle,
    scale: f64,
    font_manager: &FontManager,
) {
    let font_size_px = (style.font_size * scale) as f32;
    if font_size_px <= 0.0 || text.is_empty() {
        return;
    }

    // Try to get font data and render with ab_glyph
    let font_data = font_manager.resolve_font(&style.font_name)
        .or_else(|| font_manager.best_font_data());

    if let Some(data) = font_data {
        if let Ok(font) = FontRef::try_from_slice(data) {
            render_text_with_font(
                pixels, img_width, img_height, x, y, text, style, scale, &font,
            );
            return;
        }
    }

    // Fallback: simple rectangle rendering when no font available
    render_text_fallback(pixels, img_width, img_height, x, y, text, style, scale);
}

/// ab_glyphフォントを使用してテキストをレンダリング
fn render_text_with_font(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    x: f64,
    y: f64,
    text: &str,
    style: &FontStyle,
    scale: f64,
    font: &FontRef,
) {
    let font_size_px = (style.font_size * scale) as f32;
    let px_scale = PxScale::from(font_size_px);
    let scaled_font = font.as_scaled(px_scale);

    let ascent = scaled_font.ascent();
    let start_x = (x * scale) as f32;
    let start_y = (y * scale) as f32 + ascent;

    let mut cursor_x = start_x;

    for ch in text.chars() {
        let glyph_id = font.glyph_id(ch);
        let advance = scaled_font.h_advance(glyph_id);

        if !ch.is_whitespace() {
            let glyph = glyph_id.with_scale_and_position(
                px_scale,
                ab_glyph::point(cursor_x, start_y),
            );

            if let Some(outlined) = font.outline_glyph(glyph) {
                let bounds = outlined.px_bounds();
                outlined.draw(|gx, gy, coverage| {
                    const MIN_COVERAGE: f32 = 0.01;
                    if coverage > MIN_COVERAGE {
                        let px = (bounds.min.x as i32 + gx as i32) as u32;
                        let py = (bounds.min.y as i32 + gy as i32) as u32;
                        if px < img_width && py < img_height {
                            let idx = ((py * img_width + px) * 4) as usize;
                            if idx + 3 < pixels.len() {
                                let alpha = coverage.min(1.0);
                                pixels[idx] = blend_channel(pixels[idx], style.color.r, alpha);
                                pixels[idx + 1] = blend_channel(pixels[idx + 1], style.color.g, alpha);
                                pixels[idx + 2] = blend_channel(pixels[idx + 2], style.color.b, alpha);
                                pixels[idx + 3] = 255;
                            }
                        }
                    }
                });
            }
        }

        cursor_x += advance;
    }
}

/// フォントが利用できない場合の簡易テキスト描画フォールバック
fn render_text_fallback(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    x: f64,
    y: f64,
    text: &str,
    style: &FontStyle,
    scale: f64,
) {
    let px = (x * scale) as i32;
    let py = (y * scale) as i32;
    let font_px = (style.font_size * scale) as i32;
    let char_width = font_px * 6 / 10;

    let mut cursor_x = px;
    for ch in text.chars() {
        let cw = if ch.is_ascii() { char_width } else { font_px };

        if !ch.is_whitespace() {
            for dy in 2..font_px.saturating_sub(2) {
                for dx in 1..cw.saturating_sub(1) {
                    let pixel_x = (cursor_x + dx) as u32;
                    let pixel_y = (py + dy) as u32;
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

        cursor_x += cw;
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

/// Parse path commands into separate subpaths (helper function)
/// Each subpath starts with MoveTo and ends with Close or another MoveTo
fn parse_path_commands_to_subpaths(
    commands: &[crate::converter::PathCommand],
    scale: f64,
) -> Vec<Vec<(f64, f64)>> {
    use crate::converter::PathCommand;

    let mut subpaths: Vec<Vec<(f64, f64)>> = Vec::new();
    let mut current_subpath: Vec<(f64, f64)> = Vec::new();
    let mut subpath_start = (0.0, 0.0);
    let mut cx = 0.0;
    let mut cy = 0.0;

    for cmd in commands {
        match cmd {
            PathCommand::MoveTo(x, y) => {
                // Start a new subpath
                if !current_subpath.is_empty() {
                    subpaths.push(std::mem::take(&mut current_subpath));
                }
                cx = *x * scale;
                cy = *y * scale;
                subpath_start = (cx, cy);
                current_subpath.push((cx, cy));
            }
            PathCommand::LineTo(x, y) => {
                cx = *x * scale;
                cy = *y * scale;
                current_subpath.push((cx, cy));
            }
            PathCommand::QuadTo(qcx, qcy, x, y) => {
                let qcx = *qcx * scale;
                let qcy = *qcy * scale;
                let ex = *x * scale;
                let ey = *y * scale;
                let steps = 8;
                for i in 1..=steps {
                    let t = i as f64 / steps as f64;
                    let it = 1.0 - t;
                    let px = it * it * cx + 2.0 * it * t * qcx + t * t * ex;
                    let py = it * it * cy + 2.0 * it * t * qcy + t * t * ey;
                    current_subpath.push((px, py));
                }
                cx = ex;
                cy = ey;
            }
            PathCommand::CubicTo(cx1, cy1, cx2, cy2, x, y) => {
                let c1x = *cx1 * scale;
                let c1y = *cy1 * scale;
                let c2x = *cx2 * scale;
                let c2y = *cy2 * scale;
                let ex = *x * scale;
                let ey = *y * scale;
                let steps = 12;
                for i in 1..=steps {
                    let t = i as f64 / steps as f64;
                    let it = 1.0 - t;
                    let px = it*it*it*cx + 3.0*it*it*t*c1x + 3.0*it*t*t*c2x + t*t*t*ex;
                    let py = it*it*it*cy + 3.0*it*it*t*c1y + 3.0*it*t*t*c2y + t*t*t*ey;
                    current_subpath.push((px, py));
                }
                cx = ex;
                cy = ey;
            }
            PathCommand::ArcTo(_rx, _ry, _rot, _large, _sweep, x, y) => {
                // Approximate arc as line segments
                let ex = *x * scale;
                let ey = *y * scale;
                let steps = 12;
                for i in 1..=steps {
                    let t = i as f64 / steps as f64;
                    let px = cx + (ex - cx) * t;
                    let py = cy + (ey - cy) * t;
                    current_subpath.push((px, py));
                }
                cx = ex;
                cy = ey;
            }
            PathCommand::Close => {
                // Close current subpath to its starting point
                if !current_subpath.is_empty() && current_subpath[0] != (cx, cy) {
                    current_subpath.push(subpath_start);
                }
                // After closing, the current point becomes the subpath start
                cx = subpath_start.0;
                cy = subpath_start.1;
            }
        }
    }

    // Add the last subpath if it exists
    if !current_subpath.is_empty() {
        subpaths.push(current_subpath);
    }

    subpaths
}

/// パスをピクセルバッファに描画（多角形塗りつぶし + ストローク）
fn render_path_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    commands: &[crate::converter::PathCommand],
    fill: Option<&Color>,
    stroke: Option<&Color>,
    _stroke_width: f64,
    scale: f64,
) {
    // Parse path commands into separate subpaths using helper function
    let subpaths = parse_path_commands_to_subpaths(commands, scale);

    if subpaths.is_empty() {
        return;
    }

    // Fill all subpaths using even-odd fill rule (handles shapes with holes)
    if let Some(fill_color) = fill {
        if !subpaths.is_empty() {
            // Find the bounding box for all subpaths combined
            let mut min_y = f64::INFINITY;
            let mut max_y = f64::NEG_INFINITY;
            for subpath in &subpaths {
                for &(_, y) in subpath {
                    min_y = min_y.min(y);
                    max_y = max_y.max(y);
                }
            }
            let min_y = min_y.max(0.0) as u32;
            let max_y = max_y.min(img_height as f64) as u32;

            // Scanline algorithm with even-odd fill rule
            for scan_y in min_y..max_y {
                let y_f = scan_y as f64 + 0.5;
                let mut intersections: Vec<f64> = Vec::new();

                // Collect intersections from ALL subpaths
                for subpath in &subpaths {
                    if subpath.len() < 2 {
                        continue;
                    }

                    for i in 0..subpath.len().saturating_sub(1) {
                        let (x1, y1) = subpath[i];
                        let (x2, y2) = subpath[i + 1];

                        if (y1 <= y_f && y2 > y_f) || (y2 <= y_f && y1 > y_f) {
                            let t = (y_f - y1) / (y2 - y1);
                            intersections.push(x1 + t * (x2 - x1));
                        }
                    }
                }

                intersections.sort_by(|a, b| {
                    // Use total ordering to ensure deterministic results
                    a.partial_cmp(b).unwrap_or_else(|| {
                        // Handle NaN cases: NaN values come last
                        if a.is_nan() && b.is_nan() {
                            std::cmp::Ordering::Equal
                        } else if a.is_nan() {
                            std::cmp::Ordering::Greater
                        } else {
                            std::cmp::Ordering::Less
                        }
                    })
                });

                // Even-odd rule: fill between pairs of intersections
                for pair in intersections.chunks(2) {
                    if pair.len() == 2 {
                        let x_start = pair[0].max(0.0) as u32;
                        let x_end = (pair[1] as u32).min(img_width);
                        for px in x_start..x_end {
                            set_pixel(pixels, img_width, px, scan_y, fill_color);
                        }
                    }
                }
            }
        }
    }

    // Stroke each subpath
    if let Some(stroke_color) = stroke {
        for subpath in &subpaths {
            for i in 0..subpath.len().saturating_sub(1) {
                let (x1, y1) = subpath[i];
                let (x2, y2) = subpath[i + 1];
                render_line_to_pixels(
                    pixels, img_width, img_height,
                    x1, y1, x2, y2, 1.0, stroke_color,
                );
            }
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

/// アルファブレンド: 背景チャンネルと前景チャンネルをアルファ値で合成
#[inline]
fn blend_channel(bg: u8, fg: u8, alpha: f32) -> u8 {
    (bg as f32 * (1.0 - alpha) + fg as f32 * alpha) as u8
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

/// 楕円のストローク（輪郭）をピクセルバッファに描画
fn render_ellipse_stroke_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    cx: f64,
    cy: f64,
    rx: f64,
    ry: f64,
    color: &Color,
    stroke_width: f64,
) {
    // Use scanline-based approach for ellipse stroke
    // Draw the outer ellipse and exclude the inner ellipse
    let outer_rx = rx + stroke_width / 2.0;
    let outer_ry = ry + stroke_width / 2.0;
    let inner_rx = (rx - stroke_width / 2.0).max(0.0);
    let inner_ry = (ry - stroke_width / 2.0).max(0.0);

    let x0 = (cx - outer_rx).max(0.0) as u32;
    let y0 = (cy - outer_ry).max(0.0) as u32;
    let x1 = ((cx + outer_rx) as u32).min(img_width);
    let y1 = ((cy + outer_ry) as u32).min(img_height);

    for py in y0..y1 {
        for px in x0..x1 {
            let dx_outer = (px as f64 - cx) / outer_rx;
            let dy_outer = (py as f64 - cy) / outer_ry;
            let dist_outer = dx_outer * dx_outer + dy_outer * dy_outer;

            if dist_outer <= 1.0 {
                // Point is inside outer ellipse
                if inner_rx > 0.0 && inner_ry > 0.0 {
                    let dx_inner = (px as f64 - cx) / inner_rx;
                    let dy_inner = (py as f64 - cy) / inner_ry;
                    let dist_inner = dx_inner * dx_inner + dy_inner * dy_inner;

                    if dist_inner > 1.0 {
                        // Point is outside inner ellipse -> it's in the stroke area
                        set_pixel(pixels, img_width, px, py, color);
                    }
                } else {
                    // No inner ellipse (stroke is too wide), fill everything
                    set_pixel(pixels, img_width, px, py, color);
                }
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

/// 楕円クリップされた画像をピクセルバッファに描画
fn render_ellipse_image_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    cx: f64,
    cy: f64,
    rx: f64,
    ry: f64,
    data: &[u8],
    _mime_type: &str,
) {
    // Decode the image
    let decoded = if let Some(img) = decode_png_image(data) {
        img
    } else if let Some(img) = decode_jpeg_image(data) {
        img
    } else {
        // Fallback: render placeholder
        render_rect_to_pixels(
            pixels, img_width, img_height,
            cx - rx, cy - ry, rx * 2.0, ry * 2.0,
            Some(&Color::rgb(220, 220, 220)),
            Some(&Color::rgb(180, 180, 180)),
            1.0,
        );
        return;
    };

    // Calculate bounds
    let x0 = (cx - rx).max(0.0) as u32;
    let y0 = (cy - ry).max(0.0) as u32;
    let x1 = ((cx + rx) as u32).min(img_width);
    let y1 = ((cy + ry) as u32).min(img_height);

    // Render image with elliptical clipping
    for py in y0..y1 {
        for px in x0..x1 {
            let dx = (px as f64 - cx) / rx;
            let dy = (py as f64 - cy) / ry;

            // Check if pixel is inside ellipse
            if dx * dx + dy * dy <= 1.0 {
                // Map pixel to source image coordinates
                let src_x = ((px as f64 - (cx - rx)) / (rx * 2.0) * decoded.width as f64) as u32;
                let src_y = ((py as f64 - (cy - ry)) / (ry * 2.0) * decoded.height as f64) as u32;

                if src_x < decoded.width && src_y < decoded.height {
                    let src_idx = ((src_y * decoded.width + src_x) * 4) as usize;
                    if src_idx + 3 < decoded.pixels.len() {
                        let src_a = decoded.pixels[src_idx + 3] as f64 / 255.0;

                        if src_a > 0.99 {
                            // Opaque pixel: direct copy
                            set_pixel(pixels, img_width, px, py, &Color {
                                r: decoded.pixels[src_idx],
                                g: decoded.pixels[src_idx + 1],
                                b: decoded.pixels[src_idx + 2],
                                a: 255,
                            });
                        } else if src_a > 0.01 {
                            // Semi-transparent: alpha blend
                            let dst_idx = ((py * img_width + px) * 4) as usize;
                            if dst_idx + 3 < pixels.len() {
                                pixels[dst_idx] = (pixels[dst_idx] as f64 * (1.0 - src_a)
                                    + decoded.pixels[src_idx] as f64 * src_a) as u8;
                                pixels[dst_idx + 1] = (pixels[dst_idx + 1] as f64 * (1.0 - src_a)
                                    + decoded.pixels[src_idx + 1] as f64 * src_a) as u8;
                                pixels[dst_idx + 2] = (pixels[dst_idx + 2] as f64 * (1.0 - src_a)
                                    + decoded.pixels[src_idx + 2] as f64 * src_a) as u8;
                            }
                        }
                    }
                }
            }
        }
    }
}

/// パスクリップされた画像をピクセルバッファに描画
fn render_path_image_to_pixels(
    pixels: &mut [u8],
    img_width: u32,
    img_height: u32,
    commands: &[crate::converter::PathCommand],
    data: &[u8],
    _mime_type: &str,
    scale: f64,
) {
    // Decode the image first
    let decoded = if let Some(img) = decode_png_image(data) {
        img
    } else if let Some(img) = decode_jpeg_image(data) {
        img
    } else {
        // Fallback: render placeholder
        return;
    };

    // Parse path commands into separate subpaths for clipping mask using helper function
    let subpaths = parse_path_commands_to_subpaths(commands, scale);

    if subpaths.is_empty() {
        return;
    }

    // Calculate bounding box of the path
    let mut min_x = f64::INFINITY;
    let mut min_y = f64::INFINITY;
    let mut max_x = f64::NEG_INFINITY;
    let mut max_y = f64::NEG_INFINITY;

    for subpath in &subpaths {
        for &(x, y) in subpath {
            min_x = min_x.min(x);
            min_y = min_y.min(y);
            max_x = max_x.max(x);
            max_y = max_y.max(y);
        }
    }

    let x0 = min_x.max(0.0) as u32;
    let y0 = min_y.max(0.0) as u32;
    let x1 = (max_x as u32).min(img_width);
    let y1 = (max_y as u32).min(img_height);
    let path_width = max_x - min_x;
    let path_height = max_y - min_y;

    if path_width <= 0.0 || path_height <= 0.0 {
        return;
    }

    // Render image with path clipping using scanline algorithm
    for py in y0..y1 {
        let y_f = py as f64 + 0.5;

        // Find all intersections with the path at this scanline
        let mut intersections: Vec<f64> = Vec::new();

        for subpath in &subpaths {
            if subpath.len() < 2 {
                continue;
            }

            for i in 0..subpath.len().saturating_sub(1) {
                let (x1, y1) = subpath[i];
                let (x2, y2) = subpath[i + 1];

                if (y1 <= y_f && y2 > y_f) || (y2 <= y_f && y1 > y_f) {
                    let t = (y_f - y1) / (y2 - y1);
                    intersections.push(x1 + t * (x2 - x1));
                }
            }
        }

        intersections.sort_by(|a, b| {
            // Use total ordering to ensure deterministic results
            a.partial_cmp(b).unwrap_or_else(|| {
                // Handle NaN cases: NaN values come last
                if a.is_nan() && b.is_nan() {
                    std::cmp::Ordering::Equal
                } else if a.is_nan() {
                    std::cmp::Ordering::Greater
                } else {
                    std::cmp::Ordering::Less
                }
            })
        });

        // Fill pixels between pairs of intersections
        for pair in intersections.chunks(2) {
            if pair.len() == 2 {
                let x_start = (pair[0].max(0.0) as u32).max(x0);
                let x_end = ((pair[1] as u32).min(img_width)).min(x1);

                for px in x_start..x_end {
                    // Map pixel to source image coordinates
                    let src_x = ((px as f64 - min_x) / path_width * decoded.width as f64) as u32;
                    let src_y = ((py as f64 - min_y) / path_height * decoded.height as f64) as u32;

                    if src_x < decoded.width && src_y < decoded.height {
                        let src_idx = ((src_y * decoded.width + src_x) * 4) as usize;
                        if src_idx + 3 < decoded.pixels.len() {
                            let src_a = decoded.pixels[src_idx + 3] as f64 / 255.0;

                            if src_a > 0.99 {
                                set_pixel(pixels, img_width, px, py, &Color {
                                    r: decoded.pixels[src_idx],
                                    g: decoded.pixels[src_idx + 1],
                                    b: decoded.pixels[src_idx + 2],
                                    a: 255,
                                });
                            } else if src_a > 0.01 {
                                let dst_idx = ((py * img_width + px) * 4) as usize;
                                if dst_idx + 3 < pixels.len() {
                                    pixels[dst_idx] = (pixels[dst_idx] as f64 * (1.0 - src_a)
                                        + decoded.pixels[src_idx] as f64 * src_a) as u8;
                                    pixels[dst_idx + 1] = (pixels[dst_idx + 1] as f64 * (1.0 - src_a)
                                        + decoded.pixels[src_idx + 1] as f64 * src_a) as u8;
                                    pixels[dst_idx + 2] = (pixels[dst_idx + 2] as f64 * (1.0 - src_a)
                                        + decoded.pixels[src_idx + 2] as f64 * src_a) as u8;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
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

/// JPEG画像をデコード（jpeg-decoderクレートによる完全実装）
fn decode_jpeg_image(data: &[u8]) -> Option<DecodedImage> {
    // Check JPEG magic bytes
    if data.len() < 4 || data[0] != 0xFF || data[1] != 0xD8 {
        return None;
    }

    let mut decoder = jpeg_decoder::Decoder::new(std::io::Cursor::new(data));
    let raw_pixels = decoder.decode().ok()?;
    let info = decoder.info()?;
    let width = info.width as u32;
    let height = info.height as u32;

    if width == 0 || height == 0 {
        return None;
    }

    // Convert to RGBA based on pixel format
    let rgba = match info.pixel_format {
        jpeg_decoder::PixelFormat::RGB24 => {
            let mut rgba = Vec::with_capacity((width * height * 4) as usize);
            for chunk in raw_pixels.chunks(3) {
                if chunk.len() >= 3 {
                    rgba.push(chunk[0]);
                    rgba.push(chunk[1]);
                    rgba.push(chunk[2]);
                    rgba.push(255);
                }
            }
            rgba
        }
        jpeg_decoder::PixelFormat::L8 => {
            let mut rgba = Vec::with_capacity((width * height * 4) as usize);
            for &g in &raw_pixels {
                rgba.push(g);
                rgba.push(g);
                rgba.push(g);
                rgba.push(255);
            }
            rgba
        }
        jpeg_decoder::PixelFormat::CMYK32 => {
            let mut rgba = Vec::with_capacity((width * height * 4) as usize);
            for chunk in raw_pixels.chunks(4) {
                if chunk.len() >= 4 {
                    // CMYK → RGB conversion
                    let c = chunk[0] as f64 / 255.0;
                    let m = chunk[1] as f64 / 255.0;
                    let y = chunk[2] as f64 / 255.0;
                    let k = chunk[3] as f64 / 255.0;
                    let r = (255.0 * (1.0 - c) * (1.0 - k)) as u8;
                    let g = (255.0 * (1.0 - m) * (1.0 - k)) as u8;
                    let b = (255.0 * (1.0 - y) * (1.0 - k)) as u8;
                    rgba.push(r);
                    rgba.push(g);
                    rgba.push(b);
                    rgba.push(255);
                }
            }
            rgba
        }
        _ => return None,
    };

    Some(DecodedImage { width, height, pixels: rgba })
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
