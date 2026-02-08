// formats/smartart.rs - SmartArt/ダイアグラム描画モジュール
//
// OOXML dsp:drawing で定義されたSmartArtダイアグラムを解析し、
// ドキュメントモデルの PageElement に変換して描画します。

use crate::converter::{Color, FontStyle, PageElement, TextAlign};

/// SmartArtシェイプ
#[derive(Debug, Clone)]
struct SmartArtShape {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
    text: String,
    fill_color: Option<Color>,
    is_rounded: bool,
}

/// SmartArtテキストフォントサイズのシェイプ高さに対する比率
const FONT_SIZE_RATIO: f64 = 0.25;
/// SmartArtテキストフォントサイズ上限
const MAX_FONT_SIZE: f64 = 12.0;
/// SmartArtテキストフォントサイズ下限
const MIN_FONT_SIZE: f64 = 6.0;

/// SmartArt XMLを解析してPageElementのリストを生成
pub fn render_smartart(
    drawing_xml: &str,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) -> Vec<PageElement> {
    let shapes = parse_smartart_shapes(drawing_xml);
    if shapes.is_empty() {
        return render_placeholder(x, y, width, height, "[SmartArt]");
    }

    let mut elements = Vec::new();

    // Calculate bounding box of all shapes to scale to target area
    let min_x = shapes.iter().map(|s| s.x).fold(f64::MAX, f64::min);
    let min_y = shapes.iter().map(|s| s.y).fold(f64::MAX, f64::min);
    let max_x = shapes.iter().map(|s| s.x + s.width).fold(f64::MIN, f64::max);
    let max_y = shapes.iter().map(|s| s.y + s.height).fold(f64::MIN, f64::max);

    let src_w = (max_x - min_x).max(1.0);
    let src_h = (max_y - min_y).max(1.0);

    let margin = 8.0;
    let target_w = width - margin * 2.0;
    let target_h = height - margin * 2.0;

    let scale_x = target_w / src_w;
    let scale_y = target_h / src_h;
    let scale = scale_x.min(scale_y); // Uniform scaling

    let offset_x = x + margin + (target_w - src_w * scale) / 2.0;
    let offset_y = y + margin + (target_h - src_h * scale) / 2.0;

    for shape in &shapes {
        let sx = offset_x + (shape.x - min_x) * scale;
        let sy = offset_y + (shape.y - min_y) * scale;
        let sw = shape.width * scale;
        let sh = shape.height * scale;

        let fill = shape.fill_color.unwrap_or(Color::rgb(91, 155, 213));

        if shape.is_rounded {
            // Render rounded rect as ellipse
            elements.push(PageElement::Ellipse {
                cx: sx + sw / 2.0,
                cy: sy + sh / 2.0,
                rx: sw / 2.0,
                ry: sh / 2.0,
                fill: Some(fill),
                stroke: Some(Color::WHITE),
                stroke_width: 1.0,
            });
        } else {
            elements.push(PageElement::Rect {
                x: sx, y: sy, width: sw, height: sh,
                fill: Some(fill),
                stroke: Some(Color::WHITE),
                stroke_width: 1.0,
            });
        }

        // Text
        if !shape.text.is_empty() {
            let font_size = (sh * FONT_SIZE_RATIO).min(MAX_FONT_SIZE).max(MIN_FONT_SIZE);
            elements.push(PageElement::Text {
                x: sx + 4.0,
                y: sy + sh / 2.0 - font_size / 2.0,
                width: sw - 8.0,
                text: shape.text.clone(),
                style: FontStyle {
                    font_size,
                    color: Color::WHITE,
                    bold: true,
                    ..FontStyle::default()
                },
                align: TextAlign::Center,
            });
        }
    }

    elements
}

/// SmartArt XMLからシェイプを解析
fn parse_smartart_shapes(xml: &str) -> Vec<SmartArtShape> {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    let mut shapes = Vec::new();
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();

    // EMU to pt conversion
    const EMU_TO_PT: f64 = 72.0 / 914400.0;

    let mut in_sp = false;
    let mut in_xfrm = false;
    let mut in_text = false;
    let mut in_solid_fill = false;
    let mut is_rounded = false;

    let mut cur_x = 0.0f64;
    let mut cur_y = 0.0f64;
    let mut cur_w = 0.0f64;
    let mut cur_h = 0.0f64;
    let mut cur_text = String::new();
    let mut cur_fill: Option<Color> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");

                match name {
                    "sp" => {
                        in_sp = true;
                        cur_x = 0.0;
                        cur_y = 0.0;
                        cur_w = 0.0;
                        cur_h = 0.0;
                        cur_text = String::new();
                        cur_fill = None;
                        is_rounded = false;
                    }
                    "xfrm" if in_sp => { in_xfrm = true; }
                    "t" if in_sp => { in_text = true; }
                    "solidFill" if in_sp => { in_solid_fill = true; }
                    "prstGeom" if in_sp => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"prst" {
                                let v = std::str::from_utf8(&attr.value).unwrap_or("");
                                if v.contains("round") || v == "ellipse" {
                                    is_rounded = true;
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");

                match name {
                    "off" if in_xfrm => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = std::str::from_utf8(&attr.value).unwrap_or("0");
                            match key {
                                "x" => cur_x = val.parse::<f64>().unwrap_or(0.0) * EMU_TO_PT,
                                "y" => cur_y = val.parse::<f64>().unwrap_or(0.0) * EMU_TO_PT,
                                _ => {}
                            }
                        }
                    }
                    "ext" if in_xfrm => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = std::str::from_utf8(&attr.value).unwrap_or("0");
                            match key {
                                "cx" => cur_w = val.parse::<f64>().unwrap_or(0.0) * EMU_TO_PT,
                                "cy" => cur_h = val.parse::<f64>().unwrap_or(0.0) * EMU_TO_PT,
                                _ => {}
                            }
                        }
                    }
                    "schemeClr" if in_solid_fill => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                let scheme = std::str::from_utf8(&attr.value).unwrap_or("");
                                cur_fill = Some(resolve_scheme_color(scheme));
                            }
                        }
                    }
                    "srgbClr" if in_solid_fill => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                let hex = std::str::from_utf8(&attr.value).unwrap_or("");
                                cur_fill = parse_hex_color(hex);
                            }
                        }
                    }
                    "prstGeom" if in_sp => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"prst" {
                                let v = std::str::from_utf8(&attr.value).unwrap_or("");
                                if v.contains("round") || v == "ellipse" {
                                    is_rounded = true;
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref t)) if in_text => {
                if let Ok(text) = t.unescape() {
                    if !cur_text.is_empty() {
                        cur_text.push(' ');
                    }
                    cur_text.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");

                match name {
                    "sp" => {
                        if in_sp && cur_w > 0.0 && cur_h > 0.0 {
                            shapes.push(SmartArtShape {
                                x: cur_x,
                                y: cur_y,
                                width: cur_w,
                                height: cur_h,
                                text: cur_text.clone(),
                                fill_color: cur_fill,
                                is_rounded,
                            });
                        }
                        in_sp = false;
                    }
                    "xfrm" => { in_xfrm = false; }
                    "t" => { in_text = false; }
                    "solidFill" => { in_solid_fill = false; }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    shapes
}

fn resolve_scheme_color(scheme: &str) -> Color {
    match scheme {
        "accent1" => Color::rgb(91, 155, 213),
        "accent2" => Color::rgb(237, 125, 49),
        "accent3" => Color::rgb(165, 165, 165),
        "accent4" => Color::rgb(255, 192, 0),
        "accent5" => Color::rgb(68, 114, 196),
        "accent6" => Color::rgb(112, 173, 71),
        "dk1" | "tx1" => Color::rgb(0, 0, 0),
        "lt1" | "bg1" => Color::rgb(255, 255, 255),
        _ => Color::rgb(91, 155, 213),
    }
}

fn parse_hex_color(hex: &str) -> Option<Color> {
    if hex.len() >= 6 {
        let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
        let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
        let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
        Some(Color::rgb(r, g, b))
    } else {
        None
    }
}

fn render_placeholder(x: f64, y: f64, width: f64, height: f64, label: &str) -> Vec<PageElement> {
    vec![
        PageElement::Rect {
            x, y, width, height,
            fill: Some(Color::rgb(240, 240, 240)),
            stroke: Some(Color::rgb(200, 200, 200)),
            stroke_width: 1.0,
        },
        PageElement::Text {
            x: x + 10.0,
            y: y + height / 2.0 - 8.0,
            width: width - 20.0,
            text: label.to_string(),
            style: FontStyle { font_size: 12.0, ..FontStyle::default() },
            align: TextAlign::Center,
        },
    ]
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_smartart_shapes() {
        let xml = r#"<?xml version="1.0"?>
        <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <dsp:spTree>
                <dsp:sp>
                    <dsp:spPr>
                        <a:xfrm><a:off x="914400" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="roundRect"/>
                        <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
                    </dsp:spPr>
                    <dsp:txBody>
                        <a:p><a:r><a:t>ABC</a:t></a:r></a:p>
                    </dsp:txBody>
                </dsp:sp>
                <dsp:sp>
                    <dsp:spPr>
                        <a:xfrm><a:off x="0" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:solidFill><a:schemeClr val="accent2"/></a:solidFill>
                    </dsp:spPr>
                    <dsp:txBody>
                        <a:p><a:r><a:t>DEF</a:t></a:r></a:p>
                    </dsp:txBody>
                </dsp:sp>
            </dsp:spTree>
        </dsp:drawing>"#;

        let shapes = parse_smartart_shapes(xml);
        assert_eq!(shapes.len(), 2);
        assert_eq!(shapes[0].text, "ABC");
        assert!(shapes[0].is_rounded);
        assert_eq!(shapes[1].text, "DEF");
        assert!(!shapes[1].is_rounded);
    }

    #[test]
    fn test_render_smartart() {
        let xml = r#"<?xml version="1.0"?>
        <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <dsp:spTree>
                <dsp:sp>
                    <dsp:spPr>
                        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                    </dsp:spPr>
                    <dsp:txBody><a:p><a:r><a:t>Test</a:t></a:r></a:p></dsp:txBody>
                </dsp:sp>
            </dsp:spTree>
        </dsp:drawing>"#;

        let elements = render_smartart(xml, 0.0, 0.0, 400.0, 300.0);
        assert!(!elements.is_empty());
    }
}
