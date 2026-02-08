// formats/chart.rs - Officeチャート描画モジュール
//
// OOXML (c:chartSpace) で定義されたチャートを解析し、
// ドキュメントモデルの PageElement に変換して描画します。
// 棒グラフ、円グラフ、面グラフ等の主要チャートタイプに対応。

use crate::converter::{Color, PageElement, TextAlign, FontStyle};

/// チャートデータ系列
#[derive(Debug, Clone)]
struct ChartSeries {
    name: String,
    categories: Vec<String>,
    values: Vec<f64>,
    color: Color,
}

/// チャートタイプ
#[derive(Debug, Clone)]
enum ChartType {
    Bar { direction: BarDirection, grouping: String },
    Pie3D,
    Pie,
    Area,
    Line,
    Scatter,
}

#[derive(Debug, Clone)]
enum BarDirection {
    Column,  // 縦棒
    Bar,     // 横棒
}

/// チャート定義
#[derive(Debug, Clone)]
struct ChartDef {
    chart_type: ChartType,
    title: Option<String>,
    series: Vec<ChartSeries>,
}

/// テーマカラーパレット（アクセントカラー）
const CHART_COLORS: [Color; 6] = [
    Color { r: 91, g: 155, b: 213, a: 255 },   // accent1
    Color { r: 237, g: 125, b: 49, a: 255 },    // accent2
    Color { r: 165, g: 165, b: 165, a: 255 },   // accent3
    Color { r: 255, g: 192, b: 0, a: 255 },     // accent4
    Color { r: 68, g: 114, b: 196, a: 255 },    // accent5
    Color { r: 112, g: 173, b: 71, a: 255 },    // accent6
];

/// チャートXMLを解析してPageElementのリストを生成
pub fn render_chart(
    chart_xml: &str,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) -> Vec<PageElement> {
    let chart_def = parse_chart_xml(chart_xml);
    match chart_def {
        Some(def) => render_chart_def(&def, x, y, width, height),
        None => {
            // チャートを解析できない場合はプレースホルダーを描画
            vec![
                PageElement::Rect {
                    x, y, width, height,
                    fill: Some(Color::rgb(245, 245, 245)),
                    stroke: Some(Color::rgb(200, 200, 200)),
                    stroke_width: 1.0,
                },
                PageElement::Text {
                    x: x + 10.0,
                    y: y + height / 2.0 - 8.0,
                    width: width - 20.0,
                    text: "[Chart]".to_string(),
                    style: FontStyle { font_size: 12.0, ..FontStyle::default() },
                    align: TextAlign::Center,
                },
            ]
        }
    }
}

/// チャートXMLを解析
fn parse_chart_xml(xml: &str) -> Option<ChartDef> {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();

    let mut chart_type: Option<ChartType> = None;
    let mut series_list: Vec<ChartSeries> = Vec::new();
    let mut title: Option<String> = None;

    // Parsing state
    let mut in_chart_type = false;
    let mut current_chart_type_name = String::new();
    let mut in_ser = false;
    let mut in_tx = false;
    let mut in_cat = false;
    let mut in_val = false;
    let mut in_v = false;
    let mut in_title = false;
    let mut in_str_cache = false;
    let mut in_num_cache = false;

    let mut cur_ser_name = String::new();
    let mut cur_categories: Vec<String> = Vec::new();
    let mut cur_values: Vec<f64> = Vec::new();
    let mut ser_idx: usize = 0;

    // For bar chart direction
    let mut bar_dir = BarDirection::Column;
    let mut bar_grouping = "clustered".to_string();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");

                match name {
                    "barChart" => {
                        current_chart_type_name = "barChart".to_string();
                        in_chart_type = true;
                    }
                    "pie3DChart" => {
                        current_chart_type_name = "pie3DChart".to_string();
                        in_chart_type = true;
                    }
                    "pieChart" => {
                        current_chart_type_name = "pieChart".to_string();
                        in_chart_type = true;
                    }
                    "areaChart" => {
                        current_chart_type_name = "areaChart".to_string();
                        in_chart_type = true;
                    }
                    "lineChart" => {
                        current_chart_type_name = "lineChart".to_string();
                        in_chart_type = true;
                    }
                    "scatterChart" => {
                        current_chart_type_name = "scatterChart".to_string();
                        in_chart_type = true;
                    }
                    "barDir" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                let v = std::str::from_utf8(&attr.value).unwrap_or("col");
                                bar_dir = if v == "bar" { BarDirection::Bar } else { BarDirection::Column };
                            }
                        }
                    }
                    "grouping" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                bar_grouping = std::str::from_utf8(&attr.value).unwrap_or("clustered").to_string();
                            }
                        }
                    }
                    "ser" if in_chart_type => {
                        in_ser = true;
                        cur_ser_name = String::new();
                        cur_categories = Vec::new();
                        cur_values = Vec::new();
                    }
                    "tx" if in_ser => { in_tx = true; }
                    "cat" if in_ser => { in_cat = true; }
                    "val" if in_ser => { in_val = true; }
                    "strCache" => { in_str_cache = true; }
                    "numCache" => { in_num_cache = true; }
                    "v" => { in_v = true; }
                    "title" if !in_ser => { in_title = true; }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");

                match name {
                    "barChart" | "pie3DChart" | "pieChart" | "areaChart" | "lineChart" | "scatterChart" => {
                        in_chart_type = false;
                        chart_type = Some(match current_chart_type_name.as_str() {
                            "barChart" => ChartType::Bar { direction: bar_dir.clone(), grouping: bar_grouping.clone() },
                            "pie3DChart" => ChartType::Pie3D,
                            "pieChart" => ChartType::Pie,
                            "areaChart" => ChartType::Area,
                            "lineChart" => ChartType::Line,
                            "scatterChart" => ChartType::Scatter,
                            _ => ChartType::Bar { direction: BarDirection::Column, grouping: "clustered".to_string() },
                        });
                    }
                    "ser" if in_ser => {
                        let color = CHART_COLORS[ser_idx % CHART_COLORS.len()];
                        series_list.push(ChartSeries {
                            name: cur_ser_name.clone(),
                            categories: cur_categories.clone(),
                            values: cur_values.clone(),
                            color,
                        });
                        ser_idx += 1;
                        in_ser = false;
                    }
                    "tx" => { in_tx = false; }
                    "cat" => { in_cat = false; }
                    "val" => { in_val = false; }
                    "strCache" => { in_str_cache = false; }
                    "numCache" => { in_num_cache = false; }
                    "v" => { in_v = false; }
                    "title" => { in_title = false; }
                    _ => {}
                }
            }
            Ok(Event::Text(ref t)) if in_v => {
                let text = t.unescape().unwrap_or_default().to_string();
                if in_tx && in_str_cache {
                    cur_ser_name = text;
                } else if in_cat && in_str_cache {
                    cur_categories.push(text);
                } else if in_val && in_num_cache {
                    if let Ok(v) = text.parse::<f64>() {
                        cur_values.push(v);
                    }
                } else if in_title {
                    title = Some(text);
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");
                if name == "barDir" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"val" {
                            let v = std::str::from_utf8(&attr.value).unwrap_or("col");
                            bar_dir = if v == "bar" { BarDirection::Bar } else { BarDirection::Column };
                        }
                    }
                }
                if name == "grouping" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"val" {
                            bar_grouping = std::str::from_utf8(&attr.value).unwrap_or("clustered").to_string();
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    let ct = chart_type?;
    Some(ChartDef {
        chart_type: ct,
        title,
        series: series_list,
    })
}

/// チャート定義からPageElementリストを生成
fn render_chart_def(
    def: &ChartDef,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) -> Vec<PageElement> {
    let mut elements = Vec::new();

    // 背景
    elements.push(PageElement::Rect {
        x, y, width, height,
        fill: Some(Color::WHITE),
        stroke: Some(Color::rgb(200, 200, 200)),
        stroke_width: 0.5,
    });

    let margin = 20.0;
    let title_height = if def.title.is_some() { 25.0 } else { 0.0 };
    let legend_height = 20.0;

    let plot_x = x + margin + 30.0; // extra space for y-axis labels
    let plot_y = y + margin + title_height;
    let plot_w = width - margin * 2.0 - 30.0;
    let plot_h = height - margin * 2.0 - title_height - legend_height;

    // タイトル
    if let Some(ref title_text) = def.title {
        elements.push(PageElement::Text {
            x: x + width / 2.0 - 50.0,
            y: y + 8.0,
            width: 100.0,
            text: title_text.clone(),
            style: FontStyle {
                font_size: 11.0,
                bold: true,
                ..FontStyle::default()
            },
            align: TextAlign::Center,
        });
    }

    match &def.chart_type {
        ChartType::Bar { direction, grouping } => {
            render_bar_chart(&mut elements, &def.series, plot_x, plot_y, plot_w, plot_h, direction, grouping);
        }
        ChartType::Pie3D | ChartType::Pie => {
            render_pie_chart(&mut elements, &def.series, plot_x, plot_y, plot_w, plot_h, matches!(def.chart_type, ChartType::Pie3D));
        }
        ChartType::Area => {
            render_area_chart(&mut elements, &def.series, plot_x, plot_y, plot_w, plot_h);
        }
        ChartType::Line => {
            render_line_chart(&mut elements, &def.series, plot_x, plot_y, plot_w, plot_h);
        }
        ChartType::Scatter => {
            render_line_chart(&mut elements, &def.series, plot_x, plot_y, plot_w, plot_h);
        }
    }

    // 凡例
    render_legend(&mut elements, &def.series, x + margin, y + height - legend_height, plot_w);

    elements
}

/// 棒グラフの描画
fn render_bar_chart(
    elements: &mut Vec<PageElement>,
    series: &[ChartSeries],
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    direction: &BarDirection,
    _grouping: &str,
) {
    if series.is_empty() {
        return;
    }

    // Find value range
    let max_val = series.iter()
        .flat_map(|s| s.values.iter())
        .cloned()
        .fold(0.0f64, f64::max)
        .max(0.001);

    let num_categories = series[0].categories.len().max(1);
    let num_series = series.len().max(1);

    // Draw axes
    elements.push(PageElement::Line {
        x1: x, y1: y + h, x2: x + w, y2: y + h,
        width: 1.0, color: Color::rgb(128, 128, 128),
    });
    elements.push(PageElement::Line {
        x1: x, y1: y, x2: x, y2: y + h,
        width: 1.0, color: Color::rgb(128, 128, 128),
    });

    // Draw gridlines
    let num_gridlines = 5;
    for i in 1..=num_gridlines {
        let gy = y + h - (h * i as f64 / num_gridlines as f64);
        elements.push(PageElement::Line {
            x1: x, y1: gy, x2: x + w, y2: gy,
            width: 0.3, color: Color::rgb(220, 220, 220),
        });
        // Y-axis label
        let label_val = max_val * i as f64 / num_gridlines as f64;
        elements.push(PageElement::Text {
            x: x - 28.0, y: gy - 4.0, width: 26.0,
            text: format!("{:.1}", label_val),
            style: FontStyle { font_size: 7.0, color: Color::rgb(100, 100, 100), ..FontStyle::default() },
            align: TextAlign::Right,
        });
    }

    match direction {
        BarDirection::Column => {
            let category_width = w / num_categories as f64;
            let bar_width = category_width * 0.7 / num_series as f64;
            let bar_gap = category_width * 0.15;

            for (ci, _) in (0..num_categories).enumerate() {
                // Category label
                let cat_label = series[0].categories.get(ci).cloned().unwrap_or_default();
                elements.push(PageElement::Text {
                    x: x + ci as f64 * category_width + 2.0,
                    y: y + h + 3.0,
                    width: category_width - 4.0,
                    text: cat_label,
                    style: FontStyle { font_size: 7.0, color: Color::rgb(100, 100, 100), ..FontStyle::default() },
                    align: TextAlign::Center,
                });

                for (si, ser) in series.iter().enumerate() {
                    let val = ser.values.get(ci).copied().unwrap_or(0.0);
                    let bar_h = (val / max_val) * h;
                    let bx = x + ci as f64 * category_width + bar_gap + si as f64 * bar_width;
                    let by = y + h - bar_h;

                    elements.push(PageElement::Rect {
                        x: bx, y: by, width: bar_width, height: bar_h,
                        fill: Some(ser.color),
                        stroke: None,
                        stroke_width: 0.0,
                    });
                }
            }
        }
        BarDirection::Bar => {
            let category_height = h / num_categories as f64;
            let bar_height = category_height * 0.7 / num_series as f64;
            let bar_gap = category_height * 0.15;

            for (ci, _) in (0..num_categories).enumerate() {
                for (si, ser) in series.iter().enumerate() {
                    let val = ser.values.get(ci).copied().unwrap_or(0.0);
                    let bar_w = (val / max_val) * w;
                    let by = y + ci as f64 * category_height + bar_gap + si as f64 * bar_height;

                    elements.push(PageElement::Rect {
                        x, y: by, width: bar_w, height: bar_height,
                        fill: Some(ser.color),
                        stroke: None,
                        stroke_width: 0.0,
                    });
                }
            }
        }
    }
}

/// 円グラフの描画
fn render_pie_chart(
    elements: &mut Vec<PageElement>,
    series: &[ChartSeries],
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    is_3d: bool,
) {
    if series.is_empty() || series[0].values.is_empty() {
        return;
    }

    let values = &series[0].values;
    let categories = &series[0].categories;
    let total: f64 = values.iter().sum();
    if total <= 0.0 {
        return;
    }

    let cx = x + w / 2.0;
    let cy = y + h / 2.0;
    let rx = (w.min(h) / 2.0) * 0.85;
    let ry = if is_3d { rx * 0.6 } else { rx }; // 3D effect: compress vertically

    // For 3D pie, draw the "side" first
    if is_3d {
        let depth = 12.0;
        // Draw side of pie (simplified: full ellipse side)
        for d in (0..depth as u32).rev() {
            let side_cy = cy + d as f64;
            elements.push(PageElement::Ellipse {
                cx, cy: side_cy, rx, ry,
                fill: Some(Color::rgb(150, 150, 150)),
                stroke: None,
                stroke_width: 0.0,
            });
        }
    }

    // Draw pie slices as colored ellipse segments
    // Since we can't easily do arc segments with our current primitives,
    // we approximate each slice as a filled triangle fan from center
    let num_points = 360;
    let mut start_angle = -std::f64::consts::FRAC_PI_2; // Start from top

    for (i, &val) in values.iter().enumerate() {
        let sweep = (val / total) * 2.0 * std::f64::consts::PI;
        let end_angle = start_angle + sweep;
        let color = CHART_COLORS[i % CHART_COLORS.len()];

        // For each slice, draw a filled polygon approximation using rectangles
        let mid_angle = start_angle + sweep / 2.0;
        let points_in_slice = ((sweep / (2.0 * std::f64::consts::PI)) * num_points as f64) as usize;

        if points_in_slice > 0 {
            // Draw slice using small triangular-ish rects from center
            let step = sweep / points_in_slice as f64;
            for j in 0..points_in_slice {
                let a1 = start_angle + j as f64 * step;
                let a2 = a1 + step;

                let x1 = cx + rx * a1.cos();
                let y1 = cy + ry * a1.sin();
                let x2 = cx + rx * a2.cos();
                let y2 = cy + ry * a2.sin();

                // Draw a line from center to edge for each point
                elements.push(PageElement::Line {
                    x1: cx, y1: cy, x2: x1, y2: y1,
                    width: 2.0, color,
                });
                elements.push(PageElement::Line {
                    x1: x1, y1: y1, x2: x2, y2: y2,
                    width: 2.0, color,
                });
            }

            // Fill the sector with small rectangles (radial fill approximation)
            let fill_steps = (rx.max(ry) * 0.5) as usize;
            for r_step in 0..fill_steps {
                let r_frac = r_step as f64 / fill_steps as f64;
                let curr_rx = rx * r_frac;
                let curr_ry = ry * r_frac;

                let angle_steps = (points_in_slice * 2).max(4);
                let a_step = sweep / angle_steps as f64;
                for a_idx in 0..angle_steps {
                    let a = start_angle + a_idx as f64 * a_step;
                    let px = cx + curr_rx * a.cos();
                    let py = cy + curr_ry * a.sin();
                    // Set pixel-like small rect
                    elements.push(PageElement::Rect {
                        x: px - 1.0, y: py - 1.0, width: 2.5, height: 2.5,
                        fill: Some(color),
                        stroke: None,
                        stroke_width: 0.0,
                    });
                }
            }
        }

        // Label
        let label_dist = rx * 0.65;
        let label_x = cx + label_dist * mid_angle.cos();
        let label_y = cy + (ry * 0.65) * mid_angle.sin();
        let pct = (val / total * 100.0) as u32;
        let label_text = if let Some(cat) = categories.get(i) {
            format!("{} ({}%)", cat, pct)
        } else {
            format!("{}%", pct)
        };

        elements.push(PageElement::Text {
            x: label_x - 25.0,
            y: label_y - 5.0,
            width: 50.0,
            text: label_text,
            style: FontStyle { font_size: 6.0, ..FontStyle::default() },
            align: TextAlign::Center,
        });

        start_angle = end_angle;
    }

    // Outline ellipse
    elements.push(PageElement::Ellipse {
        cx, cy, rx, ry,
        fill: None,
        stroke: Some(Color::rgb(100, 100, 100)),
        stroke_width: 0.5,
    });
}

/// 面グラフの描画
fn render_area_chart(
    elements: &mut Vec<PageElement>,
    series: &[ChartSeries],
    x: f64,
    y: f64,
    w: f64,
    h: f64,
) {
    if series.is_empty() {
        return;
    }

    let max_val = series.iter()
        .flat_map(|s| s.values.iter())
        .cloned()
        .fold(0.0f64, f64::max)
        .max(0.001);

    let num_categories = series.iter().map(|s| s.values.len()).max().unwrap_or(1).max(1);

    // Draw axes
    elements.push(PageElement::Line {
        x1: x, y1: y + h, x2: x + w, y2: y + h,
        width: 1.0, color: Color::rgb(128, 128, 128),
    });
    elements.push(PageElement::Line {
        x1: x, y1: y, x2: x, y2: y + h,
        width: 1.0, color: Color::rgb(128, 128, 128),
    });

    // Draw gridlines
    let num_gridlines = 5;
    for i in 1..=num_gridlines {
        let gy = y + h - (h * i as f64 / num_gridlines as f64);
        elements.push(PageElement::Line {
            x1: x, y1: gy, x2: x + w, y2: gy,
            width: 0.3, color: Color::rgb(220, 220, 220),
        });
    }

    // Draw each series as filled area
    for ser in series.iter().rev() {
        let color_fill = Color {
            r: ser.color.r,
            g: ser.color.g,
            b: ser.color.b,
            a: 128, // semi-transparent
        };

        let step = w / (num_categories - 1).max(1) as f64;

        // Fill area with vertical strips
        for i in 0..ser.values.len().saturating_sub(1) {
            let v1 = ser.values[i];
            let v2 = ser.values[i + 1];
            let h1 = (v1 / max_val) * h;
            let h2 = (v2 / max_val) * h;

            let x1 = x + i as f64 * step;
            let x2 = x + (i + 1) as f64 * step;
            let strip_w = x2 - x1;
            let avg_h = (h1 + h2) / 2.0;

            elements.push(PageElement::Rect {
                x: x1,
                y: y + h - avg_h,
                width: strip_w,
                height: avg_h,
                fill: Some(color_fill),
                stroke: None,
                stroke_width: 0.0,
            });
        }

        // Draw the line on top
        for i in 0..ser.values.len().saturating_sub(1) {
            let v1 = ser.values[i];
            let v2 = ser.values[i + 1];
            let px1 = x + i as f64 * step;
            let py1 = y + h - (v1 / max_val) * h;
            let px2 = x + (i + 1) as f64 * step;
            let py2 = y + h - (v2 / max_val) * h;

            elements.push(PageElement::Line {
                x1: px1, y1: py1, x2: px2, y2: py2,
                width: 1.5, color: ser.color,
            });
        }
    }

    // Category labels
    if !series.is_empty() {
        let step = w / (num_categories - 1).max(1) as f64;
        for (i, cat) in series[0].categories.iter().enumerate() {
            elements.push(PageElement::Text {
                x: x + i as f64 * step - 15.0,
                y: y + h + 3.0,
                width: 30.0,
                text: cat.clone(),
                style: FontStyle { font_size: 7.0, color: Color::rgb(100, 100, 100), ..FontStyle::default() },
                align: TextAlign::Center,
            });
        }
    }
}

/// 折れ線グラフの描画
fn render_line_chart(
    elements: &mut Vec<PageElement>,
    series: &[ChartSeries],
    x: f64,
    y: f64,
    w: f64,
    h: f64,
) {
    if series.is_empty() {
        return;
    }

    let max_val = series.iter()
        .flat_map(|s| s.values.iter())
        .cloned()
        .fold(0.0f64, f64::max)
        .max(0.001);

    let num_categories = series.iter().map(|s| s.values.len()).max().unwrap_or(1).max(1);

    // Draw axes
    elements.push(PageElement::Line {
        x1: x, y1: y + h, x2: x + w, y2: y + h,
        width: 1.0, color: Color::rgb(128, 128, 128),
    });
    elements.push(PageElement::Line {
        x1: x, y1: y, x2: x, y2: y + h,
        width: 1.0, color: Color::rgb(128, 128, 128),
    });

    // Draw gridlines
    let num_gridlines = 5;
    for i in 1..=num_gridlines {
        let gy = y + h - (h * i as f64 / num_gridlines as f64);
        elements.push(PageElement::Line {
            x1: x, y1: gy, x2: x + w, y2: gy,
            width: 0.3, color: Color::rgb(220, 220, 220),
        });
    }

    let step = w / (num_categories - 1).max(1) as f64;

    for ser in series {
        // Draw lines
        for i in 0..ser.values.len().saturating_sub(1) {
            let v1 = ser.values[i];
            let v2 = ser.values[i + 1];
            let px1 = x + i as f64 * step;
            let py1 = y + h - (v1 / max_val) * h;
            let px2 = x + (i + 1) as f64 * step;
            let py2 = y + h - (v2 / max_val) * h;

            elements.push(PageElement::Line {
                x1: px1, y1: py1, x2: px2, y2: py2,
                width: 2.0, color: ser.color,
            });
        }

        // Draw data points
        for (i, &val) in ser.values.iter().enumerate() {
            let px = x + i as f64 * step;
            let py = y + h - (val / max_val) * h;
            elements.push(PageElement::Ellipse {
                cx: px, cy: py, rx: 3.0, ry: 3.0,
                fill: Some(ser.color),
                stroke: Some(Color::WHITE),
                stroke_width: 1.0,
            });
        }
    }
}

/// 凡例の描画
fn render_legend(
    elements: &mut Vec<PageElement>,
    series: &[ChartSeries],
    x: f64,
    y: f64,
    _width: f64,
) {
    let mut lx = x;
    for ser in series {
        // Color swatch
        elements.push(PageElement::Rect {
            x: lx, y: y + 2.0, width: 10.0, height: 10.0,
            fill: Some(ser.color),
            stroke: None,
            stroke_width: 0.0,
        });
        // Label
        elements.push(PageElement::Text {
            x: lx + 12.0, y: y + 2.0, width: 60.0,
            text: ser.name.clone(),
            style: FontStyle { font_size: 7.0, color: Color::rgb(80, 80, 80), ..FontStyle::default() },
            align: TextAlign::Left,
        });
        lx += 80.0;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_bar_chart() {
        let xml = r#"<?xml version="1.0"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:chart>
                <c:plotArea>
                    <c:barChart>
                        <c:barDir val="col"/>
                        <c:grouping val="clustered"/>
                        <c:ser>
                            <c:idx val="0"/>
                            <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series 1</c:v></c:pt></c:strCache></c:strRef></c:tx>
                            <c:cat><c:strRef><c:strCache>
                                <c:pt idx="0"><c:v>A</c:v></c:pt>
                                <c:pt idx="1"><c:v>B</c:v></c:pt>
                            </c:strCache></c:strRef></c:cat>
                            <c:val><c:numRef><c:numCache>
                                <c:pt idx="0"><c:v>4.3</c:v></c:pt>
                                <c:pt idx="1"><c:v>2.5</c:v></c:pt>
                            </c:numCache></c:numRef></c:val>
                        </c:ser>
                    </c:barChart>
                </c:plotArea>
            </c:chart>
        </c:chartSpace>"#;

        let def = parse_chart_xml(xml);
        assert!(def.is_some());
        let def = def.unwrap();
        assert!(matches!(def.chart_type, ChartType::Bar { .. }));
        assert_eq!(def.series.len(), 1);
        assert_eq!(def.series[0].name, "Series 1");
        assert_eq!(def.series[0].values, vec![4.3, 2.5]);
    }

    #[test]
    fn test_render_bar_chart_elements() {
        let elements = render_chart(
            r#"<?xml version="1.0"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart><c:plotArea>
                    <c:barChart>
                        <c:barDir val="col"/>
                        <c:ser>
                            <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>S1</c:v></c:pt></c:strCache></c:strRef></c:tx>
                            <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Cat1</c:v></c:pt></c:strCache></c:strRef></c:cat>
                            <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:val>
                        </c:ser>
                    </c:barChart>
                </c:plotArea></c:chart>
            </c:chartSpace>"#,
            0.0, 0.0, 400.0, 300.0,
        );
        assert!(!elements.is_empty());
        // Should have background rect, axes, gridlines, bars, labels, legend
        assert!(elements.len() > 5);
    }

    #[test]
    fn test_parse_pie_chart() {
        let xml = r#"<?xml version="1.0"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea>
                <c:pie3DChart>
                    <c:ser>
                        <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>
                        <c:cat><c:strRef><c:strCache>
                            <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                            <c:pt idx="1"><c:v>Q2</c:v></c:pt>
                        </c:strCache></c:strRef></c:cat>
                        <c:val><c:numRef><c:numCache>
                            <c:pt idx="0"><c:v>8.2</c:v></c:pt>
                            <c:pt idx="1"><c:v>3.2</c:v></c:pt>
                        </c:numCache></c:numRef></c:val>
                    </c:ser>
                </c:pie3DChart>
            </c:plotArea></c:chart>
        </c:chartSpace>"#;

        let def = parse_chart_xml(xml);
        assert!(def.is_some());
        let def = def.unwrap();
        assert!(matches!(def.chart_type, ChartType::Pie3D));
    }

    #[test]
    fn test_parse_area_chart() {
        let xml = r#"<?xml version="1.0"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea>
                <c:areaChart>
                    <c:ser>
                        <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Data</c:v></c:pt></c:strCache></c:strRef></c:tx>
                        <c:val><c:numRef><c:numCache>
                            <c:pt idx="0"><c:v>1</c:v></c:pt>
                            <c:pt idx="1"><c:v>3</c:v></c:pt>
                            <c:pt idx="2"><c:v>2</c:v></c:pt>
                        </c:numCache></c:numRef></c:val>
                    </c:ser>
                </c:areaChart>
            </c:plotArea></c:chart>
        </c:chartSpace>"#;

        let def = parse_chart_xml(xml);
        assert!(def.is_some());
        let def = def.unwrap();
        assert!(matches!(def.chart_type, ChartType::Area));
    }
}
