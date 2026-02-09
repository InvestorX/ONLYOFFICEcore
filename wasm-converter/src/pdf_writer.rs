// pdf_writer.rs - 軽量PDF生成エンジン
//
// 外部クレートに依存せず、PDF 1.4仕様に準拠したPDFバイト列を直接生成します。
// 日本語テキスト（Unicode）をサポートします。

use crate::converter::{Color, Document, FontStyle, GradientStop, GradientType, Page, PageElement, Table, TextAlign};
use crate::font_manager::FontManager;

/// ページ内の画像XObject情報
struct PdfImageXObject {
    name: String,       // e.g. "Im0"
    obj_id: u32,        // PDF object ID for the image XObject
    #[allow(dead_code)]
    smask_id: Option<u32>, // Optional SMask object ID for alpha
}

/// PDFオブジェクト
struct PdfObject {
    id: u32,
    data: Vec<u8>,
}

/// PDF生成器
/// フォントマネージャーを参照として保持し、外部フォントを含むすべてのフォントを利用可能にします。
pub struct PdfWriter<'a> {
    objects: Vec<PdfObject>,
    next_id: u32,
    page_ids: Vec<u32>,
    font_manager: &'a FontManager,
}

impl<'a> PdfWriter<'a> {
    pub fn new(font_manager: &'a FontManager) -> Self {
        Self {
            objects: Vec::new(),
            next_id: 1,
            page_ids: Vec::new(),
            font_manager,
        }
    }

    fn alloc_id(&mut self) -> u32 {
        let id = self.next_id;
        self.next_id += 1;
        id
    }

    fn add_object(&mut self, id: u32, data: Vec<u8>) {
        self.objects.push(PdfObject { id, data });
    }

    /// ドキュメントをPDFバイト列に変換
    pub fn render(&mut self, doc: &Document) -> Vec<u8> {
        // ab_glyphでパース可能なフォントデータを見つけてコピー
        // CIDToGIDMapとフォント埋め込みの両方で同じフォントを使用することを保証
        let usable_font_data: Option<Vec<u8>> = find_usable_font_data(self.font_manager);
        let has_font = usable_font_data.is_some();

        // IDを事前割り当て
        let catalog_id = self.alloc_id();
        let pages_id = self.alloc_id();
        let font_id = self.alloc_id();

        // Fallback standard font (Helvetica) for when no embedded font available
        let fallback_font_id = self.alloc_id();

        let cid_font_id = self.alloc_id();
        let descriptor_id = self.alloc_id();
        let tounicode_id = self.alloc_id();

        // CIDToGIDMap用のID
        let cid_to_gid_map_id = self.alloc_id();

        // フォント埋め込み用のID（パース可能なフォントデータがある場合のみ）
        let font_file_id = if usable_font_data.is_some() {
            Some(self.alloc_id())
        } else {
            None
        };

        // ページオブジェクトのID割り当て
        let mut page_content_pairs: Vec<(u32, u32)> = Vec::new();
        for _ in &doc.pages {
            let page_id = self.alloc_id();
            let content_id = self.alloc_id();
            page_content_pairs.push((page_id, content_id));
        }

        // カタログ
        self.add_object(
            catalog_id,
            format!(
                "<< /Type /Catalog /Pages {} 0 R >>",
                pages_id
            )
            .into_bytes(),
        );

        // ページツリー
        let page_refs: Vec<String> = page_content_pairs
            .iter()
            .map(|(pid, _)| format!("{} 0 R", pid))
            .collect();
        self.add_object(
            pages_id,
            format!(
                "<< /Type /Pages /Kids [{}] /Count {} >>",
                page_refs.join(" "),
                doc.pages.len()
            )
            .into_bytes(),
        );

        // ToUnicode CMap（日本語テキスト用）
        let tounicode_stream = self.create_tounicode_cmap();
        let tounicode_compressed = tounicode_stream.as_bytes().to_vec();
        self.add_object(
            tounicode_id,
            format!(
                "<< /Length {} >>\nstream\n{}\nendstream",
                tounicode_compressed.len(),
                tounicode_stream
            )
            .into_bytes(),
        );

        // CIDToGIDMapストリームを生成（フォントのcmapテーブルに基づく）
        // usable_font_dataと同じデータを使用してマッピングの一致を保証
        {
            let cid_to_gid_data = Self::build_cid_to_gid_map_from_data(usable_font_data.as_deref());
            let mut stream_data = format!(
                "<< /Length {} >>\nstream\n",
                cid_to_gid_data.len()
            ).into_bytes();
            stream_data.extend_from_slice(&cid_to_gid_data);
            stream_data.extend_from_slice(b"\nendstream");
            self.add_object(cid_to_gid_map_id, stream_data);
        }

        // CIDFont
        let mut cid_font_dict = format!(
            "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /NotoSansJP \
             /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> \
             /DW 1000 \
             /FontDescriptor {} 0 R \
             /CIDToGIDMap {} 0 R",
            descriptor_id, cid_to_gid_map_id
        );
        cid_font_dict.push_str(" >>");
        self.add_object(cid_font_id, cid_font_dict.into_bytes());

        // フォントディスクリプタ
        let mut desc = format!(
            "<< /Type /FontDescriptor /FontName /NotoSansJP \
             /Flags 4 /ItalicAngle 0 /Ascent 880 /Descent -120 \
             /CapHeight 733 /StemV 80 \
             /FontBBox [-200 -200 1200 1000]"
        );
        if let Some(ff_id) = font_file_id {
            desc.push_str(&format!(" /FontFile2 {} 0 R", ff_id));
        }
        desc.push_str(" >>");
        self.add_object(descriptor_id, desc.into_bytes());

        // Type0フォント（複合フォント）- /F1
        self.add_object(
            font_id,
            format!(
                "<< /Type /Font /Subtype /Type0 /BaseFont /NotoSansJP \
                 /Encoding /Identity-H \
                 /DescendantFonts [{} 0 R] \
                 /ToUnicode {} 0 R >>",
                cid_font_id, tounicode_id
            )
            .into_bytes(),
        );

        // 標準フォント（Helvetica）- /F2: フォント未埋め込み時のフォールバック
        self.add_object(
            fallback_font_id,
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
                .to_vec(),
        );

        // フォントファイルの埋め込み（CIDToGIDMapと同じusable_font_dataを使用）
        if let (Some(ff_id), Some(ref font_data)) =
            (font_file_id, &usable_font_data)
        {
            self.add_object(
                ff_id,
                format!(
                    "<< /Length {} /Length1 {} >>\nstream\n",
                    font_data.len(),
                    font_data.len()
                )
                .into_bytes(),
            );
            // フォントストリームは特別処理（後でバイナリデータを追加）
            if let Some(obj) = self.objects.last_mut() {
                obj.data.extend_from_slice(font_data);
                obj.data.extend_from_slice(b"\nendstream");
            }
        }

        // 各ページ
        for (i, page) in doc.pages.iter().enumerate() {
            let (page_id, content_id) = page_content_pairs[i];

            // ページ内の画像を収集してXObjectを作成
            let image_xobjects = self.create_page_image_xobjects(page);

            // ページコンテンツストリーム（画像参照付き）
            let content = self.render_page_content(page, has_font, &image_xobjects);
            self.add_object(
                content_id,
                format!(
                    "<< /Length {} >>\nstream\n{}\nendstream",
                    content.len(),
                    String::from_utf8_lossy(&content)
                )
                .into_bytes(),
            );

            // XObjectリソース辞書を構築
            let xobj_dict = if image_xobjects.is_empty() {
                String::new()
            } else {
                let refs: Vec<String> = image_xobjects.iter()
                    .map(|img| format!("/{} {} 0 R", img.name, img.obj_id))
                    .collect();
                format!(" /XObject << {} >>", refs.join(" "))
            };

            // ページオブジェクト（/F1: CIDフォント, /F2: Helveticaフォールバック + XObject）
            self.add_object(
                page_id,
                format!(
                    "<< /Type /Page /Parent {} 0 R \
                     /MediaBox [0 0 {} {}] \
                     /Contents {} 0 R \
                     /Resources << /Font << /F1 {} 0 R /F2 {} 0 R >>{} >> >>",
                    pages_id, page.width, page.height, content_id, font_id,
                    fallback_font_id, xobj_dict
                )
                .into_bytes(),
            );
            self.page_ids.push(page_id);
        }

        // PDF出力
        self.serialize(catalog_id)
    }

    /// ページコンテンツのPDFストリームを生成
    fn render_page_content(&self, page: &Page, has_font: bool, image_xobjects: &[PdfImageXObject]) -> Vec<u8> {
        let mut stream = Vec::new();
        let mut img_idx = 0usize; // 画像XObjectカウンター

        for element in &page.elements {
            match element {
                PageElement::Text {
                    x,
                    y,
                    width,
                    text,
                    style,
                    align,
                } => {
                    self.render_text(&mut stream, *x, *y, *width, text, style, *align, page.height, has_font);
                }
                PageElement::Line {
                    x1,
                    y1,
                    x2,
                    y2,
                    width,
                    color,
                } => {
                    let py1 = page.height - y1;
                    let py2 = page.height - y2;
                    stream.extend_from_slice(
                        format!(
                            "{} {} {} RG\n{} w\n{} {} m\n{} {} l\nS\n",
                            color.r as f64 / 255.0,
                            color.g as f64 / 255.0,
                            color.b as f64 / 255.0,
                            width,
                            x1,
                            py1,
                            x2,
                            py2
                        )
                        .as_bytes(),
                    );
                }
                PageElement::Rect {
                    x,
                    y,
                    width,
                    height,
                    fill,
                    stroke,
                    stroke_width,
                } => {
                    let py = page.height - y - height;
                    if let Some(fill_color) = fill {
                        stream.extend_from_slice(
                            format!(
                                "{} {} {} rg\n{} {} {} {} re\nf\n",
                                fill_color.r as f64 / 255.0,
                                fill_color.g as f64 / 255.0,
                                fill_color.b as f64 / 255.0,
                                x,
                                py,
                                width,
                                height
                            )
                            .as_bytes(),
                        );
                    }
                    if let Some(stroke_color) = stroke {
                        stream.extend_from_slice(
                            format!(
                                "{} {} {} RG\n{} w\n{} {} {} {} re\nS\n",
                                stroke_color.r as f64 / 255.0,
                                stroke_color.g as f64 / 255.0,
                                stroke_color.b as f64 / 255.0,
                                stroke_width,
                                x,
                                py,
                                width,
                                height
                            )
                            .as_bytes(),
                        );
                    }
                }
                PageElement::Image {
                    x: img_x,
                    y: img_y,
                    width: img_w,
                    height: img_h,
                    ..
                } => {
                    // 画像XObjectを配置
                    if let Some(xobj) = image_xobjects.get(img_idx) {
                        let py = page.height - img_y - img_h;
                        stream.extend_from_slice(
                            format!(
                                "q\n{} 0 0 {} {} {} cm\n/{} Do\nQ\n",
                                img_w, img_h, img_x, py, xobj.name
                            )
                            .as_bytes(),
                        );
                    }
                    img_idx += 1;
                }
                PageElement::GradientRect {
                    x,
                    y,
                    width: w,
                    height: h,
                    stops,
                    gradient_type,
                } => {
                    // Approximate gradient with multiple thin strips
                    self.render_gradient_rect(
                        &mut stream, *x, *y, *w, *h, stops, gradient_type, page.height,
                    );
                }
                PageElement::Ellipse {
                    cx,
                    cy,
                    rx,
                    ry,
                    fill,
                    stroke,
                    stroke_width,
                } => {
                    self.render_ellipse(
                        &mut stream, *cx, *cy, *rx, *ry, fill, stroke, *stroke_width,
                        page.height,
                    );
                }
                PageElement::EllipseImage {
                    cx,
                    cy,
                    rx,
                    ry,
                    data: _,
                    mime_type: _,
                    stroke,
                    stroke_width,
                } => {
                    // 楕円領域に画像XObjectを配置
                    let img_x = *cx - *rx;
                    let img_y = *cy - *ry;
                    let img_w = *rx * 2.0;
                    let img_h = *ry * 2.0;
                    if let Some(xobj) = image_xobjects.get(img_idx) {
                        let py = page.height - img_y - img_h;
                        stream.extend_from_slice(
                            format!(
                                "q\n{} 0 0 {} {} {} cm\n/{} Do\nQ\n",
                                img_w, img_h, img_x, py, xobj.name
                            )
                            .as_bytes(),
                        );
                    }
                    img_idx += 1;
                    // Draw ellipse outline if stroke is specified
                    if stroke.is_some() && *stroke_width > 0.0 {
                        self.render_ellipse(
                            &mut stream, *cx, *cy, *rx, *ry, &None, stroke, *stroke_width,
                            page.height,
                        );
                    }
                }
                PageElement::TableBlock {
                    x,
                    y,
                    width,
                    table,
                } => {
                    self.render_table(&mut stream, *x, *y, *width, table, page.height, has_font);
                }
                PageElement::Path {
                    commands,
                    fill,
                    stroke,
                    stroke_width,
                } => {
                    self.render_path(
                        &mut stream, commands, fill, stroke, *stroke_width, page.height,
                    );
                }
                PageElement::PathImage {
                    commands,
                    data: _,
                    mime_type: _,
                    stroke,
                    stroke_width,
                } => {
                    // パスのバウンディングボックスに画像XObjectを配置
                    let mut min_x = f64::INFINITY;
                    let mut min_y = f64::INFINITY;
                    let mut max_x = f64::NEG_INFINITY;
                    let mut max_y = f64::NEG_INFINITY;

                    for cmd in commands.iter() {
                        match cmd {
                            crate::converter::PathCommand::MoveTo(x, y)
                            | crate::converter::PathCommand::LineTo(x, y) => {
                                min_x = min_x.min(*x);
                                min_y = min_y.min(*y);
                                max_x = max_x.max(*x);
                                max_y = max_y.max(*y);
                            }
                            crate::converter::PathCommand::QuadTo(cx, cy, x, y) => {
                                min_x = min_x.min(*cx).min(*x);
                                min_y = min_y.min(*cy).min(*y);
                                max_x = max_x.max(*cx).max(*x);
                                max_y = max_y.max(*cy).max(*y);
                            }
                            crate::converter::PathCommand::CubicTo(cx1, cy1, cx2, cy2, x, y) => {
                                min_x = min_x.min(*cx1).min(*cx2).min(*x);
                                min_y = min_y.min(*cy1).min(*cy2).min(*y);
                                max_x = max_x.max(*cx1).max(*cx2).max(*x);
                                max_y = max_y.max(*cy1).max(*cy2).max(*y);
                            }
                            crate::converter::PathCommand::ArcTo(rx, ry, _, _, _, x, y) => {
                                min_x = min_x.min(*x).min(*x - rx);
                                min_y = min_y.min(*y).min(*y - ry);
                                max_x = max_x.max(*x).max(*x + rx);
                                max_y = max_y.max(*y).max(*y + ry);
                            }
                            crate::converter::PathCommand::Close => {}
                        }
                    }

                    if min_x < max_x && min_y < max_y {
                        if let Some(xobj) = image_xobjects.get(img_idx) {
                            let img_w = max_x - min_x;
                            let img_h = max_y - min_y;
                            let py = page.height - min_y - img_h;
                            stream.extend_from_slice(
                                format!(
                                    "q\n{} 0 0 {} {} {} cm\n/{} Do\nQ\n",
                                    img_w, img_h, min_x, py, xobj.name
                                )
                                .as_bytes(),
                            );
                        }
                    }
                    img_idx += 1;

                    // Render the path stroke if specified
                    if stroke.is_some() && *stroke_width > 0.0 {
                        self.render_path(
                            &mut stream, commands, &None, stroke, *stroke_width, page.height,
                        );
                    }
                }
            }
        }

        stream
    }

    /// テキストをPDFストリームに出力
    fn render_text(
        &self,
        stream: &mut Vec<u8>,
        x: f64,
        y: f64,
        _width: f64,
        text: &str,
        style: &FontStyle,
        _align: TextAlign,
        page_height: f64,
        has_font: bool,
    ) {
        if text.is_empty() {
            return;
        }
        // 制御文字をサニタイズしてPDFテキストオペレータの破損を防止
        // - 改行・復帰・タブはスペースに正規化して単語境界を保持
        // - それ以外の制御文字は除去
        let clean_text: String = text
            .chars()
            .map(|c| match c {
                '\n' | '\r' | '\t' => ' ',
                _ => c,
            })
            .filter(|c| !c.is_control())
            .collect();
        if clean_text.is_empty() {
            return;
        }
        let pdf_y = page_height - y - style.font_size;

        if has_font {
            // CIDフォント（/F1）: UTF-16BEヘックス文字列で全Unicode対応
            let hex_text = self.text_to_pdf_hex(&clean_text);

            stream.extend_from_slice(
                format!(
                    "BT\n/F1 {} Tf\n{} {} {} rg\n{} {} Td\n<{}> Tj\nET\n",
                    style.font_size,
                    style.color.r as f64 / 255.0,
                    style.color.g as f64 / 255.0,
                    style.color.b as f64 / 255.0,
                    x,
                    pdf_y,
                    hex_text
                )
                .as_bytes(),
            );
        } else {
            // フォールバック（/F2 Helvetica）: WinAnsiEncoding（Latin-1）
            // 非ASCII文字を '?' に置換してLatin-1の範囲内に収める
            let safe_text = text_to_winansi(&clean_text);
            let escaped = pdf_escape_string(&safe_text);

            stream.extend_from_slice(
                format!(
                    "BT\n/F2 {} Tf\n{} {} {} rg\n{} {} Td\n({}) Tj\nET\n",
                    style.font_size,
                    style.color.r as f64 / 255.0,
                    style.color.g as f64 / 255.0,
                    style.color.b as f64 / 255.0,
                    x,
                    pdf_y,
                    escaped
                )
                .as_bytes(),
            );
        }
    }

    /// テーブルをPDFストリームに出力
    fn render_table(
        &self,
        stream: &mut Vec<u8>,
        x: f64,
        y: f64,
        width: f64,
        table: &Table,
        page_height: f64,
        has_font: bool,
    ) {
        let base_row_height = 20.0;
        let padding = 4.0;
        let line_spacing = 1.3;
        let num_cols = table.column_widths.len().max(1);
        let col_width = if table.column_widths.is_empty() {
            width / num_cols as f64
        } else {
            0.0 // 個別幅を使用
        };

        // 各行の高さを事前計算（セルごとの行高見積もりの最大に基づく動的高さ）
        let row_heights: Vec<f64> = table.rows.iter().map(|row| {
            let row_max_height = row
                .iter()
                .map(|cell| {
                    let lines = if cell.text.is_empty() {
                        1
                    } else {
                        cell.text.split('\n').count().max(1)
                    };
                    let fs = cell.style.font_size;
                    fs * line_spacing * lines as f64 + padding * 2.0
                })
                .max_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal))
                .unwrap_or(base_row_height);
            row_max_height.max(base_row_height)
        }).collect();

        let mut row_y_offset = 0.0f64;
        for (row_idx, row) in table.rows.iter().enumerate() {
            let row_y = y + row_y_offset;
            let rh = row_heights[row_idx];
            let mut cell_x = x;

            for (col_idx, cell) in row.iter().enumerate() {
                let cw = if table.column_widths.is_empty() || col_idx >= table.column_widths.len() {
                    col_width
                } else {
                    table.column_widths[col_idx]
                };

                // セル枠線
                let py = page_height - row_y - rh;
                stream.extend_from_slice(
                    format!(
                        "0.8 0.8 0.8 RG\n0.5 w\n{} {} {} {} re\nS\n",
                        cell_x, py, cw, rh
                    )
                    .as_bytes(),
                );

                // セルテキスト（複数行対応）
                if !cell.text.is_empty() {
                    let fs = cell.style.font_size;
                    let font_name = if has_font { "F1" } else { "F2" };
                    let lines: Vec<&str> = cell.text.split('\n').collect();
                    for (line_idx, line) in lines.iter().enumerate() {
                        let text_y = page_height - row_y - padding - fs
                            - (line_idx as f64 * fs * line_spacing);
                        if line.is_empty() {
                            // 空行はYオフセットだけ進める（描画はスキップ）
                            continue;
                        }
                        if has_font {
                            let hex_text = self.text_to_pdf_hex(line);
                            stream.extend_from_slice(
                                format!(
                                    "BT\n/{} {} Tf\n{} {} {} rg\n{} {} Td\n<{}> Tj\nET\n",
                                    font_name,
                                    fs,
                                    cell.style.color.r as f64 / 255.0,
                                    cell.style.color.g as f64 / 255.0,
                                    cell.style.color.b as f64 / 255.0,
                                    cell_x + padding,
                                    text_y,
                                    hex_text
                                )
                                .as_bytes(),
                            );
                        } else {
                            let safe_text = text_to_winansi(line);
                            let escaped = pdf_escape_string(&safe_text);
                            stream.extend_from_slice(
                                format!(
                                    "BT\n/{} {} Tf\n{} {} {} rg\n{} {} Td\n({}) Tj\nET\n",
                                    font_name,
                                    fs,
                                    cell.style.color.r as f64 / 255.0,
                                    cell.style.color.g as f64 / 255.0,
                                    cell.style.color.b as f64 / 255.0,
                                    cell_x + padding,
                                    text_y,
                                    escaped
                                )
                                .as_bytes(),
                            );
                        }
                    }
                }

                cell_x += cw;
            }
            row_y_offset += rh;
        }
    }

    /// グラデーション矩形をPDFストリームに出力（ストライプ近似）
    #[allow(clippy::too_many_arguments)]
    fn render_gradient_rect(
        &self,
        stream: &mut Vec<u8>,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        stops: &[GradientStop],
        gradient_type: &GradientType,
        page_height: f64,
    ) {
        if stops.is_empty() || w <= 0.0 || h <= 0.0 {
            return;
        }

        // Approximate with thin horizontal/vertical strips
        const GRADIENT_STRIP_COUNT: u32 = 50;
        let num_strips = GRADIENT_STRIP_COUNT;
        let strip_height = h / num_strips as f64;

        for i in 0..num_strips {
            let t = match gradient_type {
                GradientType::Linear(angle) => {
                    let local_y = i as f64 / num_strips as f64;
                    let cos_a = angle.cos();
                    let sin_a = angle.sin();
                    (0.5 * sin_a + local_y * cos_a).clamp(0.0, 1.0)
                }
                GradientType::Radial => {
                    let local_y = (i as f64 / num_strips as f64 - 0.5).abs() * 2.0;
                    local_y.min(1.0)
                }
            };

            let color = Self::interpolate_gradient_color(stops, t);
            let strip_y = page_height - y - (i + 1) as f64 * strip_height;
            stream.extend_from_slice(
                format!(
                    "{:.4} {:.4} {:.4} rg\n{} {} {} {} re\nf\n",
                    color.r as f64 / 255.0,
                    color.g as f64 / 255.0,
                    color.b as f64 / 255.0,
                    x,
                    strip_y,
                    w,
                    strip_height + 0.5 // Slight overlap to avoid gaps
                )
                .as_bytes(),
            );
        }
    }

    /// グラデーション停止点間の色を補間
    fn interpolate_gradient_color(stops: &[GradientStop], t: f64) -> Color {
        if stops.is_empty() {
            return Color::WHITE;
        }
        if stops.len() == 1 || t <= stops[0].position {
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
                let lt = (t - stops[i].position) / range;
                let c1 = &stops[i].color;
                let c2 = &stops[i + 1].color;
                return Color {
                    r: (c1.r as f64 + (c2.r as f64 - c1.r as f64) * lt) as u8,
                    g: (c1.g as f64 + (c2.g as f64 - c1.g as f64) * lt) as u8,
                    b: (c1.b as f64 + (c2.b as f64 - c1.b as f64) * lt) as u8,
                    a: 255,
                };
            }
        }
        stops[0].color
    }

    /// 楕円をPDFストリームに出力（ベジェ曲線近似）
    #[allow(clippy::too_many_arguments)]
    fn render_ellipse(
        &self,
        stream: &mut Vec<u8>,
        cx: f64,
        cy: f64,
        rx: f64,
        ry: f64,
        fill: &Option<Color>,
        stroke: &Option<Color>,
        stroke_width: f64,
        page_height: f64,
    ) {
        let pcy = page_height - cy;
        // Bezier approximation of ellipse: kappa = 4 * (sqrt(2) - 1) / 3
        let k = 0.5522847498;
        let kx = rx * k;
        let ky = ry * k;

        let path = format!(
            "{} {} m\n{} {} {} {} {} {} c\n{} {} {} {} {} {} c\n{} {} {} {} {} {} c\n{} {} {} {} {} {} c\n",
            cx + rx, pcy,
            cx + rx, pcy + ky, cx + kx, pcy + ry, cx, pcy + ry,
            cx - kx, pcy + ry, cx - rx, pcy + ky, cx - rx, pcy,
            cx - rx, pcy - ky, cx - kx, pcy - ry, cx, pcy - ry,
            cx + kx, pcy - ry, cx + rx, pcy - ky, cx + rx, pcy,
        );

        if let Some(fill_color) = fill {
            stream.extend_from_slice(
                format!(
                    "{} {} {} rg\n{}f\n",
                    fill_color.r as f64 / 255.0,
                    fill_color.g as f64 / 255.0,
                    fill_color.b as f64 / 255.0,
                    path
                )
                .as_bytes(),
            );
        }
        if let Some(stroke_color) = stroke {
            stream.extend_from_slice(
                format!(
                    "{} {} {} RG\n{} w\n{}S\n",
                    stroke_color.r as f64 / 255.0,
                    stroke_color.g as f64 / 255.0,
                    stroke_color.b as f64 / 255.0,
                    stroke_width,
                    path
                )
                .as_bytes(),
            );
        }
    }

    /// テキストをPDF用UTF-16BEヘックス文字列に変換
    fn text_to_pdf_hex(&self, text: &str) -> String {
        let mut hex = String::new();
        for ch in text.chars() {
            let code = ch as u32;
            if code <= 0xFFFF {
                hex.push_str(&format!("{:04X}", code));
            } else {
                // サロゲートペア
                let code = code - 0x10000;
                let high = 0xD800 + (code >> 10);
                let low = 0xDC00 + (code & 0x3FF);
                hex.push_str(&format!("{:04X}{:04X}", high, low));
            }
        }
        hex
    }

    /// パスをPDFストリームに出力
    fn render_path(
        &self,
        stream: &mut Vec<u8>,
        commands: &[crate::converter::PathCommand],
        fill: &Option<Color>,
        stroke: &Option<Color>,
        stroke_width: f64,
        page_height: f64,
    ) {
        use crate::converter::PathCommand;

        let mut path_str = String::new();
        let mut cur_x = 0.0_f64;
        let mut cur_y = 0.0_f64;
        for cmd in commands {
            match cmd {
                PathCommand::MoveTo(x, y) => {
                    cur_x = *x;
                    cur_y = *y;
                    path_str.push_str(&format!("{} {} m\n", x, page_height - y));
                }
                PathCommand::LineTo(x, y) => {
                    cur_x = *x;
                    cur_y = *y;
                    path_str.push_str(&format!("{} {} l\n", x, page_height - y));
                }
                PathCommand::QuadTo(qcx, qcy, x, y) => {
                    // Convert quadratic to cubic: cp1 = cur + 2/3*(qc - cur), cp2 = end + 2/3*(qc - end)
                    let cp1x = cur_x + 2.0 / 3.0 * (qcx - cur_x);
                    let cp1y = cur_y + 2.0 / 3.0 * (qcy - cur_y);
                    let cp2x = x + 2.0 / 3.0 * (qcx - x);
                    let cp2y = y + 2.0 / 3.0 * (qcy - y);
                    cur_x = *x;
                    cur_y = *y;
                    path_str.push_str(&format!(
                        "{} {} {} {} {} {} c\n",
                        cp1x, page_height - cp1y, cp2x, page_height - cp2y, x, page_height - y
                    ));
                }
                PathCommand::CubicTo(cx1, cy1, cx2, cy2, x, y) => {
                    cur_x = *x;
                    cur_y = *y;
                    path_str.push_str(&format!(
                        "{} {} {} {} {} {} c\n",
                        cx1, page_height - cy1, cx2, page_height - cy2, x, page_height - y
                    ));
                }
                PathCommand::ArcTo(_rx, _ry, _rot, _large, _sweep, x, y) => {
                    // Approximate as line (proper arc-to-bezier conversion is complex)
                    cur_x = *x;
                    cur_y = *y;
                    path_str.push_str(&format!("{} {} l\n", x, page_height - y));
                }
                PathCommand::Close => {
                    path_str.push_str("h\n");
                }
            }
        }

        if let Some(fill_color) = fill {
            stream.extend_from_slice(
                format!(
                    "{} {} {} rg\n{}f\n",
                    fill_color.r as f64 / 255.0,
                    fill_color.g as f64 / 255.0,
                    fill_color.b as f64 / 255.0,
                    path_str
                )
                .as_bytes(),
            );
        }
        if let Some(stroke_color) = stroke {
            stream.extend_from_slice(
                format!(
                    "{} {} {} RG\n{} w\n{}S\n",
                    stroke_color.r as f64 / 255.0,
                    stroke_color.g as f64 / 255.0,
                    stroke_color.b as f64 / 255.0,
                    stroke_width,
                    path_str
                )
                .as_bytes(),
            );
        }
    }

    /// ToUnicode CMapを生成
    fn create_tounicode_cmap(&self) -> String {
        // Identity CMap: CIDがそのままUnicodeコードポイントに対応
        "/CIDInit /ProcSet findresource begin\n\
         12 dict begin\n\
         begincmap\n\
         /CIDSystemInfo\n\
         << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n\
         /CMapName /Adobe-Identity-UCS def\n\
         /CMapType 2 def\n\
         1 begincodespacerange\n\
         <0000> <FFFF>\n\
         endcodespacerange\n\
         1 beginbfrange\n\
         <0000> <FFFF> <0000>\n\
         endbfrange\n\
         endcmap\n\
         CMapName currentdict /CMap defineresource pop\n\
         end\n\
         end"
            .to_string()
    }

    /// CIDToGIDMapストリームを生成
    /// フォントのcmapテーブルからUnicodeコードポイント(CID)→グリフID(GID)のマッピングを構築
    /// 指定されたフォントデータを使用（find_usable_font_dataで検証済みのもの）
    fn build_cid_to_gid_map_from_data(font_data: Option<&[u8]>) -> Vec<u8> {
        use ab_glyph::{Font, FontRef};
        // 65536 entries × 2 bytes each = 131072 bytes
        let mut map = vec![0u8; 65536 * 2];

        if let Some(data) = font_data {
            if let Ok(font) = FontRef::try_from_slice(data) {
                for code_point in 0u32..=0xFFFF {
                    if let Some(ch) = char::from_u32(code_point) {
                        let glyph_id = font.glyph_id(ch);
                        let gid = glyph_id.0;
                        let offset = (code_point as usize) * 2;
                        map[offset] = (gid >> 8) as u8;
                        map[offset + 1] = (gid & 0xFF) as u8;
                    }
                }
            }
        }
        map
    }

    /// ページ内の画像要素からPDF XObjectを作成
    fn create_page_image_xobjects(&mut self, page: &Page) -> Vec<PdfImageXObject> {
        let mut xobjects = Vec::new();
        let mut counter = 0u32;

        for element in &page.elements {
            let image_data: Option<(&[u8], &str)> = match element {
                PageElement::Image { data, mime_type, .. } => Some((data, mime_type)),
                PageElement::EllipseImage { data, mime_type, .. } => Some((data, mime_type)),
                PageElement::PathImage { data, mime_type, .. } => Some((data, mime_type)),
                _ => None,
            };

            if let Some((data, mime_type)) = image_data {
                let name = format!("Im{}", counter);
                let is_jpeg = mime_type.contains("jpeg") || mime_type.contains("jpg")
                    || (data.len() >= 2 && data[0] == 0xFF && data[1] == 0xD8);

                if is_jpeg {
                    if let Some((w, h)) = extract_jpeg_dimensions(data) {
                        let obj_id = self.alloc_id();
                        let mut obj_data = format!(
                            "<< /Type /XObject /Subtype /Image /Width {} /Height {} \
                             /ColorSpace /DeviceRGB /BitsPerComponent 8 \
                             /Filter /DCTDecode /Length {} >>\nstream\n",
                            w, h, data.len()
                        ).into_bytes();
                        obj_data.extend_from_slice(data);
                        obj_data.extend_from_slice(b"\nendstream");
                        self.add_object(obj_id, obj_data);
                        xobjects.push(PdfImageXObject { name, obj_id, smask_id: None });
                    }
                } else if let Some((w, h, rgb_data, alpha_data)) = decode_image_to_raw_rgb(data) {
                    let obj_id = self.alloc_id();

                    // アルファチャンネルがあればSMaskを作成
                    let smask_id = if let Some(ref alpha) = alpha_data {
                        let sid = self.alloc_id();
                        let compressed_alpha = miniz_oxide::deflate::compress_to_vec(alpha, 6);
                        let mut smask_data = format!(
                            "<< /Type /XObject /Subtype /Image /Width {} /Height {} \
                             /ColorSpace /DeviceGray /BitsPerComponent 8 \
                             /Filter /FlateDecode /Length {} >>\nstream\n",
                            w, h, compressed_alpha.len()
                        ).into_bytes();
                        smask_data.extend_from_slice(&compressed_alpha);
                        smask_data.extend_from_slice(b"\nendstream");
                        self.add_object(sid, smask_data);
                        Some(sid)
                    } else {
                        None
                    };

                    let compressed_rgb = miniz_oxide::deflate::compress_to_vec(&rgb_data, 6);
                    let smask_ref = smask_id.map_or(String::new(),
                        |sid| format!(" /SMask {} 0 R", sid));
                    let mut obj_data = format!(
                        "<< /Type /XObject /Subtype /Image /Width {} /Height {} \
                         /ColorSpace /DeviceRGB /BitsPerComponent 8 \
                         /Filter /FlateDecode /Length {}{} >>\nstream\n",
                        w, h, compressed_rgb.len(), smask_ref
                    ).into_bytes();
                    obj_data.extend_from_slice(&compressed_rgb);
                    obj_data.extend_from_slice(b"\nendstream");
                    self.add_object(obj_id, obj_data);
                    xobjects.push(PdfImageXObject { name, obj_id, smask_id });
                }
                counter += 1;
            }
        }

        xobjects
    }

    /// PDFバイト列をシリアライズ
    fn serialize(&self, catalog_id: u32) -> Vec<u8> {
        let mut output = Vec::new();

        // ヘッダー
        output.extend_from_slice(b"%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");

        // オブジェクト書き出しとオフセット記録
        let mut offsets: Vec<(u32, usize)> = Vec::new();

        for obj in &self.objects {
            offsets.push((obj.id, output.len()));
            output.extend_from_slice(format!("{} 0 obj\n", obj.id).as_bytes());
            output.extend_from_slice(&obj.data);
            output.extend_from_slice(b"\nendobj\n\n");
        }

        // 相互参照テーブル
        let xref_offset = output.len();
        output.extend_from_slice(b"xref\n");

        let max_id = self.next_id;
        output.extend_from_slice(format!("0 {}\n", max_id).as_bytes());
        output.extend_from_slice(b"0000000000 65535 f \n");

        for id in 1..max_id {
            if let Some((_, off)) = offsets.iter().find(|(oid, _)| *oid == id) {
                output.extend_from_slice(format!("{:010} 00000 n \n", off).as_bytes());
            } else {
                // オブジェクトが存在しないIDはfree(f)エントリとして出力
                output.extend_from_slice(b"0000000000 00000 f \n");
            }
        }

        // トレーラー
        output.extend_from_slice(
            format!(
                "trailer\n<< /Size {} /Root {} 0 R >>\n",
                max_id, catalog_id
            )
            .as_bytes(),
        );
        output.extend_from_slice(format!("startxref\n{}\n%%EOF\n", xref_offset).as_bytes());

        output
    }
}

/// ab_glyphでパース可能なフォントデータを見つける（フリー関数版）
/// 外部フォント → 内蔵フォントの順に試行し、パース可能な最初のフォントデータをコピーして返す
fn find_usable_font_data(font_manager: &FontManager) -> Option<Vec<u8>> {
    use ab_glyph::FontRef;
    // まず外部フォントを順に試す
    for (_, data) in font_manager.external_fonts_iter() {
        if !data.is_empty() && FontRef::try_from_slice(data).is_ok() {
            return Some(data.to_vec());
        }
    }
    // 内蔵フォントを試す
    if let Some(data) = font_manager.builtin_japanese_font() {
        if FontRef::try_from_slice(data).is_ok() {
            return Some(data.to_vec());
        }
    }
    // LINE Seed JPを試す
    if let Some(data) = font_manager.builtin_line_seed_jp() {
        if FontRef::try_from_slice(data).is_ok() {
            return Some(data.to_vec());
        }
    }
    None
}

/// テキストをASCIIに制限してHelveticaフォールバック用に使用
/// PDF literal string `(...)` はバイト列として解釈されるため、
/// Rust String（UTF-8）のマルチバイト文字を直接書くと破損する。
/// Helvetica（/WinAnsiEncoding）は基本的にLatin-1相当だが、
/// UTF-8→Latin-1変換の複雑さを避け、安全にASCII範囲に制限する。
fn text_to_winansi(text: &str) -> String {
    text.chars()
        .map(|c| {
            let cp = u32::from(c);
            if cp >= 0x20 && cp <= 0x7E {
                c
            } else {
                '?'
            }
        })
        .collect()
}

/// PDFリテラル文字列用エスケープ
fn pdf_escape_string(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '\\' => result.push_str("\\\\"),
            '(' => result.push_str("\\("),
            ')' => result.push_str("\\)"),
            _ => result.push(c),
        }
    }
    result
}

/// ドキュメントをPDFバイト列に変換する便利関数
pub fn render_to_pdf(doc: &Document) -> Vec<u8> {
    let fm = FontManager::new();
    render_to_pdf_with_fonts(doc, &fm)
}

/// フォントマネージャーを使用してドキュメントをPDFバイト列に変換
/// 外部から読み込んだフォントを使用する場合はこちらを使用してください。
pub fn render_to_pdf_with_fonts(doc: &Document, font_manager: &FontManager) -> Vec<u8> {
    let mut writer = PdfWriter::new(font_manager);
    writer.render(doc)
}

/// JPEGバイト列から画像の幅と高さを抽出する
fn extract_jpeg_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    if data.len() < 4 || data[0] != 0xFF || data[1] != 0xD8 {
        return None;
    }

    let mut i = 2;
    while i + 1 < data.len() {
        if data[i] != 0xFF {
            i += 1;
            continue;
        }
        let marker = data[i + 1];
        i += 2;

        // SOF markers (SOF0-SOF3, SOF5-SOF7, SOF9-SOF11, SOF13-SOF15)
        if matches!(marker, 0xC0..=0xC3 | 0xC5..=0xC7 | 0xC9..=0xCB | 0xCD..=0xCF) {
            if i + 7 <= data.len() {
                let height = ((data[i + 3] as u32) << 8) | (data[i + 4] as u32);
                let width = ((data[i + 5] as u32) << 8) | (data[i + 6] as u32);
                if width > 0 && height > 0 {
                    return Some((width, height));
                }
            }
            return None;
        }

        // Skip non-SOF markers by reading segment length
        if i + 1 < data.len() {
            let seg_len = ((data[i] as usize) << 8) | (data[i + 1] as usize);
            if seg_len < 2 {
                return None;
            }
            i += seg_len;
        } else {
            break;
        }
    }
    None
}

/// 画像データ（PNG/JPEG）をRGBバイト列にデコードする
/// 戻り値: (width, height, rgb_data, optional_alpha_data)
fn decode_image_to_raw_rgb(data: &[u8]) -> Option<(u32, u32, Vec<u8>, Option<Vec<u8>>)> {
    // Try PNG first
    let is_png = data.len() >= 4 && data[0..4] == [0x89, 0x50, 0x4E, 0x47];
    if is_png {
        return decode_png_to_raw_rgb(data);
    }
    // Try JPEG
    let is_jpeg = data.len() >= 2 && data[0] == 0xFF && data[1] == 0xD8;
    if is_jpeg {
        return decode_jpeg_to_raw_rgb(data);
    }
    None
}

/// PNGデータをRGBバイト列にデコード
fn decode_png_to_raw_rgb(data: &[u8]) -> Option<(u32, u32, Vec<u8>, Option<Vec<u8>>)> {
    let decoder = png::Decoder::new(std::io::Cursor::new(data));
    let mut reader = decoder.read_info().ok()?;
    let mut buf = vec![0u8; reader.output_buffer_size()];
    let info = reader.next_frame(&mut buf).ok()?;
    let width = info.width;
    let height = info.height;
    let pixels = &buf[..info.buffer_size()];

    match info.color_type {
        png::ColorType::Rgba => {
            let npixels = (width * height) as usize;
            let mut rgb = Vec::with_capacity(npixels * 3);
            let mut alpha = Vec::with_capacity(npixels);
            let mut has_transparency = false;
            for chunk in pixels.chunks_exact(4) {
                rgb.push(chunk[0]);
                rgb.push(chunk[1]);
                rgb.push(chunk[2]);
                alpha.push(chunk[3]);
                if chunk[3] != 255 { has_transparency = true; }
            }
            let alpha_data = if has_transparency { Some(alpha) } else { None };
            Some((width, height, rgb, alpha_data))
        }
        png::ColorType::Rgb => {
            Some((width, height, pixels.to_vec(), None))
        }
        png::ColorType::GrayscaleAlpha => {
            let npixels = (width * height) as usize;
            let mut rgb = Vec::with_capacity(npixels * 3);
            let mut alpha = Vec::with_capacity(npixels);
            let mut has_transparency = false;
            for chunk in pixels.chunks_exact(2) {
                rgb.push(chunk[0]);
                rgb.push(chunk[0]);
                rgb.push(chunk[0]);
                alpha.push(chunk[1]);
                if chunk[1] != 255 { has_transparency = true; }
            }
            let alpha_data = if has_transparency { Some(alpha) } else { None };
            Some((width, height, rgb, alpha_data))
        }
        png::ColorType::Grayscale => {
            let npixels = (width * height) as usize;
            let mut rgb = Vec::with_capacity(npixels * 3);
            for &g in pixels.iter() {
                rgb.push(g);
                rgb.push(g);
                rgb.push(g);
            }
            Some((width, height, rgb, None))
        }
        png::ColorType::Indexed => {
            // Indexed PNG: expand palette
            let info2 = reader.info();
            let palette = info2.palette.as_ref()?;
            let trns = info2.trns.as_ref();
            let npixels = (width * height) as usize;
            let mut rgb = Vec::with_capacity(npixels * 3);
            let mut alpha = Vec::with_capacity(npixels);
            let mut has_transparency = false;
            for &idx in pixels.iter().take(npixels) {
                let base = (idx as usize) * 3;
                if base + 2 < palette.len() {
                    rgb.push(palette[base]);
                    rgb.push(palette[base + 1]);
                    rgb.push(palette[base + 2]);
                } else {
                    rgb.push(0);
                    rgb.push(0);
                    rgb.push(0);
                }
                let a = trns.and_then(|t| t.get(idx as usize).copied()).unwrap_or(255);
                alpha.push(a);
                if a != 255 { has_transparency = true; }
            }
            let alpha_data = if has_transparency { Some(alpha) } else { None };
            Some((width, height, rgb, alpha_data))
        }
    }
}

/// JPEGデータをRGBバイト列にデコード
fn decode_jpeg_to_raw_rgb(data: &[u8]) -> Option<(u32, u32, Vec<u8>, Option<Vec<u8>>)> {
    let mut decoder = jpeg_decoder::Decoder::new(std::io::Cursor::new(data));
    let pixels = decoder.decode().ok()?;
    let info = decoder.info()?;
    let width = info.width as u32;
    let height = info.height as u32;

    let rgb = match info.pixel_format {
        jpeg_decoder::PixelFormat::RGB24 => pixels,
        jpeg_decoder::PixelFormat::L8 => {
            let mut rgb_data = Vec::with_capacity(pixels.len() * 3);
            for &g in &pixels {
                rgb_data.push(g);
                rgb_data.push(g);
                rgb_data.push(g);
            }
            rgb_data
        }
        jpeg_decoder::PixelFormat::CMYK32 => {
            let mut rgb_data = Vec::with_capacity((pixels.len() / 4) * 3);
            for chunk in pixels.chunks_exact(4) {
                let c = chunk[0] as f64 / 255.0;
                let m = chunk[1] as f64 / 255.0;
                let y = chunk[2] as f64 / 255.0;
                let k = chunk[3] as f64 / 255.0;
                rgb_data.push(((1.0 - c) * (1.0 - k) * 255.0) as u8);
                rgb_data.push(((1.0 - m) * (1.0 - k) * 255.0) as u8);
                rgb_data.push(((1.0 - y) * (1.0 - k) * 255.0) as u8);
            }
            rgb_data
        }
        _ => {
            // Unsupported pixel format: fallback, treat bytes as grayscale
            let mut rgb_data = Vec::with_capacity(pixels.len() * 3);
            for &g in &pixels {
                rgb_data.push(g);
                rgb_data.push(g);
                rgb_data.push(g);
            }
            rgb_data
        }
    };

    Some((width, height, rgb, None))
}
