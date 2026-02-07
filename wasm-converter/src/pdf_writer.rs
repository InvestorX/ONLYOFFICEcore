// pdf_writer.rs - 軽量PDF生成エンジン
//
// 外部クレートに依存せず、PDF 1.4仕様に準拠したPDFバイト列を直接生成します。
// 日本語テキスト（Unicode）をサポートします。

use crate::converter::{Color, Document, FontStyle, GradientStop, GradientType, Page, PageElement, Table, TextAlign};
use crate::font_manager::FontManager;

/// PDFオブジェクト
struct PdfObject {
    id: u32,
    data: Vec<u8>,
}

/// PDF生成器
pub struct PdfWriter {
    objects: Vec<PdfObject>,
    next_id: u32,
    page_ids: Vec<u32>,
    font_manager: FontManager,
}

impl PdfWriter {
    pub fn new() -> Self {
        Self {
            objects: Vec::new(),
            next_id: 1,
            page_ids: Vec::new(),
            font_manager: FontManager::new(),
        }
    }

    pub fn with_font_manager(font_manager: FontManager) -> Self {
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
        // IDを事前割り当て
        let catalog_id = self.alloc_id();
        let pages_id = self.alloc_id();
        let font_id = self.alloc_id();
        let cid_font_id = self.alloc_id();
        let descriptor_id = self.alloc_id();
        let tounicode_id = self.alloc_id();

        // フォント埋め込み用のID（フォントデータがある場合）
        let font_file_id = if self.font_manager.has_builtin_japanese_font() {
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

        // CIDFont
        let mut cid_font_dict = format!(
            "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /NotoSansJP \
             /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> \
             /DW 1000 \
             /FontDescriptor {} 0 R",
            descriptor_id
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

        // Type0フォント（複合フォント）
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

        // フォントファイルの埋め込み
        if let (Some(ff_id), Some(font_data)) =
            (font_file_id, self.font_manager.builtin_japanese_font())
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

            // ページコンテンツストリーム
            let content = self.render_page_content(page);
            self.add_object(
                content_id,
                format!(
                    "<< /Length {} >>\nstream\n{}\nendstream",
                    content.len(),
                    String::from_utf8_lossy(&content)
                )
                .into_bytes(),
            );

            // ページオブジェクト
            self.add_object(
                page_id,
                format!(
                    "<< /Type /Page /Parent {} 0 R \
                     /MediaBox [0 0 {} {}] \
                     /Contents {} 0 R \
                     /Resources << /Font << /F1 {} 0 R >> >> >>",
                    pages_id, page.width, page.height, content_id, font_id
                )
                .into_bytes(),
            );
            self.page_ids.push(page_id);
        }

        // PDF出力
        self.serialize(catalog_id)
    }

    /// ページコンテンツのPDFストリームを生成
    fn render_page_content(&self, page: &Page) -> Vec<u8> {
        let mut stream = Vec::new();

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
                    self.render_text(&mut stream, *x, *y, *width, text, style, *align, page.height);
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
                    // Draw a placeholder rectangle for the image position in PDF
                    // (Full image XObject embedding is complex; we show the bounding box)
                    let py = page.height - img_y - img_h;
                    stream.extend_from_slice(
                        format!(
                            "0.88 0.88 0.88 rg\n{} {} {} {} re\nf\n\
                             0.7 0.7 0.7 RG\n0.5 w\n{} {} {} {} re\nS\n",
                            img_x, py, img_w, img_h,
                            img_x, py, img_w, img_h
                        )
                        .as_bytes(),
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
                PageElement::TableBlock {
                    x,
                    y,
                    width,
                    table,
                } => {
                    self.render_table(&mut stream, *x, *y, *width, table, page.height);
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
    ) {
        if text.is_empty() {
            return;
        }
        let pdf_y = page_height - y - style.font_size;

        // テキストをUTF-16BEに変換してPDFヘックス文字列として出力
        let hex_text = self.text_to_pdf_hex(text);

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
    ) {
        let row_height = 20.0;
        let padding = 4.0;
        let num_cols = table.column_widths.len().max(1);
        let col_width = if table.column_widths.is_empty() {
            width / num_cols as f64
        } else {
            0.0 // 個別幅を使用
        };

        for (row_idx, row) in table.rows.iter().enumerate() {
            let row_y = y + row_idx as f64 * row_height;
            let mut cell_x = x;

            for (col_idx, cell) in row.iter().enumerate() {
                let cw = if table.column_widths.is_empty() || col_idx >= table.column_widths.len() {
                    col_width
                } else {
                    table.column_widths[col_idx]
                };

                // セル枠線
                let py = page_height - row_y - row_height;
                stream.extend_from_slice(
                    format!(
                        "0.8 0.8 0.8 RG\n0.5 w\n{} {} {} {} re\nS\n",
                        cell_x, py, cw, row_height
                    )
                    .as_bytes(),
                );

                // セルテキスト
                if !cell.text.is_empty() {
                    let hex_text = self.text_to_pdf_hex(&cell.text);
                    let text_y = page_height - row_y - row_height + padding;
                    stream.extend_from_slice(
                        format!(
                            "BT\n/F1 {} Tf\n0 0 0 rg\n{} {} Td\n<{}> Tj\nET\n",
                            cell.style.font_size,
                            cell_x + padding,
                            text_y,
                            hex_text
                        )
                        .as_bytes(),
                    );
                }

                cell_x += cw;
            }
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
        // BOM
        hex.push_str("FEFF");
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
            let offset = offsets
                .iter()
                .find(|(oid, _)| *oid == id)
                .map(|(_, off)| *off)
                .unwrap_or(0);
            output.extend_from_slice(format!("{:010} 00000 n \n", offset).as_bytes());
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

/// ドキュメントをPDFバイト列に変換する便利関数
pub fn render_to_pdf(doc: &Document) -> Vec<u8> {
    let mut writer = PdfWriter::new();
    writer.render(doc)
}

/// フォントマネージャーを使用してドキュメントをPDFバイト列に変換
pub fn render_to_pdf_with_fonts(doc: &Document, font_manager: FontManager) -> Vec<u8> {
    let mut writer = PdfWriter::with_font_manager(font_manager);
    writer.render(doc)
}
