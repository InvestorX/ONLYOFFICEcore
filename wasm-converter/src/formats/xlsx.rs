// formats/xlsx.rs - XLSX/XLS/ODS変換モジュール
//
// calamine クレートを使用してスプレッドシートファイルを読み込み、
// テーブル形式でドキュメントモデルに変換します。

use crate::converter::{
    ConvertError, Document, DocumentConverter, FontStyle, Metadata, Page, PageElement, Table,
    TableCell,
};
use calamine::{open_workbook_auto_from_rs, Data, Reader};

/// スプレッドシートコンバーター
pub struct XlsxConverter;

impl XlsxConverter {
    pub fn new() -> Self {
        Self
    }
}

impl DocumentConverter for XlsxConverter {
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError> {
        let cursor = std::io::Cursor::new(input);
        let mut workbook = open_workbook_auto_from_rs(cursor)
            .map_err(|e| ConvertError::new("XLSX", &format!("ワークブックを開けません: {}", e)))?;

        let sheet_names: Vec<String> = workbook.sheet_names().to_vec();

        let mut doc = Document::new();
        doc.metadata = Metadata {
            title: Some("Spreadsheet".to_string()),
            creator: Some("WASM Document Converter".to_string()),
            ..Default::default()
        };

        for sheet_name in &sheet_names {
            if let Ok(range) = workbook.worksheet_range(sheet_name) {
                let pages = render_sheet_to_pages(sheet_name, &range);
                doc.pages.extend(pages);
            }
        }

        if doc.pages.is_empty() {
            doc.pages.push(Page::a4());
        }

        Ok(doc)
    }

    fn supported_extensions(&self) -> &[&str] {
        &["xlsx", "xls", "ods"]
    }

    fn format_name(&self) -> &str {
        "XLSX"
    }
}

/// シートデータをページに変換
fn render_sheet_to_pages(sheet_name: &str, range: &calamine::Range<Data>) -> Vec<Page> {
    let margin = 40.0;
    let page_width = 595.28;
    let page_height = 841.89;
    let usable_width = page_width - margin * 2.0;
    let usable_height = page_height - margin * 2.0;
    let row_height = 20.0;
    let header_height = 30.0;
    let rows_per_page = ((usable_height - header_height) / row_height) as usize;

    let (row_count, col_count) = range.get_size();
    if row_count == 0 || col_count == 0 {
        return vec![Page::a4()];
    }

    let col_width = usable_width / col_count.max(1) as f64;
    let column_widths: Vec<f64> = (0..col_count).map(|_| col_width).collect();

    let mut pages = Vec::new();

    // 全行を収集
    let all_rows: Vec<Vec<String>> = range
        .rows()
        .map(|row| {
            row.iter()
                .map(|cell| match cell {
                    Data::Int(i) => i.to_string(),
                    Data::Float(f) => format!("{:.2}", f),
                    Data::String(s) => s.clone(),
                    Data::Bool(b) => b.to_string(),
                    Data::DateTime(dt) => format!("{}", dt),
                    Data::DateTimeIso(s) => s.clone(),
                    Data::DurationIso(s) => s.clone(),
                    Data::Error(e) => format!("#ERR: {:?}", e),
                    Data::Empty => String::new(),
                })
                .collect()
        })
        .collect();

    for chunk in all_rows.chunks(rows_per_page.max(1)) {
        let mut page = Page::a4();

        // シート名ヘッダー
        page.elements.push(PageElement::Text {
            x: margin,
            y: margin,
            width: usable_width,
            text: format!("📊 {}", sheet_name),
            style: FontStyle {
                font_size: 14.0,
                bold: true,
                ..FontStyle::default()
            },
            align: crate::converter::TextAlign::Left,
        });

        // テーブルデータ
        let table_rows: Vec<Vec<TableCell>> = chunk
            .iter()
            .map(|row| {
                let mut cells: Vec<TableCell> = row
                    .iter()
                    .map(|text| TableCell::new(text))
                    .collect();
                while cells.len() < col_count {
                    cells.push(TableCell::new(""));
                }
                cells
            })
            .collect();

        let table = Table {
            rows: table_rows,
            column_widths: column_widths.clone(),
        };

        page.elements.push(PageElement::TableBlock {
            x: margin,
            y: margin + header_height,
            width: usable_width,
            table,
        });

        pages.push(page);
    }

    if pages.is_empty() {
        pages.push(Page::a4());
    }

    pages
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_format_name() {
        let converter = XlsxConverter::new();
        assert_eq!(converter.format_name(), "XLSX");
        assert_eq!(converter.supported_extensions(), &["xlsx", "xls", "ods"]);
    }
}
