// formats/csv_conv.rs - CSV変換モジュール
//
// CSVファイルを読み込み、テーブル形式でドキュメントモデルに変換します。

use crate::converter::{
    ConvertError, Document, DocumentConverter, Metadata, Page, PageElement,
    Table, TableCell,
};

/// CSVコンバーター
pub struct CsvConverter;

impl CsvConverter {
    pub fn new() -> Self {
        Self
    }
}

impl DocumentConverter for CsvConverter {
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError> {
        let text = String::from_utf8_lossy(input);
        let mut reader = csv::ReaderBuilder::new()
            .has_headers(false)
            .flexible(true)
            .from_reader(text.as_bytes());

        let mut rows: Vec<Vec<String>> = Vec::new();
        let mut max_cols = 0;

        for result in reader.records() {
            let record = result
                .map_err(|e| ConvertError::new("CSV", &format!("CSVパースエラー: {}", e)))?;
            let row: Vec<String> = record.iter().map(|s| s.to_string()).collect();
            max_cols = max_cols.max(row.len());
            rows.push(row);
        }

        if rows.is_empty() {
            let mut doc = Document::new();
            doc.pages.push(Page::a4());
            return Ok(doc);
        }

        // テーブルレイアウト
        let margin = 40.0;
        let page_width = 595.28;
        let page_height = 841.89;
        let usable_width = page_width - margin * 2.0;
        let row_height = 22.0;
        let header_height = 26.0;
        let usable_height = page_height - margin * 2.0;

        let col_width = usable_width / max_cols.max(1) as f64;
        let column_widths: Vec<f64> = (0..max_cols).map(|_| col_width).collect();

        // ページごとに行を分割
        let rows_per_page = ((usable_height - header_height) / row_height) as usize;

        let mut doc = Document::new();
        doc.metadata = Metadata {
            title: Some("CSV Document".to_string()),
            ..Default::default()
        };

        for (chunk_idx, chunk) in rows.chunks(rows_per_page.max(1)).enumerate() {
            let mut page = Page::a4();

            // テーブルセルを構築
            let table_rows: Vec<Vec<TableCell>> = chunk
                .iter()
                .enumerate()
                .map(|(row_idx, row)| {
                    let mut cells: Vec<TableCell> = row
                        .iter()
                        .map(|cell_text| {
                            let mut cell = TableCell::new(cell_text);
                            // 最初のページの最初の行はヘッダースタイル
                            if chunk_idx == 0 && row_idx == 0 {
                                cell.style.bold = true;
                                cell.style.font_size = 11.0;
                            }
                            cell
                        })
                        .collect();
                    // 列数を揃える
                    while cells.len() < max_cols {
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
                y: margin,
                width: usable_width,
                table,
            });

            doc.pages.push(page);
        }

        if doc.pages.is_empty() {
            doc.pages.push(Page::a4());
        }

        Ok(doc)
    }

    fn supported_extensions(&self) -> &[&str] {
        &["csv"]
    }

    fn format_name(&self) -> &str {
        "CSV"
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_csv() {
        let input = b"Name,Age,City\nAlice,30,Tokyo\nBob,25,Osaka";
        let converter = CsvConverter::new();
        let doc = converter.convert(input).unwrap();
        assert!(!doc.pages.is_empty());
    }

    #[test]
    fn test_empty_csv() {
        let input = b"";
        let converter = CsvConverter::new();
        let doc = converter.convert(input).unwrap();
        assert_eq!(doc.pages.len(), 1);
    }

    #[test]
    fn test_japanese_csv() {
        let input = "名前,年齢,都市\n太郎,30,東京\n花子,25,大阪".as_bytes();
        let converter = CsvConverter::new();
        let doc = converter.convert(input).unwrap();
        assert!(!doc.pages.is_empty());
    }
}
