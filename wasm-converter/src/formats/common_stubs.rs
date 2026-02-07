// formats/common_stubs.rs - 未実装フォーマットのスタブ
//
// 複雑なバイナリフォーマット（DOC, PPT, XPS, DjVu等）については、
// 将来の実装に向けたスタブを提供します。

use crate::converter::{
    Color, ConvertError, Document, DocumentConverter, FontStyle, Page, PageElement, TextAlign,
};

/// スタブコンバーター
/// まだ完全に実装されていないフォーマット用のプレースホルダーです。
pub struct StubConverter {
    format_name: String,
    #[allow(dead_code)]
    extensions: Vec<String>,
}

impl StubConverter {
    pub fn new(format_name: &str, extensions: &[&str]) -> Self {
        Self {
            format_name: format_name.to_string(),
            extensions: extensions.iter().map(|e| e.to_string()).collect(),
        }
    }
}

impl DocumentConverter for StubConverter {
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError> {
        let mut doc = Document::new();
        let mut page = Page::a4();
        let margin = 50.0;

        // フォーマット情報ページを生成
        page.elements.push(PageElement::Text {
            x: margin,
            y: margin,
            width: 495.28,
            text: format!("📄 {} ドキュメント", self.format_name),
            style: FontStyle {
                font_size: 18.0,
                bold: true,
                ..FontStyle::default()
            },
            align: TextAlign::Left,
        });

        page.elements.push(PageElement::Line {
            x1: margin,
            y1: margin + 30.0,
            x2: 545.28,
            y2: margin + 30.0,
            width: 1.0,
            color: Color::rgb(100, 100, 100),
        });

        page.elements.push(PageElement::Text {
            x: margin,
            y: margin + 50.0,
            width: 495.28,
            text: format!(
                "このファイルは {} フォーマットです。",
                self.format_name
            ),
            style: FontStyle::default(),
            align: TextAlign::Left,
        });

        page.elements.push(PageElement::Text {
            x: margin,
            y: margin + 70.0,
            width: 495.28,
            text: format!("ファイルサイズ: {} バイト", input.len()),
            style: FontStyle::default(),
            align: TextAlign::Left,
        });

        page.elements.push(PageElement::Text {
            x: margin,
            y: margin + 110.0,
            width: 495.28,
            text: "⚠ このフォーマットの完全な変換は開発中です。".to_string(),
            style: FontStyle {
                color: Color::rgb(200, 100, 0),
                ..FontStyle::default()
            },
            align: TextAlign::Left,
        });

        page.elements.push(PageElement::Text {
            x: margin,
            y: margin + 140.0,
            width: 495.28,
            text: "現在サポートされているフォーマット:".to_string(),
            style: FontStyle {
                bold: true,
                ..FontStyle::default()
            },
            align: TextAlign::Left,
        });

        let supported = [
            "✅ TXT (テキストファイル) - 完全サポート",
            "✅ CSV (カンマ区切り) - 完全サポート",
            "✅ RTF (リッチテキスト) - テキスト抽出",
            "✅ DOCX (Word文書) - テキスト抽出",
            "✅ XLSX/XLS/ODS (スプレッドシート) - テーブル表示",
            &format!("🔧 {} - 開発中", self.format_name),
        ];

        for (i, line) in supported.iter().enumerate() {
            page.elements.push(PageElement::Text {
                x: margin + 20.0,
                y: margin + 165.0 + i as f64 * 20.0,
                width: 475.28,
                text: line.to_string(),
                style: FontStyle {
                    font_size: 9.0,
                    ..FontStyle::default()
                },
                align: TextAlign::Left,
            });
        }

        doc.pages.push(page);
        Ok(doc)
    }

    fn supported_extensions(&self) -> &[&str] {
        // この関数のライフタイムの関係で空スライスを返す
        &[]
    }

    fn format_name(&self) -> &str {
        &self.format_name
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_stub_converter() {
        let converter = StubConverter::new("DOC", &["doc"]);
        let doc = converter.convert(b"dummy data").unwrap();
        assert_eq!(doc.pages.len(), 1);
        assert!(!doc.pages[0].elements.is_empty());
    }
}
