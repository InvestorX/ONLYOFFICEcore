// converter.rs - コア変換トレイトとドキュメントモデル定義
//
// すべてのフォーマットコンバーターが実装すべきトレイトと、
// 中間ドキュメント表現を定義します。

use serde::{Deserialize, Serialize};

/// 変換エラー
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConvertError {
    pub message: String,
    pub format: String,
}

impl ConvertError {
    pub fn new(format: &str, message: &str) -> Self {
        Self {
            message: message.to_string(),
            format: format.to_string(),
        }
    }
}

impl std::fmt::Display for ConvertError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "[{}] {}", self.format, self.message)
    }
}

/// ドキュメント変換トレイト
/// 各フォーマットのコンバーターはこのトレイトを実装します。
pub trait DocumentConverter {
    /// 入力バイト列からドキュメントモデルに変換
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError>;

    /// サポートする拡張子のリスト
    fn supported_extensions(&self) -> &[&str];

    /// フォーマット名
    fn format_name(&self) -> &str;
}

/// 色表現 (RGBA)
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
    pub a: u8,
}

impl Color {
    pub const BLACK: Color = Color {
        r: 0,
        g: 0,
        b: 0,
        a: 255,
    };
    pub const WHITE: Color = Color {
        r: 255,
        g: 255,
        b: 255,
        a: 255,
    };
    pub const fn rgb(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b, a: 255 }
    }
}

/// テキストの水平揃え
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub enum TextAlign {
    Left,
    Center,
    Right,
}

/// フォントスタイル
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontStyle {
    pub font_name: String,
    pub font_size: f64,
    pub bold: bool,
    pub italic: bool,
    pub color: Color,
}

impl Default for FontStyle {
    fn default() -> Self {
        Self {
            font_name: "NotoSansJP".to_string(),
            font_size: 10.0,
            bold: false,
            italic: false,
            color: Color::BLACK,
        }
    }
}

/// テーブルセル
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    pub text: String,
    pub style: FontStyle,
    pub col_span: u32,
    pub row_span: u32,
}

impl TableCell {
    pub fn new(text: &str) -> Self {
        Self {
            text: text.to_string(),
            style: FontStyle::default(),
            col_span: 1,
            row_span: 1,
        }
    }
}

/// テーブル
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Table {
    pub rows: Vec<Vec<TableCell>>,
    pub column_widths: Vec<f64>,
}

/// ページ要素
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum PageElement {
    /// テキストブロック
    Text {
        x: f64,
        y: f64,
        width: f64,
        text: String,
        style: FontStyle,
        align: TextAlign,
    },
    /// 画像
    Image {
        x: f64,
        y: f64,
        width: f64,
        height: f64,
        data: Vec<u8>,
        mime_type: String,
    },
    /// 直線
    Line {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        width: f64,
        color: Color,
    },
    /// 矩形
    Rect {
        x: f64,
        y: f64,
        width: f64,
        height: f64,
        fill: Option<Color>,
        stroke: Option<Color>,
        stroke_width: f64,
    },
    /// テーブル
    TableBlock {
        x: f64,
        y: f64,
        width: f64,
        table: Table,
    },
}

/// ページ
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Page {
    /// 幅（ポイント単位、1pt = 1/72インチ）
    pub width: f64,
    /// 高さ（ポイント単位）
    pub height: f64,
    /// ページ要素
    pub elements: Vec<PageElement>,
}

impl Page {
    /// A4サイズのページを作成
    pub fn a4() -> Self {
        Self {
            width: 595.28,  // 210mm
            height: 841.89, // 297mm
            elements: Vec::new(),
        }
    }

    /// レターサイズのページを作成
    pub fn letter() -> Self {
        Self {
            width: 612.0,
            height: 792.0,
            elements: Vec::new(),
        }
    }
}

/// ドキュメントメタデータ
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Metadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
}

/// 中間ドキュメント表現
/// すべてのフォーマットはまずこの構造に変換され、
/// その後PDFまたは画像に出力されます。
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Document {
    pub pages: Vec<Page>,
    pub metadata: Metadata,
}

impl Document {
    pub fn new() -> Self {
        Self {
            pages: Vec::new(),
            metadata: Metadata::default(),
        }
    }

    /// 単純なテキストドキュメントを作成するヘルパー
    /// テキストをA4ページに自動レイアウトします。
    pub fn from_text_lines(lines: &[String], style: &FontStyle) -> Self {
        let margin = 50.0;
        let line_height = style.font_size * 1.5;
        let page_width = 595.28;
        let page_height = 841.89;
        let usable_height = page_height - margin * 2.0;
        let max_lines_per_page = (usable_height / line_height) as usize;

        let mut doc = Document::new();

        for chunk in lines.chunks(max_lines_per_page.max(1)) {
            let mut page = Page::a4();
            let mut y = margin;

            for line in chunk {
                page.elements.push(PageElement::Text {
                    x: margin,
                    y,
                    width: page_width - margin * 2.0,
                    text: line.clone(),
                    style: style.clone(),
                    align: TextAlign::Left,
                });
                y += line_height;
            }

            doc.pages.push(page);
        }

        if doc.pages.is_empty() {
            doc.pages.push(Page::a4());
        }

        doc
    }
}

/// 出力フォーマット
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub enum OutputFormat {
    /// 単一PDFファイル
    Pdf,
    /// ページごとのPNG画像をZIPにまとめたもの
    ImagesZip,
}

/// 入力ファイルのフォーマットを拡張子から判定
pub fn detect_format(filename: &str) -> Option<&'static str> {
    let ext = filename.rsplit('.').next()?.to_lowercase();
    match ext.as_str() {
        "doc" => Some("doc"),
        "docx" => Some("docx"),
        "odt" => Some("odt"),
        "rtf" => Some("rtf"),
        "txt" => Some("txt"),
        "epub" => Some("epub"),
        "xps" => Some("xps"),
        "djvu" | "djv" => Some("djvu"),
        "xls" => Some("xls"),
        "xlsx" => Some("xlsx"),
        "ods" => Some("ods"),
        "csv" => Some("csv"),
        "ppt" => Some("ppt"),
        "pptx" => Some("pptx"),
        "odp" => Some("odp"),
        _ => None,
    }
}
