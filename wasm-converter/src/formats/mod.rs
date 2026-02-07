// formats/mod.rs - フォーマットモジュール索引
//
// サポートする全ドキュメントフォーマットのコンバーターを管理します。

pub mod txt;
pub mod csv_conv;
pub mod rtf;
pub mod docx;
pub mod xlsx;
pub mod common_stubs;

use crate::converter::{ConvertError, Document, DocumentConverter};

/// ファイル拡張子に基づいて適切なコンバーターで変換を実行
pub fn convert_by_extension(ext: &str, data: &[u8]) -> Result<Document, ConvertError> {
    match ext.to_lowercase().as_str() {
        "txt" => txt::TxtConverter::new().convert(data),
        "csv" => csv_conv::CsvConverter::new().convert(data),
        "rtf" => rtf::RtfConverter::new().convert(data),
        "docx" => docx::DocxConverter::new().convert(data),
        "xlsx" | "xls" | "ods" => xlsx::XlsxConverter::new().convert(data),
        "doc" => common_stubs::StubConverter::new("DOC", &["doc"]).convert(data),
        "odt" => common_stubs::StubConverter::new("ODT", &["odt"]).convert(data),
        "epub" => common_stubs::StubConverter::new("EPUB", &["epub"]).convert(data),
        "xps" => common_stubs::StubConverter::new("XPS", &["xps"]).convert(data),
        "djvu" | "djv" => common_stubs::StubConverter::new("DjVu", &["djvu", "djv"]).convert(data),
        "ppt" => common_stubs::StubConverter::new("PPT", &["ppt"]).convert(data),
        "pptx" => common_stubs::StubConverter::new("PPTX", &["pptx"]).convert(data),
        "odp" => common_stubs::StubConverter::new("ODP", &["odp"]).convert(data),
        _ => Err(ConvertError::new(
            "unknown",
            &format!("サポートされていないフォーマットです: {}", ext),
        )),
    }
}

/// サポートされているフォーマットの一覧を取得
pub fn supported_formats() -> Vec<(&'static str, &'static [&'static str])> {
    vec![
        ("テキスト", &["txt"][..]),
        ("CSV", &["csv"][..]),
        ("RTF", &["rtf"][..]),
        ("DOCX (Microsoft Word)", &["docx"][..]),
        ("DOC (Microsoft Word 旧形式)", &["doc"][..]),
        ("ODT (OpenDocument Text)", &["odt"][..]),
        ("EPUB (電子書籍)", &["epub"][..]),
        ("XPS", &["xps"][..]),
        ("DjVu", &["djvu", "djv"][..]),
        ("XLSX (Microsoft Excel)", &["xlsx"][..]),
        ("XLS (Microsoft Excel 旧形式)", &["xls"][..]),
        ("ODS (OpenDocument Spreadsheet)", &["ods"][..]),
        ("PPTX (Microsoft PowerPoint)", &["pptx"][..]),
        ("PPT (Microsoft PowerPoint 旧形式)", &["ppt"][..]),
        ("ODP (OpenDocument Presentation)", &["odp"][..]),
    ]
}
