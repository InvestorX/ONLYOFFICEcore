// tests/integration_test.rs - 統合テスト
//
// 各フォーマットの変換フロー全体をテストします。

use wasm_document_converter::converter::{detect_format, Document, FontStyle};
use wasm_document_converter::formats;
use wasm_document_converter::pdf_writer;
use wasm_document_converter::image_renderer;
use wasm_document_converter::font_manager::FontManager;

#[test]
fn test_txt_to_pdf_full_pipeline() {
    let input = "Hello World\nこんにちは世界\n\nRust + WebAssembly".as_bytes();
    let doc = formats::convert_by_extension("txt", input).unwrap();
    assert!(!doc.pages.is_empty());

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF"));
    assert!(pdf.len() > 100);
}

#[test]
fn test_csv_to_pdf_full_pipeline() {
    let input = b"Name,Score\nTaro,100\nHanako,95";
    let doc = formats::convert_by_extension("csv", input).unwrap();
    assert!(!doc.pages.is_empty());

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_txt_to_images_zip_full_pipeline() {
    let input = "Test Document\nPage 1 content".as_bytes();
    let doc = formats::convert_by_extension("txt", input).unwrap();

    let font_manager = FontManager::new();
    let zip_data = image_renderer::render_to_images_zip(&doc, &font_manager);
    assert!(zip_data.starts_with(b"PK"));
    assert!(zip_data.len() > 100);
}

#[test]
fn test_rtf_to_pdf_full_pipeline() {
    let input = br#"{\rtf1\ansi Hello RTF World}"#;
    let doc = formats::convert_by_extension("rtf", input).unwrap();
    assert!(!doc.pages.is_empty());

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_stub_format_produces_info_page() {
    let input = b"dummy binary data";
    let doc = formats::convert_by_extension("doc", input).unwrap();
    assert_eq!(doc.pages.len(), 1);
    // ページに要素が含まれていること（情報ページ）
    assert!(!doc.pages[0].elements.is_empty());
}

#[test]
fn test_detect_all_formats() {
    let test_cases = vec![
        ("file.doc", Some("doc")),
        ("file.docx", Some("docx")),
        ("file.odt", Some("odt")),
        ("file.rtf", Some("rtf")),
        ("file.txt", Some("txt")),
        ("file.epub", Some("epub")),
        ("file.xps", Some("xps")),
        ("file.djvu", Some("djvu")),
        ("file.xls", Some("xls")),
        ("file.xlsx", Some("xlsx")),
        ("file.ods", Some("ods")),
        ("file.csv", Some("csv")),
        ("file.ppt", Some("ppt")),
        ("file.pptx", Some("pptx")),
        ("file.odp", Some("odp")),
        ("file.unknown", None),
    ];

    for (filename, expected) in test_cases {
        assert_eq!(detect_format(filename), expected, "Failed for: {}", filename);
    }
}

#[test]
fn test_supported_formats_count() {
    let formats = formats::supported_formats();
    // 最低15フォーマットがサポートされていること
    assert!(formats.len() >= 15, "Expected >= 15 formats, got {}", formats.len());
}

#[test]
fn test_document_from_text_lines_pagination() {
    let style = FontStyle::default();
    // 大量の行を生成してページ分割をテスト
    let lines: Vec<String> = (0..200).map(|i| format!("Line {}", i)).collect();
    let doc = Document::from_text_lines(&lines, &style);
    // 複数ページに分割されること
    assert!(doc.pages.len() > 1, "Expected multiple pages for 200 lines");
}

#[test]
fn test_unsupported_format_error() {
    let result = formats::convert_by_extension("xyz", b"data");
    assert!(result.is_err());
}

#[test]
fn test_japanese_text_in_pdf() {
    let input = "日本語テスト\n漢字、ひらがな、カタカナ\nRust WebAssembly変換".as_bytes();
    let doc = formats::convert_by_extension("txt", input).unwrap();
    let pdf = pdf_writer::render_to_pdf(&doc);

    // PDFヘッダーが正しいこと
    assert!(pdf.starts_with(b"%PDF-1.4"));
    // EOF マーカーが含まれること
    let pdf_str = String::from_utf8_lossy(&pdf);
    assert!(pdf_str.contains("%%EOF"));
}
