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

/// サンプルTXTファイルをPDFに変換し、出力ファイルを検証するテスト
#[test]
fn test_sample_txt_file_to_pdf_output() {
    use std::io::Write;

    // サンプル日本語テキストファイルを作成
    let sample_text = "\
WASM ドキュメントコンバーター テストファイル
========================================

このファイルは、Rust + WebAssembly ドキュメントコンバーターの
動作確認用サンプルファイルです。

日本語テスト：
  漢字：東京都渋谷区
  ひらがな：こんにちは
  カタカナ：コンバーター
  英数字混在：2026年2月7日 Release v0.1.0

Lorem ipsum dolor sit amet, consectetur adipiscing elit.
Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.

--- テスト完了 ---
";

    // TXT → Document → PDF パイプライン
    let doc = formats::convert_by_extension("txt", sample_text.as_bytes()).unwrap();
    assert!(!doc.pages.is_empty(), "ドキュメントにページがありません");

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF-1.4"), "PDFヘッダーが不正です");
    assert!(pdf.len() > 500, "PDFファイルが小さすぎます: {} bytes", pdf.len());

    // PDFの構造を検証
    let pdf_str = String::from_utf8_lossy(&pdf);
    assert!(pdf_str.contains("%%EOF"), "PDFにEOFマーカーがありません");
    assert!(pdf_str.contains("/Type /Catalog"), "PDFにカタログがありません");
    assert!(pdf_str.contains("/Type /Page"), "PDFにページがありません");
    assert!(pdf_str.contains("xref"), "PDFに相互参照テーブルがありません");
    assert!(pdf_str.contains("trailer"), "PDFにトレーラーがありません");

    // 出力ファイルを/tmpに保存して確認可能にする
    let out_path = "/tmp/wasm_converter_test_output.pdf";
    let mut f = std::fs::File::create(out_path).unwrap();
    f.write_all(&pdf).unwrap();
    eprintln!("✅ サンプルPDF出力: {} ({} bytes)", out_path, pdf.len());
}

/// サンプルCSVファイルをPDFに変換し、出力ファイルを検証するテスト
#[test]
fn test_sample_csv_file_to_pdf_output() {
    use std::io::Write;

    let sample_csv = "\
名前,年齢,都市,スコア
田中太郎,30,東京,95
佐藤花子,25,大阪,88
鈴木一郎,35,名古屋,92
高橋美咲,28,福岡,97
伊藤健太,32,札幌,85
";

    let doc = formats::convert_by_extension("csv", sample_csv.as_bytes()).unwrap();
    assert!(!doc.pages.is_empty());

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF-1.4"));
    assert!(pdf.len() > 500);

    let out_path = "/tmp/wasm_converter_test_csv.pdf";
    let mut f = std::fs::File::create(out_path).unwrap();
    f.write_all(&pdf).unwrap();
    eprintln!("✅ サンプルCSV→PDF出力: {} ({} bytes)", out_path, pdf.len());
}

/// サンプルRTFファイルをPDFに変換するテスト
#[test]
fn test_sample_rtf_file_to_pdf_output() {
    use std::io::Write;

    let sample_rtf = r#"{\rtf1\ansi\deff0
{\fonttbl{\f0 Times New Roman;}}
\f0\fs24
This is a sample RTF document.\par
It contains multiple paragraphs.\par
\par
{\b Bold text} and {\i italic text} are supported.\par
Numbers: 1, 2, 3, 4, 5\par
}"#;

    let doc = formats::convert_by_extension("rtf", sample_rtf.as_bytes()).unwrap();
    assert!(!doc.pages.is_empty());

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF-1.4"));

    let out_path = "/tmp/wasm_converter_test_rtf.pdf";
    let mut f = std::fs::File::create(out_path).unwrap();
    f.write_all(&pdf).unwrap();
    eprintln!("✅ サンプルRTF→PDF出力: {} ({} bytes)", out_path, pdf.len());
}

/// サンプルTXTファイルを画像ZIPに変換するテスト
#[test]
fn test_sample_txt_to_images_zip_output() {
    use std::io::Write;

    let sample_text = "画像変換テスト\n\nページ1: こんにちは世界\nRust + WebAssembly";

    let doc = formats::convert_by_extension("txt", sample_text.as_bytes()).unwrap();
    let font_manager = FontManager::new();
    let zip_data = image_renderer::render_to_images_zip(&doc, &font_manager);

    assert!(zip_data.starts_with(b"PK"), "ZIPヘッダーが不正です");
    assert!(zip_data.len() > 100);

    // ZIPの中身を検証
    let cursor = std::io::Cursor::new(&zip_data);
    let archive = zip::ZipArchive::new(cursor).unwrap();
    assert!(archive.len() > 0, "ZIPにファイルがありません");
    eprintln!("✅ ZIP内のページ画像数: {}", archive.len());

    let out_path = "/tmp/wasm_converter_test_pages.zip";
    let mut f = std::fs::File::create(out_path).unwrap();
    f.write_all(&zip_data).unwrap();
    eprintln!("✅ サンプル画像ZIP出力: {} ({} bytes)", out_path, zip_data.len());
}

/// LINE Seed JPフォントの管理テスト
#[test]
fn test_line_seed_jp_font_manager() {
    let mut fm = FontManager::new();

    // 外部フォントとしてLINE Seed JPを追加（ダミーデータで管理機能のみ検証）
    fm.add_font("LINESeedJP-Regular".to_string(), vec![0; 100]);

    let fonts = fm.available_fonts();
    assert!(fonts.contains(&"LINESeedJP-Regular".to_string()));

    // 外部フォントは名前の完全一致で取得可能
    assert!(fm.get_font_data("LINESeedJP-Regular").is_some());

    // 内蔵フォントのキーワードマッチ（embed-fontsなしでは内蔵データは空）
    // "LINE Seed" キーワードは builtin_line_seed_jp() を参照するため、
    // embed-fontsフィーチャー無効時はNoneが返る
    let builtin_result = fm.get_font_data("LINE Seed");
    if cfg!(feature = "embed-fonts") {
        // embed-fonts有効時は内蔵フォントが返る
        assert!(builtin_result.is_some());
    } else {
        // embed-fonts無効時は内蔵データが空なのでNone
        assert!(builtin_result.is_none());
    }
}
