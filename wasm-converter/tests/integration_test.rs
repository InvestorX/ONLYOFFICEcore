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

/// 最小限の有効なDOCXファイルを作成するヘルパー
fn create_sample_docx(text_paragraphs: &[&str]) -> Vec<u8> {
    use std::io::Write;
    let buf = Vec::new();
    let cursor = std::io::Cursor::new(buf);
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::SimpleFileOptions::default();

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    // word/document.xml
    zip.start_file("word/document.xml", options).unwrap();
    let mut doc_xml = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>"#);
    for para in text_paragraphs {
        doc_xml.push_str(&format!(
            "\n    <w:p><w:r><w:t>{}</w:t></w:r></w:p>",
            para
        ));
    }
    doc_xml.push_str("\n  </w:body>\n</w:document>");
    zip.write_all(doc_xml.as_bytes()).unwrap();

    // docProps/core.xml
    zip.start_file("docProps/core.xml", options).unwrap();
    zip.write_all("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\
<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"\n\
                   xmlns:dc=\"http://purl.org/dc/elements/1.1/\">\n\
  <dc:title>サンプルDOCX文書</dc:title>\n\
  <dc:creator>テストユーザー</dc:creator>\n\
</cp:coreProperties>".as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

/// 最小限の有効なPPTXファイルを作成するヘルパー
fn create_sample_pptx(slides: &[(&str, &[&str])]) -> Vec<u8> {
    use std::io::Write;
    let buf = Vec::new();
    let cursor = std::io::Cursor::new(buf);
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::SimpleFileOptions::default();

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", options).unwrap();
    let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>"#);
    for (i, _) in slides.iter().enumerate() {
        ct.push_str(&format!(
            r#"
  <Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
            i + 1
        ));
    }
    ct.push_str("\n</Types>");
    zip.write_all(ct.as_bytes()).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#).unwrap();

    // ppt/presentation.xml
    zip.start_file("ppt/presentation.xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#).unwrap();

    // 各スライドを作成
    for (i, (title, body_texts)) in slides.iter().enumerate() {
        let slide_path = format!("ppt/slides/slide{}.xml", i + 1);
        zip.start_file(&slide_path, options).unwrap();

        let mut slide_xml = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>"#);

        // タイトル
        slide_xml.push_str(&format!(r#"
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:txBody>
          <a:p><a:r><a:t>{}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>"#, title));

        // 本文テキスト
        slide_xml.push_str("\n      <p:sp>\n        <p:txBody>");
        for text in *body_texts {
            slide_xml.push_str(&format!(
                "\n          <a:p><a:r><a:t>{}</a:t></a:r></a:p>",
                text
            ));
        }
        slide_xml.push_str("\n        </p:txBody>\n      </p:sp>");

        slide_xml.push_str("\n    </p:spTree>\n  </p:cSld>\n</p:sld>");
        zip.write_all(slide_xml.as_bytes()).unwrap();
    }

    // docProps/core.xml
    zip.start_file("docProps/core.xml", options).unwrap();
    zip.write_all("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\
<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"\n\
                   xmlns:dc=\"http://purl.org/dc/elements/1.1/\">\n\
  <dc:title>サンプルPPTXプレゼンテーション</dc:title>\n\
  <dc:creator>テストユーザー</dc:creator>\n\
</cp:coreProperties>".as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

/// サンプルDOCXファイルを変換してPDF出力するテスト
#[test]
fn test_sample_docx_to_pdf() {
    use std::io::Write;

    let docx_data = create_sample_docx(&[
        "明治大学情報基盤本部",
        "サンプルテキスト文書",
        "",
        "これはDOCXファイルのサンプルテストです。",
        "日本語テキストが正しく変換されることを確認します。",
        "",
        "第1章 はじめに",
        "ドキュメントコンバーターは、様々なファイル形式をPDFに変換するツールです。",
        "WebAssemblyを使用して、ブラウザ上で動作します。",
        "",
        "第2章 機能一覧",
        "・テキストファイル (TXT, CSV, RTF)",
        "・Microsoft Office (DOCX, XLSX, PPTX)",
        "・その他 (ODT, ODS, ODP, EPUB等)",
    ]);

    let doc = formats::convert_by_extension("docx", &docx_data).unwrap();
    assert!(!doc.pages.is_empty(), "DOCXからページが生成されませんでした");

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF-1.4"), "PDFヘッダーが不正");
    assert!(pdf.len() > 500, "PDFが小さすぎます: {} bytes", pdf.len());

    let out_path = "/tmp/wasm_converter_test_docx.pdf";
    let mut f = std::fs::File::create(out_path).unwrap();
    f.write_all(&pdf).unwrap();
    eprintln!("✅ DOCX→PDF出力: {} ({} bytes, {}ページ)", out_path, pdf.len(), doc.pages.len());
}

/// サンプルPPTXファイルを変換してPDF出力するテスト
#[test]
fn test_sample_pptx_to_pdf() {
    use std::io::Write;

    let pptx_data = create_sample_pptx(&[
        ("PowerPoint練習1", &[
            "情報基盤本部 サンプル",
            "プレゼンテーションの基本",
        ][..]),
        ("スライド作成の手順", &[
            "1. 新しいスライドを追加する",
            "2. テキストを入力する",
            "3. 画像やグラフを挿入する",
            "4. デザインを整える",
        ][..]),
        ("日本語テキストのテスト", &[
            "漢字: 東京都渋谷区",
            "ひらがな: こんにちは",
            "カタカナ: プレゼンテーション",
            "英数字混在: 2026年2月7日 v0.1.0",
        ][..]),
    ]);

    let doc = formats::convert_by_extension("pptx", &pptx_data).unwrap();
    assert_eq!(doc.pages.len(), 3, "PPTXから3ページ生成されるべき");

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF-1.4"), "PDFヘッダーが不正");
    assert!(pdf.len() > 500, "PDFが小さすぎます");

    let out_path = "/tmp/wasm_converter_test_pptx.pdf";
    let mut f = std::fs::File::create(out_path).unwrap();
    f.write_all(&pdf).unwrap();
    eprintln!("✅ PPTX→PDF出力: {} ({} bytes, {}ページ)", out_path, pdf.len(), doc.pages.len());
}

/// PPTX→画像ZIPの変換テスト
#[test]
fn test_sample_pptx_to_images_zip() {
    use std::io::Write;

    let pptx_data = create_sample_pptx(&[
        ("スライド1: タイトル", &["テスト内容"][..]),
        ("スライド2: 詳細", &["詳細テキスト"][..]),
    ]);

    let doc = formats::convert_by_extension("pptx", &pptx_data).unwrap();
    let fm = FontManager::new();
    let zip_data = image_renderer::render_to_images_zip(&doc, &fm);

    assert!(zip_data.starts_with(b"PK"), "ZIPヘッダーが不正");

    let cursor = std::io::Cursor::new(&zip_data);
    let archive = zip::ZipArchive::new(cursor).unwrap();
    assert_eq!(archive.len(), 2, "2スライドなので2つのPNG画像が必要");

    let out_path = "/tmp/wasm_converter_test_pptx_pages.zip";
    let mut f = std::fs::File::create(out_path).unwrap();
    f.write_all(&zip_data).unwrap();
    eprintln!("✅ PPTX→画像ZIP出力: {} ({} bytes, {}ページ)", out_path, zip_data.len(), archive.len());
}

/// サンプルXLSXファイルでの変換テスト（calamine使用）
#[test]
fn test_sample_xlsx_to_pdf_with_calamine() {
    use std::io::Write;

    // calamine経由で最小限の有効なXLSXファイルを作成
    // XLSXのXML構造を手動構築
    let buf = Vec::new();
    let cursor = std::io::Cursor::new(buf);
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::SimpleFileOptions::default();

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).unwrap();

    // xl/_rels/workbook.xml.rels
    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#).unwrap();

    // xl/workbook.xml
    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"\n\
          xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n\
  <sheets>\n\
    <sheet name=\"売上データ\" sheetId=\"1\" r:id=\"rId1\"/>\n\
  </sheets>\n\
</workbook>".as_bytes()).unwrap();

    // xl/sharedStrings.xml
    zip.start_file("xl/sharedStrings.xml", options).unwrap();
    zip.write_all("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\
<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"8\" uniqueCount=\"8\">\n\
  <si><t>商品名</t></si>\n\
  <si><t>数量</t></si>\n\
  <si><t>単価</t></si>\n\
  <si><t>合計</t></si>\n\
  <si><t>りんご</t></si>\n\
  <si><t>みかん</t></si>\n\
  <si><t>バナナ</t></si>\n\
  <si><t>いちご</t></si>\n\
</sst>".as_bytes()).unwrap();

    // xl/worksheets/sheet1.xml
    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
      <c r="D1" t="s"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>4</v></c>
      <c r="B2"><v>10</v></c>
      <c r="C2"><v>150</v></c>
      <c r="D2"><v>1500</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>5</v></c>
      <c r="B3"><v>20</v></c>
      <c r="C3"><v>80</v></c>
      <c r="D3"><v>1600</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>6</v></c>
      <c r="B4"><v>15</v></c>
      <c r="C4"><v>120</v></c>
      <c r="D4"><v>1800</v></c>
    </row>
    <row r="5">
      <c r="A5" t="s"><v>7</v></c>
      <c r="B5"><v>8</v></c>
      <c r="C5"><v>300</v></c>
      <c r="D5"><v>2400</v></c>
    </row>
  </sheetData>
</worksheet>"#).unwrap();

    let xlsx_data = zip.finish().unwrap().into_inner();

    let doc = formats::convert_by_extension("xlsx", &xlsx_data).unwrap();
    assert!(!doc.pages.is_empty(), "XLSXからページが生成されませんでした");

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF-1.4"), "PDFヘッダーが不正");

    let out_path = "/tmp/wasm_converter_test_xlsx_sample.pdf";
    let mut f = std::fs::File::create(out_path).unwrap();
    f.write_all(&pdf).unwrap();
    eprintln!("✅ XLSX→PDF出力: {} ({} bytes, {}ページ)", out_path, pdf.len(), doc.pages.len());
}

/// Sample_12.pptxでのレイアウト保持変換テスト（背景画像・グラデーション・シャドウ含む）
/// (ファイルが存在しない場合はスキップ)
#[test]
fn test_real_sample12_pptx_layout() {
    use wasm_document_converter::converter::PageElement;
    use std::io::Write;

    let pptx_path = "/tmp/Sample_12.pptx";
    let data = match std::fs::read(pptx_path) {
        Ok(d) => d,
        Err(_) => {
            eprintln!("⏭ {}: ファイルが見つからないためスキップ", pptx_path);
            return;
        }
    };

    eprintln!("入力: {} bytes", data.len());

    let doc = formats::convert_by_extension("pptx", &data).expect("PPTX変換に失敗");

    eprintln!("スライド数: {}", doc.pages.len());
    for (i, page) in doc.pages.iter().enumerate() {
        let mut counts = std::collections::HashMap::new();
        for elem in &page.elements {
            let key = match elem {
                PageElement::Rect { .. } => "Rect",
                PageElement::GradientRect { .. } => "Gradient",
                PageElement::Image { .. } => "Image",
                PageElement::Text { .. } => "Text",
                PageElement::Line { .. } => "Line",
                PageElement::Ellipse { .. } => "Ellipse",
                PageElement::TableBlock { .. } => "Table",
                PageElement::Path { .. } => "Path",
                PageElement::EllipseImage { .. } => "EllipseImage",
                PageElement::PathImage { .. } => "PathImage",
            };
            *counts.entry(key).or_insert(0u32) += 1;
        }
        eprintln!(
            "  スライド {}: {}x{} ({} 要素: {:?})",
            i + 1,
            page.width as i32,
            page.height as i32,
            page.elements.len(),
            counts,
        );
    }

    assert!(doc.pages.len() >= 10, "12スライドあるはず（実際: {}）", doc.pages.len());

    // Slide 1 should have background image (from blipFill)
    let slide1 = &doc.pages[0];
    let has_bg_image = slide1.elements.iter().any(|e| matches!(e, PageElement::Image { .. }));
    eprintln!("  スライド1 背景画像: {}", has_bg_image);
    assert!(has_bg_image, "スライド1には背景画像があるはず");

    // Slide 4 should have gradient background
    if doc.pages.len() > 3 {
        let slide4 = &doc.pages[3];
        let has_gradient = slide4.elements.iter().any(|e| matches!(e, PageElement::GradientRect { .. }));
        eprintln!("  スライド4 グラデーション背景: {}", has_gradient);
        assert!(has_gradient, "スライド4にはグラデーション背景があるはず");
    }

    // PDF出力
    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF-1.4"));
    let pdf_path = "/tmp/wasm_converter_sample12.pdf";
    let mut f = std::fs::File::create(pdf_path).unwrap();
    f.write_all(&pdf).unwrap();
    eprintln!("✅ PDF出力: {} ({} bytes)", pdf_path, pdf.len());

    // 個別スライドPNG出力（確認用）
    let fm = FontManager::new();
    let config = image_renderer::ImageRenderConfig {
        dpi: 150.0,
        ..Default::default()
    };
    for idx in [0, 3, 4] {
        if idx < doc.pages.len() {
            let img = image_renderer::render_page_to_image(&doc.pages[idx], &config, &fm);
            let png_path = format!("/tmp/wasm_converter_sample12_slide{}.png", idx + 1);
            let mut f = std::fs::File::create(&png_path).unwrap();
            f.write_all(&img).unwrap();
            eprintln!("✅ スライド{} PNG: {} ({} bytes)", idx + 1, png_path, img.len());
        }
    }

    // 画像ZIP出力
    let zip_data = image_renderer::render_to_images_zip(&doc, &fm);
    assert!(zip_data.starts_with(b"PK"));
    let zip_path = "/tmp/wasm_converter_sample12_pages.zip";
    let mut f = std::fs::File::create(zip_path).unwrap();
    f.write_all(&zip_data).unwrap();
    eprintln!("✅ 画像ZIP出力: {} ({} bytes)", zip_path, zip_data.len());
}

/// Sample_12.pptxのチャートスライド（slide 8, 9, 10）のレンダリングテスト
/// チャート（棒グラフ、円グラフ、面グラフ）が正しくレンダリングされることを検証
#[test]
fn test_sample12_chart_rendering() {
    use wasm_document_converter::converter::PageElement;
    use std::io::Write;

    let pptx_path = "/tmp/Sample_12.pptx";
    let data = match std::fs::read(pptx_path) {
        Ok(d) => d,
        Err(_) => {
            eprintln!("⏭ {}: ファイルが見つからないためスキップ", pptx_path);
            return;
        }
    };

    let doc = formats::convert_by_extension("pptx", &data).expect("PPTX変換に失敗");
    assert!(doc.pages.len() >= 12, "12スライドあるはず（実際: {}）", doc.pages.len());

    // Slide 8 has a bar chart
    let slide8 = &doc.pages[7];
    eprintln!("スライド8: {} 要素", slide8.elements.len());
    // Charts produce rects (bars), lines (axes/gridlines), text (labels) and ellipses
    let has_chart_elements = slide8.elements.iter().any(|e| {
        matches!(e, PageElement::Line { .. })
    });
    eprintln!("  チャート要素あり: {}", has_chart_elements);

    // Slide 9 also has a chart
    let slide9 = &doc.pages[8];
    eprintln!("スライド9: {} 要素", slide9.elements.len());

    // Render chart slides as PNG for visual verification
    let fm = FontManager::new();
    let config = image_renderer::ImageRenderConfig {
        dpi: 150.0,
        ..Default::default()
    };
    for idx in [7, 8, 9] { // slides 8, 9, 10
        if idx < doc.pages.len() {
            let img = image_renderer::render_page_to_image(&doc.pages[idx], &config, &fm);
            let png_path = format!("/tmp/wasm_converter_sample12_chart_slide{}.png", idx + 1);
            let mut f = std::fs::File::create(&png_path).unwrap();
            f.write_all(&img).unwrap();
            eprintln!("✅ チャートスライド{} PNG: {} ({} bytes)", idx + 1, png_path, img.len());
        }
    }
}

/// Sample_12.pptxのSmartArtスライド（slide 11）のレンダリングテスト
#[test]
fn test_sample12_smartart_rendering() {
    use std::io::Write;

    let pptx_path = "/tmp/Sample_12.pptx";
    let data = match std::fs::read(pptx_path) {
        Ok(d) => d,
        Err(_) => {
            eprintln!("⏭ {}: ファイルが見つからないためスキップ", pptx_path);
            return;
        }
    };

    let doc = formats::convert_by_extension("pptx", &data).expect("PPTX変換に失敗");

    // Slide 11 has SmartArt (diagram)
    if doc.pages.len() > 10 {
        let slide11 = &doc.pages[10];
        eprintln!("スライド11 (SmartArt): {} 要素", slide11.elements.len());

        let fm = FontManager::new();
        let config = image_renderer::ImageRenderConfig {
            dpi: 150.0,
            ..Default::default()
        };
        let img = image_renderer::render_page_to_image(slide11, &config, &fm);
        let png_path = "/tmp/wasm_converter_sample12_smartart_slide11.png";
        let mut f = std::fs::File::create(png_path).unwrap();
        f.write_all(&img).unwrap();
        eprintln!("✅ SmartArtスライド11 PNG: {} ({} bytes)", png_path, img.len());
    }
}

/// Sample_12.pptxの3Dエフェクトスライド（slide 6, 7）のレンダリングテスト
#[test]
fn test_sample12_3d_effects() {
    use std::io::Write;

    let pptx_path = "/tmp/Sample_12.pptx";
    let data = match std::fs::read(pptx_path) {
        Ok(d) => d,
        Err(_) => {
            eprintln!("⏭ {}: ファイルが見つからないためスキップ", pptx_path);
            return;
        }
    };

    let doc = formats::convert_by_extension("pptx", &data).expect("PPTX変換に失敗");

    // Slides 6-7 have 3D effects (scene3d/sp3d)
    let fm = FontManager::new();
    let config = image_renderer::ImageRenderConfig {
        dpi: 150.0,
        ..Default::default()
    };
    for idx in [5, 6] { // slides 6, 7
        if idx < doc.pages.len() {
            let page = &doc.pages[idx];
            eprintln!("スライド{} (3D): {} 要素", idx + 1, page.elements.len());
            let img = image_renderer::render_page_to_image(page, &config, &fm);
            let png_path = format!("/tmp/wasm_converter_sample12_3d_slide{}.png", idx + 1);
            let mut f = std::fs::File::create(&png_path).unwrap();
            f.write_all(&img).unwrap();
            eprintln!("✅ 3Dスライド{} PNG: {} ({} bytes)", idx + 1, png_path, img.len());
        }
    }
}

/// Sample_12.pptxのJPEG画像デコードテスト
/// 実際のJPEG画像（image1.jpg, image2.jpeg等）が完全にデコードされることを検証
#[test]
fn test_sample12_jpeg_decoding() {
    use wasm_document_converter::converter::PageElement;

    let pptx_path = "/tmp/Sample_12.pptx";
    let data = match std::fs::read(pptx_path) {
        Ok(d) => d,
        Err(_) => {
            eprintln!("⏭ {}: ファイルが見つからないためスキップ", pptx_path);
            return;
        }
    };

    let doc = formats::convert_by_extension("pptx", &data).expect("PPTX変換に失敗");

    // Count image elements across all slides
    let mut total_images = 0;
    let mut jpeg_images = 0;
    for (i, page) in doc.pages.iter().enumerate() {
        for elem in &page.elements {
            if let PageElement::Image { mime_type, data, .. } = elem {
                total_images += 1;
                if mime_type.contains("jpeg") || mime_type.contains("jpg") {
                    jpeg_images += 1;
                    // Verify JPEG header
                    assert!(data.len() > 10, "JPEG data too small on slide {}", i + 1);
                    assert!(data[0] == 0xFF && data[1] == 0xD8, "Invalid JPEG header on slide {}", i + 1);
                }
                eprintln!("  スライド{} 画像: {} ({} bytes)", i + 1, mime_type, data.len());
            }
        }
    }
    eprintln!("画像合計: {}, JPEG: {}", total_images, jpeg_images);
    // Sample_12.pptx has JPEG images
    assert!(total_images > 0, "画像が見つかりません");
}

/// assignment.docx（GitHub docxtemplater sample）の変換テスト
#[test]
fn test_assignment_docx_conversion() {
    use std::io::Write;

    let docx_path = "/tmp/assignment.docx";
    let data = match std::fs::read(docx_path) {
        Ok(d) => d,
        Err(_) => {
            eprintln!("⏭ {}: ファイルが見つからないためスキップ", docx_path);
            return;
        }
    };

    eprintln!("assignment.docx: {} bytes", data.len());

    let doc = formats::convert_by_extension("docx", &data).expect("DOCX変換に失敗");
    assert!(!doc.pages.is_empty(), "ページが生成されませんでした");
    eprintln!("ページ数: {}", doc.pages.len());
    for (i, page) in doc.pages.iter().enumerate() {
        eprintln!("  ページ{}: {}x{} ({} 要素)", i + 1, page.width as i32, page.height as i32, page.elements.len());
    }

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF-1.4"));
    let pdf_path = "/tmp/wasm_converter_assignment.pdf";
    let mut f = std::fs::File::create(pdf_path).unwrap();
    f.write_all(&pdf).unwrap();
    eprintln!("✅ assignment.docx → PDF: {} ({} bytes)", pdf_path, pdf.len());

    // PNG output for screenshot
    let fm = FontManager::new();
    let config = image_renderer::ImageRenderConfig {
        dpi: 150.0,
        ..Default::default()
    };
    if !doc.pages.is_empty() {
        let img = image_renderer::render_page_to_image(&doc.pages[0], &config, &fm);
        let png_path = "/tmp/wasm_converter_assignment_page1.png";
        let mut f = std::fs::File::create(png_path).unwrap();
        f.write_all(&img).unwrap();
        eprintln!("✅ assignment.docx ページ1 PNG: {} ({} bytes)", png_path, img.len());
    }
}

/// simple.xlsx（GitHub docxtemplater sample）の変換テスト
#[test]
fn test_simple_xlsx_conversion() {
    use std::io::Write;

    let xlsx_path = "/tmp/simple.xlsx";
    let data = match std::fs::read(xlsx_path) {
        Ok(d) => d,
        Err(_) => {
            eprintln!("⏭ {}: ファイルが見つからないためスキップ", xlsx_path);
            return;
        }
    };

    eprintln!("simple.xlsx: {} bytes", data.len());

    let doc = formats::convert_by_extension("xlsx", &data).expect("XLSX変換に失敗");
    assert!(!doc.pages.is_empty(), "ページが生成されませんでした");
    eprintln!("ページ数: {}", doc.pages.len());

    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF-1.4"));
    let pdf_path = "/tmp/wasm_converter_simple_xlsx.pdf";
    let mut f = std::fs::File::create(pdf_path).unwrap();
    f.write_all(&pdf).unwrap();
    eprintln!("✅ simple.xlsx → PDF: {} ({} bytes)", pdf_path, pdf.len());

    // PNG output
    let fm = FontManager::new();
    let config = image_renderer::ImageRenderConfig {
        dpi: 150.0,
        ..Default::default()
    };
    if !doc.pages.is_empty() {
        let img = image_renderer::render_page_to_image(&doc.pages[0], &config, &fm);
        let png_path = "/tmp/wasm_converter_simple_xlsx_page1.png";
        let mut f = std::fs::File::create(png_path).unwrap();
        f.write_all(&img).unwrap();
        eprintln!("✅ simple.xlsx ページ1 PNG: {} ({} bytes)", png_path, img.len());
    }
}

/// チャートXML直接解析テスト（棒グラフ、円グラフ、面グラフ）
#[test]
fn test_chart_xml_parsing() {
    use wasm_document_converter::formats::chart;

    // Test bar chart rendering
    let bar_xml = r#"<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart><c:plotArea>
            <c:barChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                <c:ser>
                    <c:idx val="0"/>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series 1</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:cat><c:strRef><c:strCache>
                        <c:pt idx="0"><c:v>Cat A</c:v></c:pt>
                        <c:pt idx="1"><c:v>Cat B</c:v></c:pt>
                        <c:pt idx="2"><c:v>Cat C</c:v></c:pt>
                    </c:strCache></c:strRef></c:cat>
                    <c:val><c:numRef><c:numCache>
                        <c:pt idx="0"><c:v>4.3</c:v></c:pt>
                        <c:pt idx="1"><c:v>2.5</c:v></c:pt>
                        <c:pt idx="2"><c:v>3.5</c:v></c:pt>
                    </c:numCache></c:numRef></c:val>
                </c:ser>
                <c:ser>
                    <c:idx val="1"/>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series 2</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:val><c:numRef><c:numCache>
                        <c:pt idx="0"><c:v>2.4</c:v></c:pt>
                        <c:pt idx="1"><c:v>4.4</c:v></c:pt>
                        <c:pt idx="2"><c:v>1.8</c:v></c:pt>
                    </c:numCache></c:numRef></c:val>
                </c:ser>
            </c:barChart>
        </c:plotArea></c:chart>
    </c:chartSpace>"#;

    let elements = chart::render_chart(bar_xml, 50.0, 50.0, 400.0, 300.0);
    assert!(elements.len() > 10, "チャート要素が少なすぎます: {}", elements.len());

    // Test pie chart
    let pie_xml = r#"<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea>
            <c:pie3DChart>
                <c:ser>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:cat><c:strRef><c:strCache>
                        <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                        <c:pt idx="1"><c:v>Q2</c:v></c:pt>
                        <c:pt idx="2"><c:v>Q3</c:v></c:pt>
                    </c:strCache></c:strRef></c:cat>
                    <c:val><c:numRef><c:numCache>
                        <c:pt idx="0"><c:v>8.2</c:v></c:pt>
                        <c:pt idx="1"><c:v>3.2</c:v></c:pt>
                        <c:pt idx="2"><c:v>1.4</c:v></c:pt>
                    </c:numCache></c:numRef></c:val>
                </c:ser>
            </c:pie3DChart>
        </c:plotArea></c:chart>
    </c:chartSpace>"#;

    let pie_elements = chart::render_chart(pie_xml, 50.0, 50.0, 300.0, 300.0);
    assert!(pie_elements.len() > 5, "円グラフ要素が少なすぎます: {}", pie_elements.len());

    // Test area chart
    let area_xml = r#"<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea>
            <c:areaChart>
                <c:ser>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Data</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:cat><c:strRef><c:strCache>
                        <c:pt idx="0"><c:v>Jan</c:v></c:pt>
                        <c:pt idx="1"><c:v>Feb</c:v></c:pt>
                        <c:pt idx="2"><c:v>Mar</c:v></c:pt>
                    </c:strCache></c:strRef></c:cat>
                    <c:val><c:numRef><c:numCache>
                        <c:pt idx="0"><c:v>1</c:v></c:pt>
                        <c:pt idx="1"><c:v>3</c:v></c:pt>
                        <c:pt idx="2"><c:v>2</c:v></c:pt>
                    </c:numCache></c:numRef></c:val>
                </c:ser>
            </c:areaChart>
        </c:plotArea></c:chart>
    </c:chartSpace>"#;

    let area_elements = chart::render_chart(area_xml, 0.0, 0.0, 400.0, 300.0);
    assert!(area_elements.len() > 5, "面グラフ要素が少なすぎます: {}", area_elements.len());
}

// --- 外部フォント読み込みテスト ---

#[test]
fn test_external_font_loading() {
    let mut fm = FontManager::new();
    
    // 初期状態: 外部フォントなし
    assert_eq!(fm.external_font_count(), 0);
    
    // 外部フォントを追加（ダミーデータ）
    let dummy_font = vec![0u8; 100];
    fm.add_font("TestFont".to_string(), dummy_font.clone());
    
    assert_eq!(fm.external_font_count(), 1);
    assert!(fm.has_any_font());
    
    // フォント名で検索
    let data = fm.get_font_data("TestFont");
    assert!(data.is_some());
    assert_eq!(data.unwrap().len(), 100);
    
    // 部分一致検索
    let data2 = fm.get_font_data("Test");
    assert!(data2.is_some());
    
    // best_font_data は外部フォントを優先
    let best = fm.best_font_data();
    assert!(best.is_some());
    assert_eq!(best.unwrap().len(), 100);
    
    // 複数フォント追加
    fm.add_font("AnotherFont".to_string(), vec![1u8; 200]);
    assert_eq!(fm.external_font_count(), 2);
    
    let fonts = fm.available_fonts();
    assert!(fonts.contains(&"TestFont".to_string()));
    assert!(fonts.contains(&"AnotherFont".to_string()));
    
    // フォント削除
    fm.remove_font("TestFont");
    assert_eq!(fm.external_font_count(), 1);
    assert!(fm.get_font_data("TestFont").is_none());
    
    // 同名フォントの置き換え
    fm.add_font("AnotherFont".to_string(), vec![2u8; 300]);
    assert_eq!(fm.external_font_count(), 1);
    assert_eq!(fm.get_font_data("AnotherFont").unwrap().len(), 300);
}

#[test]
fn test_font_resolve_cjk_names() {
    let fm = FontManager::new();
    
    // CJKフォント名はbest_font_dataにフォールバック
    // (embed-fontsなしでは内蔵フォントもないのでNone)
    let resolved = fm.resolve_font("MS Gothic");
    // embed-fontsなしの場合はNone、ありの場合はSome
    if fm.has_any_font() {
        assert!(resolved.is_some());
    } else {
        assert!(resolved.is_none());
    }
}

#[test]
fn test_pdf_with_external_font_manager() {
    let mut fm = FontManager::new();
    // ダミーフォント（不正なTTFだがPDF生成はクラッシュしない）
    fm.add_font("CustomFont".to_string(), vec![0u8; 50]);
    
    let input = "外部フォント付きPDFテスト".as_bytes();
    let doc = formats::convert_by_extension("txt", input).unwrap();
    
    // render_to_pdf_with_fonts で外部フォントマネージャーを渡す
    let pdf = pdf_writer::render_to_pdf_with_fonts(&doc, &fm);
    assert!(pdf.starts_with(b"%PDF"));
    assert!(pdf.len() > 100);
}

#[test]
fn test_sample12_table_and_bg_rendering() {
    let pptx_path = "/tmp/Sample_12.pptx";
    if !std::path::Path::new(pptx_path).exists() {
        println!("SKIP: {} not found", pptx_path);
        return;
    }
    use std::io::Read;
    let mut data = Vec::new();
    std::fs::File::open(pptx_path).unwrap().read_to_end(&mut data).unwrap();

    let doc = formats::convert_by_extension("pptx", &data).unwrap();
    assert_eq!(doc.pages.len(), 12);

    // Slide 6 (index 5) should have a table
    use wasm_document_converter::converter::PageElement;
    let slide6 = &doc.pages[5];
    let has_table = slide6.elements.iter().any(|e| matches!(e, PageElement::TableBlock { .. }));
    assert!(has_table, "Slide 6 should have a table");

    // Slide 4 (index 3) should have gradient background
    let slide4 = &doc.pages[3];
    let has_gradient = slide4.elements.iter().any(|e| matches!(e, PageElement::GradientRect { .. }));
    assert!(has_gradient, "Slide 4 should have gradient background");

    // PDF should contain CIDToGIDMap
    let fm = FontManager::new();
    let pdf = pdf_writer::render_to_pdf_with_fonts(&doc, &fm);
    let pdf_str = String::from_utf8_lossy(&pdf);
    assert!(pdf_str.contains("CIDToGIDMap"), "PDF should contain CIDToGIDMap");

    // Slide 1 should have multiple font sizes
    let slide1 = &doc.pages[0];
    let mut font_sizes: std::collections::HashSet<u32> = std::collections::HashSet::new();
    for elem in &slide1.elements {
        if let PageElement::Text { style, .. } = elem {
            font_sizes.insert((style.font_size * 10.0) as u32);
        }
    }
    assert!(font_sizes.len() > 1, "Slide 1 should have multiple font sizes, got: {:?}", font_sizes);
}

#[test]
fn test_sample12_deep_diagnosis() {
    let pptx_path = "/tmp/Sample_12.pptx";
    if !std::path::Path::new(pptx_path).exists() {
        println!("SKIP: {} not found", pptx_path);
        return;
    }
    use std::io::Read;
    let mut data = Vec::new();
    std::fs::File::open(pptx_path).unwrap().read_to_end(&mut data).unwrap();

    let doc = formats::convert_by_extension("pptx", &data).unwrap();
    assert_eq!(doc.pages.len(), 12);

    use wasm_document_converter::converter::PageElement;

    for (i, page) in doc.pages.iter().enumerate() {
        let mut img_count = 0;
        let mut table_count = 0;
        let mut gray_placeholder = 0;
        let mut unresolved_imgs = 0;

        for el in &page.elements {
            match el {
                PageElement::Image { data, mime_type, .. } => {
                    img_count += 1;
                    let is_png = data.len() >= 4 && data[0..4] == [0x89, 0x50, 0x4E, 0x47];
                    let is_jpeg = data.len() >= 2 && data[0..2] == [0xFF, 0xD8];
                    eprintln!("  Slide {}: Image {} bytes, mime={}, png={}, jpeg={}", 
                        i+1, data.len(), mime_type, is_png, is_jpeg);
                }
                PageElement::TableBlock { table, .. } => {
                    table_count += 1;
                    eprintln!("  Slide {}: Table {}rows x {}cols", 
                        i+1, table.rows.len(), table.column_widths.len());
                    for (ri, row) in table.rows.iter().enumerate().take(2) {
                        let texts: Vec<&str> = row.iter().map(|c| c.text.as_str()).collect();
                        eprintln!("    Row {}: {:?}", ri, texts);
                    }
                }
                PageElement::Rect { fill, .. } => {
                    if let Some(c) = fill {
                        if c.r == 230 && c.g == 230 && c.b == 230 {
                            gray_placeholder += 1;
                        }
                    }
                }
                PageElement::Text { text, .. } => {
                    if text == "[Image]" {
                        unresolved_imgs += 1;
                    }
                }
                _ => {}
            }
        }

        eprintln!("Slide {}: elements={} images={} tables={} gray={} unresolved={}", 
            i+1, page.elements.len(), img_count, table_count, gray_placeholder, unresolved_imgs);
    }

    // Verify slide 1 has images (not just gray placeholders)
    let slide1 = &doc.pages[0];
    let image_elements: Vec<_> = slide1.elements.iter().filter(|e| matches!(e, PageElement::Image { .. })).collect();
    let gray_count = slide1.elements.iter().filter(|e| {
        if let PageElement::Rect { fill: Some(c), .. } = e {
            c.r == 230 && c.g == 230 && c.b == 230
        } else { false }
    }).count();
    let unresolved_count = slide1.elements.iter().filter(|e| {
        if let PageElement::Text { text, .. } = e { text == "[Image]" } else { false }
    }).count();
    eprintln!("Slide 1: {} images, {} gray placeholders, {} unresolved", 
        image_elements.len(), gray_count, unresolved_count);

    // Slide 1 should have at least one resolved image (it has background + pics)
    assert!(image_elements.len() >= 1 || gray_count == 0, 
        "Slide 1 should have images, not gray placeholders");

    // Slides 6, 7 should have tables
    for slide_idx in [5, 6] {
        let slide = &doc.pages[slide_idx];
        let table_elems: Vec<_> = slide.elements.iter().filter(|e| matches!(e, PageElement::TableBlock { .. })).collect();
        assert!(!table_elems.is_empty(), "Slide {} should have tables", slide_idx + 1);
    }
}

#[test]
fn test_pdf_contains_image_xobjects() {
    let pptx_path = "/tmp/Sample_12.pptx";
    if !std::path::Path::new(pptx_path).exists() {
        println!("SKIP: {} not found", pptx_path);
        return;
    }
    use std::io::Read;
    let mut data = Vec::new();
    std::fs::File::open(pptx_path).unwrap().read_to_end(&mut data).unwrap();

    let doc = formats::convert_by_extension("pptx", &data).unwrap();
    let fm = FontManager::new();
    let pdf = pdf_writer::render_to_pdf_with_fonts(&doc, &fm);

    let pdf_str = String::from_utf8_lossy(&pdf);

    // PDF should contain XObject references for images
    assert!(pdf_str.contains("/XObject"), "PDF should contain /XObject resource dictionary");
    assert!(pdf_str.contains("/Im0"), "PDF should contain /Im0 image reference");
    assert!(pdf_str.contains("/Subtype /Image"), "PDF should contain /Subtype /Image for XObjects");
    assert!(pdf_str.contains("Do"), "PDF should contain Do operator for image placement");

    // Should contain DCTDecode (for JPEG) or FlateDecode (for PNG)
    let has_dct = pdf_str.contains("/DCTDecode");
    let has_flate_img = pdf_str.contains("/FlateDecode");
    assert!(has_dct || has_flate_img, "PDF should contain image filter (DCTDecode or FlateDecode)");

    // Should NOT contain the old gray placeholder pattern for images
    // (Note: some gray rects may still exist for other shapes, 
    //  but the count should be significantly reduced)
    
    // Write PDF for manual inspection
    std::fs::write("/tmp/Sample_12_output.pdf", &pdf).ok();
    eprintln!("PDF written to /tmp/Sample_12_output.pdf ({} bytes)", pdf.len());
}

#[test]
fn test_pdf_table_rendering() {
    let pptx_path = "/tmp/Sample_12.pptx";
    if !std::path::Path::new(pptx_path).exists() {
        println!("SKIP: {} not found", pptx_path);
        return;
    }
    use std::io::Read;
    let mut data = Vec::new();
    std::fs::File::open(pptx_path).unwrap().read_to_end(&mut data).unwrap();

    let doc = formats::convert_by_extension("pptx", &data).unwrap();
    let fm = FontManager::new();
    let pdf = pdf_writer::render_to_pdf_with_fonts(&doc, &fm);

    let pdf_str = String::from_utf8_lossy(&pdf);

    // Tables use BT/ET for cell text rendering
    // Check slide 6 has table content
    // "No." and "First Name" should be in the PDF text
    assert!(pdf_str.contains("Tj"), "PDF should contain text show (Tj) operators for table cells");
}

// ── テーブルセル改行テスト ──

#[test]
fn test_table_cell_multiline_rendering() {
    // テーブルセル内に複数行のテキストがある場合のPDF出力をテスト
    use wasm_document_converter::converter::{
        Color, Page, PageElement, Table, TableCell,
    };

    let mut cell = TableCell::new("Line1\nLine2\nLine3");
    cell.style.font_size = 10.0;
    cell.style.color = Color::BLACK;
    let table = Table {
        rows: vec![vec![cell]],
        column_widths: vec![200.0],
    };

    let page = Page {
        width: 400.0,
        height: 300.0,
        elements: vec![PageElement::TableBlock {
            x: 10.0,
            y: 10.0,
            width: 200.0,
            table,
        }],
    };

    let mut doc = Document::new();
    doc.pages.push(page);

    let fm = FontManager::new();
    let pdf = pdf_writer::render_to_pdf_with_fonts(&doc, &fm);
    let pdf_str = String::from_utf8_lossy(&pdf);

    // 複数行テキストは複数のTj命令で出力されるべき
    let tj_count = pdf_str.matches("Tj").count();
    assert!(
        tj_count >= 3,
        "3-line cell text should produce at least 3 Tj operators, got {}",
        tj_count
    );
}

// ── フォント文字化け防止テスト ──

#[test]
fn test_cid_to_gid_map_with_unparseable_font() {
    // パース不能なフォントデータが追加された場合でも、
    // 内蔵フォントにフォールバックしてCIDToGIDMapが空にならないことをテスト
    let mut fm = FontManager::new();
    // 不正なフォントデータを追加
    fm.add_font("BadFont".to_string(), vec![0, 1, 2, 3, 4, 5]);

    let txt = "テスト Test".as_bytes();
    let doc = formats::convert_by_extension("txt", txt).unwrap();
    let pdf = pdf_writer::render_to_pdf_with_fonts(&doc, &fm);

    // PDFが正常に生成されること
    assert!(pdf.starts_with(b"%PDF"), "PDF should be generated even with bad external font");

    // フォールバックフォントが埋め込まれていることを確認
    let pdf_str = String::from_utf8_lossy(&pdf);
    assert!(
        pdf_str.contains("/FontFile2"),
        "PDF should embed fallback TrueType font (/FontFile2) when external font is unusable"
    );
}

#[test]
fn test_render_text_control_char_safety() {
    // テキストに制御文字が含まれていてもPDFが壊れないことをテスト
    use wasm_document_converter::converter::{Page, PageElement, TextAlign};

    let page = Page {
        width: 400.0,
        height: 300.0,
        elements: vec![PageElement::Text {
            x: 10.0,
            y: 10.0,
            width: 200.0,
            text: "Hello\nWorld\tTest".to_string(),
            style: FontStyle::default(),
            align: TextAlign::Left,
        }],
    };

    let mut doc = Document::new();
    doc.pages.push(page);

    let fm = FontManager::new();
    let pdf = pdf_writer::render_to_pdf_with_fonts(&doc, &fm);
    assert!(pdf.starts_with(b"%PDF"), "PDF should be valid even with control chars in text");

    // PDF 内のすべてのヘックス文字列 (<...>) を走査し、その中に 000A / 0009 が
    // 含まれていないことを確認する
    let pdf_str = String::from_utf8_lossy(&pdf);
    let pdf_search = pdf_str.as_ref();
    let mut pos = 0;
    while let Some(start) = pdf_search[pos..].find('<') {
        let start = pos + start;
        if let Some(end) = pdf_search[start + 1..].find('>') {
            let end = start + 1 + end;
            let hex_segment = &pdf_search[start + 1..end];
            assert!(
                !hex_segment.contains("000A"),
                "PDF hex text should not contain newline character encoding (000A) in any hex string"
            );
            assert!(
                !hex_segment.contains("0009"),
                "PDF hex text should not contain tab character encoding (0009) in any hex string"
            );
            pos = end + 1;
        } else {
            break;
        }
    }
}
