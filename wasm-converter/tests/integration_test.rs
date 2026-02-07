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

/// Sample_12.pptxでのレイアウト保持変換テスト
/// (ファイルが存在しない場合はスキップ)
#[test]
fn test_real_sample12_pptx_layout() {
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
        eprintln!(
            "  スライド {}: {}x{} ({} 要素)",
            i + 1,
            page.width as i32,
            page.height as i32,
            page.elements.len()
        );
    }

    assert!(doc.pages.len() >= 10, "12スライドあるはず（実際: {}）", doc.pages.len());

    // PDF出力
    let pdf = pdf_writer::render_to_pdf(&doc);
    assert!(pdf.starts_with(b"%PDF-1.4"));
    let pdf_path = "/tmp/wasm_converter_sample12.pdf";
    let mut f = std::fs::File::create(pdf_path).unwrap();
    f.write_all(&pdf).unwrap();
    eprintln!("✅ PDF出力: {} ({} bytes)", pdf_path, pdf.len());

    // 画像ZIP出力
    let fm = FontManager::new();
    let zip_data = image_renderer::render_to_images_zip(&doc, &fm);
    assert!(zip_data.starts_with(b"PK"));
    let zip_path = "/tmp/wasm_converter_sample12_pages.zip";
    let mut f = std::fs::File::create(zip_path).unwrap();
    f.write_all(&zip_data).unwrap();
    eprintln!("✅ 画像ZIP出力: {} ({} bytes)", zip_path, zip_data.len());
}
