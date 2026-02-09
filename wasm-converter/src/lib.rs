// lib.rs - WebAssemblyエントリーポイント
//
// wasm-bindgen を使用してJavaScriptから呼び出し可能なAPIを公開します。
// ドキュメント変換の全フローを統合します。

pub mod converter;
pub mod font_manager;
pub mod formats;
pub mod image_renderer;
pub mod pdf_writer;

use converter::detect_format;
use font_manager::FontManager;
use wasm_bindgen::prelude::*;

/// WASMコンバーターのメインインスタンス
/// JavaScriptからこのオブジェクトを作成して使用します。
#[wasm_bindgen]
pub struct WasmConverter {
    font_manager: FontManager,
}

#[wasm_bindgen]
impl WasmConverter {
    /// 新しいコンバーターインスタンスを作成
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            font_manager: FontManager::new(),
        }
    }

    /// 外部フォントデータを追加（実行時にフォントを読み込み）
    /// コンパイル後でも外部からフォントを追加できます。
    /// TTFまたはOTFフォーマットのバイト列を受け付けます。
    /// @param name フォント名（例: "NotoSansJP-Regular", "Meiryo"）
    /// @param data フォントファイルのバイト列（Uint8Array）
    #[wasm_bindgen(js_name = addFont)]
    pub fn add_font(&mut self, name: String, data: Vec<u8>) {
        self.font_manager.add_font(name, data);
    }

    /// 外部フォントを削除
    /// @param name 削除するフォント名
    #[wasm_bindgen(js_name = removeFont)]
    pub fn remove_font(&mut self, name: &str) {
        self.font_manager.remove_font(name);
    }

    /// 日本語内蔵フォントが利用可能かどうか
    #[wasm_bindgen(js_name = hasJapaneseFont)]
    pub fn has_japanese_font(&self) -> bool {
        self.font_manager.has_builtin_japanese_font()
    }

    /// いずれかのフォントが利用可能かどうか（外部フォント含む）
    #[wasm_bindgen(js_name = hasAnyFont)]
    pub fn has_any_font(&self) -> bool {
        self.font_manager.has_any_font()
    }

    /// 利用可能なフォント名の一覧をJSON配列で取得
    #[wasm_bindgen(js_name = listFonts)]
    pub fn list_fonts(&self) -> String {
        let fonts = self.font_manager.available_fonts();
        serde_json::to_string(&fonts).unwrap_or_else(|_| "[]".to_string())
    }

    /// 読み込まれた外部フォントの数を取得
    #[wasm_bindgen(js_name = externalFontCount)]
    pub fn external_font_count(&self) -> usize {
        self.font_manager.external_font_count()
    }

    /// サポートされているフォーマット一覧をJSON文字列で取得
    #[wasm_bindgen(js_name = supportedFormats)]
    pub fn supported_formats(&self) -> String {
        let formats = formats::supported_formats();
        let list: Vec<serde_json::Value> = formats
            .iter()
            .map(|(name, exts)| {
                serde_json::json!({
                    "name": name,
                    "extensions": exts,
                })
            })
            .collect();
        serde_json::to_string(&list).unwrap_or_else(|_| "[]".to_string())
    }

    /// ファイルをPDFに変換
    /// @param filename ファイル名（拡張子でフォーマットを判定）
    /// @param data ファイルのバイト列
    /// @returns PDFバイト列
    #[wasm_bindgen(js_name = convertToPdf)]
    pub fn convert_to_pdf(&self, filename: &str, data: &[u8]) -> Result<Vec<u8>, JsValue> {
        let ext = detect_format(filename).ok_or_else(|| {
            JsValue::from_str(&format!(
                "サポートされていないファイル形式です: {}",
                filename
            ))
        })?;

        let doc = formats::convert_by_extension(ext, data).map_err(|e| JsValue::from_str(&e.to_string()))?;

        Ok(pdf_writer::render_to_pdf_with_fonts(&doc, &self.font_manager))
    }

    /// ファイルを画像に変換してZIPで返す
    /// @param filename ファイル名
    /// @param data ファイルのバイト列
    /// @param dpi 画像の解像度（デフォルト: 150）
    /// @returns ZIPバイト列（各ページがPNG画像）
    #[wasm_bindgen(js_name = convertToImagesZip)]
    pub fn convert_to_images_zip(
        &self,
        filename: &str,
        data: &[u8],
        dpi: Option<f64>,
    ) -> Result<Vec<u8>, JsValue> {
        let ext = detect_format(filename).ok_or_else(|| {
            JsValue::from_str(&format!(
                "サポートされていないファイル形式です: {}",
                filename
            ))
        })?;

        let doc = formats::convert_by_extension(ext, data).map_err(|e| JsValue::from_str(&e.to_string()))?;

        let mut config = image_renderer::ImageRenderConfig::default();
        if let Some(d) = dpi {
            // DPIを72〜600の範囲にクランプしてOOMを防止
            config.dpi = d.clamp(72.0, 600.0);
        }

        Ok(image_renderer::render_to_images_zip_with_config(
            &doc,
            &self.font_manager,
            &config,
        ))
    }

    /// ファイルをJSON形式のドキュメントモデルに変換（デバッグ用）
    /// @param filename ファイル名
    /// @param data ファイルのバイト列
    /// @returns ドキュメントモデルのJSON文字列
    #[wasm_bindgen(js_name = convertToJson)]
    pub fn convert_to_json(&self, filename: &str, data: &[u8]) -> Result<String, JsValue> {
        let ext = detect_format(filename).ok_or_else(|| {
            JsValue::from_str(&format!(
                "サポートされていないファイル形式です: {}",
                filename
            ))
        })?;

        let doc = formats::convert_by_extension(ext, data).map_err(|e| JsValue::from_str(&e.to_string()))?;

        serde_json::to_string_pretty(&doc)
            .map_err(|e| JsValue::from_str(&format!("JSONシリアライズエラー: {}", e)))
    }
}

/// 簡易変換関数（インスタンスなしで使用可能）
/// @param filename ファイル名
/// @param data ファイルのバイト列
/// @param output_format "pdf" または "images_zip"
/// @returns 変換結果のバイト列
#[wasm_bindgen(js_name = convertDocument)]
pub fn convert_document(
    filename: &str,
    data: &[u8],
    output_format: &str,
) -> Result<Vec<u8>, JsValue> {
    let converter = WasmConverter::new();

    match output_format {
        "pdf" => converter.convert_to_pdf(filename, data),
        "images_zip" | "zip" => converter.convert_to_images_zip(filename, data, None),
        _ => Err(JsValue::from_str(&format!(
            "サポートされていない出力形式です: {} (pdf または images_zip を指定してください)",
            output_format
        ))),
    }
}

/// バージョン情報を取得
#[wasm_bindgen(js_name = getVersion)]
pub fn get_version() -> String {
    format!(
        "WASM Document Converter v{} ({}フォーマット対応)",
        env!("CARGO_PKG_VERSION"),
        formats::supported_formats().len()
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_detect_format() {
        assert_eq!(detect_format("test.txt"), Some("txt"));
        assert_eq!(detect_format("test.docx"), Some("docx"));
        assert_eq!(detect_format("test.xlsx"), Some("xlsx"));
        assert_eq!(detect_format("test.csv"), Some("csv"));
        assert_eq!(detect_format("test.unknown"), None);
    }

    #[test]
    fn test_convert_txt_to_pdf() {
        let data = "Hello, World!\nこんにちは世界！".as_bytes();
        let doc = formats::convert_by_extension("txt", data).unwrap();
        let pdf = pdf_writer::render_to_pdf(&doc);
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_convert_csv_to_pdf() {
        let data = b"Name,Age\nAlice,30\nBob,25";
        let doc = formats::convert_by_extension("csv", data).unwrap();
        let pdf = pdf_writer::render_to_pdf(&doc);
        assert!(pdf.starts_with(b"%PDF"));
    }

    #[test]
    fn test_convert_txt_to_images_zip() {
        let data = "テスト文書".as_bytes();
        let doc = formats::convert_by_extension("txt", data).unwrap();
        let fm = font_manager::FontManager::new();
        let zip_data = image_renderer::render_to_images_zip(&doc, &fm);
        // ZIPファイルの先頭はPKシグネチャ
        assert!(zip_data.starts_with(b"PK"));
    }

    #[test]
    fn test_unsupported_format() {
        let result = formats::convert_by_extension("xyz", b"data");
        assert!(result.is_err());
    }

    #[test]
    fn test_supported_formats_json() {
        let fmts = formats::supported_formats();
        assert!(fmts.len() > 10);
    }
}
