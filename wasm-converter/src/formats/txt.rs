// formats/txt.rs - テキストファイル変換モジュール
//
// TXTファイルを読み込み、自動エンコーディング検出を行い、
// ドキュメントモデルに変換します。

use crate::converter::{ConvertError, Document, DocumentConverter, FontStyle};

/// テキストファイルコンバーター
pub struct TxtConverter;

impl TxtConverter {
    pub fn new() -> Self {
        Self
    }
}

impl DocumentConverter for TxtConverter {
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError> {
        let text = decode_text(input)?;
        let lines: Vec<String> = text.lines().map(|l| l.to_string()).collect();
        let style = FontStyle::default();
        Ok(Document::from_text_lines(&lines, &style))
    }

    fn supported_extensions(&self) -> &[&str] {
        &["txt"]
    }

    fn format_name(&self) -> &str {
        "TXT"
    }
}

/// テキストのエンコーディングを自動検出してUTF-8に変換
fn decode_text(input: &[u8]) -> Result<String, ConvertError> {
    // BOMチェック
    if input.starts_with(&[0xEF, 0xBB, 0xBF]) {
        // UTF-8 BOM
        return String::from_utf8(input[3..].to_vec())
            .map_err(|e| ConvertError::new("TXT", &format!("UTF-8デコードエラー: {}", e)));
    }
    if input.starts_with(&[0xFF, 0xFE]) {
        // UTF-16LE BOM
        let (result, _, had_errors) = encoding_rs::UTF_16LE.decode(input);
        if had_errors {
            return Err(ConvertError::new("TXT", "UTF-16LEデコードエラー"));
        }
        return Ok(result.into_owned());
    }
    if input.starts_with(&[0xFE, 0xFF]) {
        // UTF-16BE BOM
        let (result, _, had_errors) = encoding_rs::UTF_16BE.decode(input);
        if had_errors {
            return Err(ConvertError::new("TXT", "UTF-16BEデコードエラー"));
        }
        return Ok(result.into_owned());
    }

    // UTF-8として試行
    if let Ok(text) = String::from_utf8(input.to_vec()) {
        return Ok(text);
    }

    // Shift_JIS（日本語）として試行
    let (result, _, had_errors) = encoding_rs::SHIFT_JIS.decode(input);
    if !had_errors {
        return Ok(result.into_owned());
    }

    // EUC-JP（日本語）として試行
    let (result, _, had_errors) = encoding_rs::EUC_JP.decode(input);
    if !had_errors {
        return Ok(result.into_owned());
    }

    // ISO-2022-JP（日本語）として試行
    let (result, _, _) = encoding_rs::ISO_2022_JP.decode(input);
    Ok(result.into_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_utf8_text() {
        let input = "Hello, World!\nこんにちは世界！".as_bytes();
        let converter = TxtConverter::new();
        let doc = converter.convert(input).unwrap();
        assert_eq!(doc.pages.len(), 1);
        assert!(!doc.pages[0].elements.is_empty());
    }

    #[test]
    fn test_empty_text() {
        let input = b"";
        let converter = TxtConverter::new();
        let doc = converter.convert(input).unwrap();
        assert_eq!(doc.pages.len(), 1);
    }

    #[test]
    fn test_utf8_bom() {
        let mut input = vec![0xEF, 0xBB, 0xBF];
        input.extend_from_slice("テスト".as_bytes());
        let result = decode_text(&input).unwrap();
        assert_eq!(result, "テスト");
    }
}
