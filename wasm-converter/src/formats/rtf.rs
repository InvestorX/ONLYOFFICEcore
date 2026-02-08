// formats/rtf.rs - RTFファイル変換モジュール
//
// RTF (Rich Text Format) ファイルを解析し、テキストを抽出して
// ドキュメントモデルに変換します。

use crate::converter::{ConvertError, Document, DocumentConverter, FontStyle};

/// RTFコンバーター
pub struct RtfConverter;

impl RtfConverter {
    pub fn new() -> Self {
        Self
    }
}

impl DocumentConverter for RtfConverter {
    fn convert(&self, input: &[u8]) -> Result<Document, ConvertError> {
        let text = String::from_utf8_lossy(input);

        // RTFヘッダーの確認
        if !text.starts_with("{\\rtf") {
            return Err(ConvertError::new("RTF", "有効なRTFファイルではありません"));
        }

        // テキストを抽出
        let extracted = extract_rtf_text(&text);
        let lines: Vec<String> = extracted.lines().map(|l| l.to_string()).collect();
        let style = FontStyle::default();
        Ok(Document::from_text_lines(&lines, &style))
    }

    fn supported_extensions(&self) -> &[&str] {
        &["rtf"]
    }

    fn format_name(&self) -> &str {
        "RTF"
    }
}

/// RTFテキストからプレーンテキストを抽出
/// 制御コードを除去し、テキスト内容のみを返します。
fn extract_rtf_text(rtf: &str) -> String {
    let mut result = String::new();
    let mut chars = rtf.chars().peekable();
    let mut depth = 0;
    let mut skip_group = false;
    let mut skip_groups: Vec<bool> = Vec::new();

    while let Some(ch) = chars.next() {
        match ch {
            '{' => {
                skip_groups.push(skip_group);
                depth += 1;
            }
            '}' => {
                if depth > 0 {
                    depth -= 1;
                }
                skip_group = skip_groups.pop().unwrap_or(false);
            }
            '\\' => {
                if skip_group {
                    // グループをスキップ中
                    continue;
                }
                // 制御語を読み取る
                let mut control_word = String::new();
                while let Some(&next_ch) = chars.peek() {
                    if next_ch.is_ascii_alphabetic() {
                        control_word.push(next_ch);
                        chars.next();
                    } else {
                        break;
                    }
                }

                // 数値パラメータをスキップ
                while let Some(&next_ch) = chars.peek() {
                    if next_ch.is_ascii_digit() || next_ch == '-' {
                        chars.next();
                    } else {
                        break;
                    }
                }

                // スペース区切りをスキップ
                if let Some(&' ') = chars.peek() {
                    chars.next();
                }

                match control_word.as_str() {
                    "par" | "line" => result.push('\n'),
                    "tab" => result.push('\t'),
                    "pict" | "fonttbl" | "colortbl" | "stylesheet" | "info" | "header"
                    | "footer" | "headerf" | "footerf" => {
                        skip_group = true;
                    }
                    _ => {}
                }

                // エスケープ文字
                if control_word.is_empty() {
                    if let Some(&next_ch) = chars.peek() {
                        match next_ch {
                            '\\' | '{' | '}' => {
                                result.push(next_ch);
                                chars.next();
                            }
                            '\'' => {
                                // ヘックスエスケープ
                                chars.next();
                                let mut hex = String::new();
                                for _ in 0..2 {
                                    if let Some(&h) = chars.peek() {
                                        if h.is_ascii_hexdigit() {
                                            hex.push(h);
                                            chars.next();
                                        }
                                    }
                                }
                                if let Ok(byte) = u8::from_str_radix(&hex, 16) {
                                    // Windows-1252としてデコード
                                    let buf = [byte];
                                    let (decoded, _, _) =
                                        encoding_rs::WINDOWS_1252.decode(&buf);
                                    result.push_str(&decoded);
                                }
                            }
                            _ => {}
                        }
                    }
                }
            }
            _ => {
                if !skip_group && depth > 0 {
                    result.push(ch);
                }
            }
        }
    }

    // 連続する空行を整理
    let mut cleaned = String::new();
    let mut prev_empty = false;
    for line in result.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            if !prev_empty {
                cleaned.push('\n');
                prev_empty = true;
            }
        } else {
            cleaned.push_str(trimmed);
            cleaned.push('\n');
            prev_empty = false;
        }
    }

    cleaned.trim().to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_rtf() {
        let input = br#"{\rtf1\ansi Hello World}"#;
        let converter = RtfConverter::new();
        let doc = converter.convert(input).unwrap();
        assert!(!doc.pages.is_empty());
    }

    #[test]
    fn test_invalid_rtf() {
        let input = b"Not an RTF file";
        let converter = RtfConverter::new();
        assert!(converter.convert(input).is_err());
    }

    #[test]
    fn test_rtf_with_paragraphs() {
        let input = br#"{\rtf1\ansi First paragraph\par Second paragraph}"#;
        let converter = RtfConverter::new();
        let doc = converter.convert(input).unwrap();
        assert!(!doc.pages.is_empty());
    }
}
