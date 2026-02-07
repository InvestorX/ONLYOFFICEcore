// font_manager.rs - フォント管理モジュール
//
// 日本語フォント（Noto Sans JP）を含むフォントの読み込みと管理を行います。
// WebAssemblyバイナリにフォントデータを内包します。

use ab_glyph::{Font, FontRef, PxScale, ScaleFont};

/// 内蔵フォント：Noto Sans JP Regular
/// ビルド時にfonts/ディレクトリからフォントファイルを読み込みます。
/// フォントが存在しない場合はフォールバック用の空バイト列を使用します。
#[cfg(feature = "embed-fonts")]
const NOTO_SANS_JP_REGULAR: &[u8] = include_bytes!("../fonts/NotoSansJP-Regular.ttf");

/// フォントが内蔵されていない場合のフォールバック
#[cfg(not(feature = "embed-fonts"))]
const NOTO_SANS_JP_REGULAR: &[u8] = &[];

/// フォントマネージャー
/// 利用可能なフォントの管理とフォントデータへのアクセスを提供します。
pub struct FontManager {
    /// 外部から読み込まれたフォントデータ
    external_fonts: Vec<(String, Vec<u8>)>,
}

impl FontManager {
    pub fn new() -> Self {
        Self {
            external_fonts: Vec::new(),
        }
    }

    /// 外部フォントデータを追加
    pub fn add_font(&mut self, name: String, data: Vec<u8>) {
        self.external_fonts.push((name, data));
    }

    /// 内蔵日本語フォントが利用可能かどうか
    pub fn has_builtin_japanese_font(&self) -> bool {
        !NOTO_SANS_JP_REGULAR.is_empty()
    }

    /// 内蔵日本語フォントデータを取得
    pub fn builtin_japanese_font(&self) -> Option<&'static [u8]> {
        if NOTO_SANS_JP_REGULAR.is_empty() {
            None
        } else {
            Some(NOTO_SANS_JP_REGULAR)
        }
    }

    /// 名前でフォントデータを取得
    pub fn get_font_data(&self, name: &str) -> Option<&[u8]> {
        // まず外部フォントを検索
        for (font_name, data) in &self.external_fonts {
            if font_name == name {
                return Some(data.as_slice());
            }
        }
        // 内蔵フォントをチェック
        if name.contains("NotoSansJP") || name.contains("Noto") || name.contains("Japanese") {
            return self.builtin_japanese_font();
        }
        None
    }

    /// フォント名のリストを取得
    pub fn available_fonts(&self) -> Vec<String> {
        let mut fonts: Vec<String> = self
            .external_fonts
            .iter()
            .map(|(name, _)| name.clone())
            .collect();
        if self.has_builtin_japanese_font() {
            fonts.push("NotoSansJP-Regular".to_string());
        }
        fonts
    }
}

/// テキストの幅を概算するヘルパー関数
/// フォントが利用可能な場合はab_glyphで正確に計測、
/// そうでない場合はフォントサイズベースで概算します。
pub fn estimate_text_width(text: &str, font_size: f64, font_data: Option<&[u8]>) -> f64 {
    if let Some(data) = font_data {
        if let Ok(font) = FontRef::try_from_slice(data) {
            let scale = PxScale::from(font_size as f32);
            let scaled = font.as_scaled(scale);
            let mut width = 0.0f32;
            for ch in text.chars() {
                let glyph_id = font.glyph_id(ch);
                let advance = scaled.h_advance(glyph_id);
                width += advance;
            }
            return width as f64;
        }
    }
    // フォールバック: 文字種別に基づく概算
    let mut width = 0.0;
    for ch in text.chars() {
        if ch.is_ascii() {
            width += font_size * 0.6; // 半角文字
        } else {
            width += font_size; // 全角文字（日本語等）
        }
    }
    width
}


