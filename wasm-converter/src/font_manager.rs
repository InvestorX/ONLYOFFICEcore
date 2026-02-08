// font_manager.rs - フォント管理モジュール
//
// 日本語フォント（Noto Sans JP, LINE Seed JP）を含むフォントの読み込みと管理を行います。
// コンパイル時に内蔵するフォント（embed-fontsフィーチャー）と、
// 実行時に外部から読み込むフォントの両方をサポートします。

use ab_glyph::{Font, FontRef, PxScale, ScaleFont};

/// 内蔵フォント：Noto Sans JP Regular（サブセット版）
/// ASCII + Latin-1 + ひらがな + カタカナ + 基本漢字（約500字）を含む
/// ビルド時にfonts/ディレクトリからフォントファイルを読み込みます。
const NOTO_SANS_JP_REGULAR: &[u8] = include_bytes!("../fonts/NotoSansJP-Regular.ttf");

/// 内蔵フォント：LINE Seed JP Regular（オプション）
#[cfg(feature = "embed-fonts")]
const LINE_SEED_JP_REGULAR: &[u8] = include_bytes!("../fonts/LINESeedJP-Regular.ttf");

/// LINE Seed JP が内蔵されていない場合のフォールバック
#[cfg(not(feature = "embed-fonts"))]
const LINE_SEED_JP_REGULAR: &[u8] = &[];

/// よく使われるOfficeフォント名から日本語フォントへのマッピング
/// これらのフォントがドキュメント内で参照されている場合、
/// 利用可能な日本語フォントにフォールバックします。
const CJK_FONT_NAMES: &[&str] = &[
    "MS Gothic", "MS Mincho", "MS PGothic", "MS PMincho",
    "Yu Gothic", "Yu Mincho", "Meiryo", "HGGothic",
    "HGMincho", "IPAGothic", "IPAMincho", "Hiragino",
    "ヒラギノ", "游ゴシック", "游明朝", "メイリオ",
    "ＭＳ ゴシック", "ＭＳ 明朝", "ＭＳ Ｐゴシック", "ＭＳ Ｐ明朝",
    "SimSun", "SimHei", "MingLiU", "PMingLiU",
    "Malgun Gothic", "Batang", "Gulim",
];

/// フォントマネージャー
/// 利用可能なフォントの管理とフォントデータへのアクセスを提供します。
/// コンパイル時内蔵フォントと実行時外部フォントの両方を管理します。
pub struct FontManager {
    /// 外部から読み込まれたフォントデータ（実行時に追加可能）
    external_fonts: Vec<(String, Vec<u8>)>,
}

impl FontManager {
    pub fn new() -> Self {
        Self {
            external_fonts: Vec::new(),
        }
    }

    /// 外部フォントデータを追加（実行時にフォントを読み込み）
    /// WASMコンパイル後でもこのメソッドでフォントを追加できます。
    /// TTFまたはOTFフォーマットのバイト列を受け付けます。
    pub fn add_font(&mut self, name: String, data: Vec<u8>) {
        // 同名のフォントが既にある場合は置き換える
        self.external_fonts.retain(|(n, _)| n != &name);
        self.external_fonts.push((name, data));
    }

    /// 外部フォントを削除
    pub fn remove_font(&mut self, name: &str) {
        self.external_fonts.retain(|(n, _)| n != name);
    }

    /// 外部フォントの数を取得
    pub fn external_font_count(&self) -> usize {
        self.external_fonts.len()
    }

    /// 内蔵日本語フォントが利用可能かどうか
    pub fn has_builtin_japanese_font(&self) -> bool {
        !NOTO_SANS_JP_REGULAR.is_empty() || !LINE_SEED_JP_REGULAR.is_empty()
    }

    /// いずれかのフォントが利用可能かどうか（外部フォント含む）
    pub fn has_any_font(&self) -> bool {
        !self.external_fonts.is_empty() || self.has_builtin_japanese_font()
    }

    /// 内蔵日本語フォントデータを取得（Noto Sans JP優先）
    pub fn builtin_japanese_font(&self) -> Option<&'static [u8]> {
        if !NOTO_SANS_JP_REGULAR.is_empty() {
            Some(NOTO_SANS_JP_REGULAR)
        } else if !LINE_SEED_JP_REGULAR.is_empty() {
            Some(LINE_SEED_JP_REGULAR)
        } else {
            None
        }
    }

    /// 内蔵LINE Seed JPフォントデータを取得
    pub fn builtin_line_seed_jp(&self) -> Option<&'static [u8]> {
        if LINE_SEED_JP_REGULAR.is_empty() {
            None
        } else {
            Some(LINE_SEED_JP_REGULAR)
        }
    }

    /// 最適なフォントデータを取得（外部フォント優先）
    /// 外部フォントが読み込まれていればそれを使用し、
    /// なければ内蔵フォントにフォールバックします。
    pub fn best_font_data(&self) -> Option<&[u8]> {
        // 外部フォントがあればそれを優先
        if let Some((_, data)) = self.external_fonts.first() {
            if !data.is_empty() {
                return Some(data.as_slice());
            }
        }
        // 内蔵フォントにフォールバック
        self.builtin_japanese_font()
    }

    /// 外部フォントのイテレータを返す
    /// PDF生成時にab_glyphでパース可能なフォントを検索するために使用
    pub fn external_fonts_iter(&self) -> impl Iterator<Item = (&str, &[u8])> {
        self.external_fonts.iter().map(|(name, data)| (name.as_str(), data.as_slice()))
    }

    /// 名前でフォントデータを取得
    pub fn get_font_data(&self, name: &str) -> Option<&[u8]> {
        // まず外部フォントを検索（完全一致）
        for (font_name, data) in &self.external_fonts {
            if font_name == name {
                return Some(data.as_slice());
            }
        }
        // 外部フォントを部分一致で検索
        let name_lower = name.to_lowercase();
        for (font_name, data) in &self.external_fonts {
            if font_name.to_lowercase().contains(&name_lower)
                || name_lower.contains(&font_name.to_lowercase())
            {
                return Some(data.as_slice());
            }
        }
        // 内蔵フォントをチェック
        if name.contains("LINESeed") || name.contains("LINE Seed") || name.contains("LineSeed") {
            return self.builtin_line_seed_jp();
        }
        if name.contains("NotoSansJP") || name.contains("Noto") || name.contains("Japanese") {
            return self.builtin_japanese_font();
        }
        None
    }

    /// ドキュメント内のフォント名を利用可能なフォントに解決
    /// フォント名が直接見つからない場合、利用可能な最適フォントにフォールバックします。
    pub fn resolve_font(&self, name: &str) -> Option<&[u8]> {
        // まず名前でそのまま検索
        if let Some(data) = self.get_font_data(name) {
            return Some(data);
        }
        // CJKフォント名の場合は即フォールバック
        if is_cjk_font_name(name) {
            return self.best_font_data();
        }
        // 西洋フォント（Calibri, Arial等）でもフォントデータがあればフォールバック
        // （文字が全く表示されないよりも代替フォントで表示する方が望ましい）
        self.best_font_data()
    }

    /// フォント名のリストを取得
    pub fn available_fonts(&self) -> Vec<String> {
        let mut fonts: Vec<String> = self
            .external_fonts
            .iter()
            .map(|(name, _)| name.clone())
            .collect();
        if !NOTO_SANS_JP_REGULAR.is_empty() {
            fonts.push("NotoSansJP-Regular".to_string());
        }
        if !LINE_SEED_JP_REGULAR.is_empty() {
            fonts.push("LINESeedJP-Regular".to_string());
        }
        fonts
    }
}

/// フォント名がCJK（日中韓）フォントかどうかを判定
fn is_cjk_font_name(name: &str) -> bool {
    let lower = name.to_lowercase();
    for cjk_name in CJK_FONT_NAMES {
        if lower.contains(&cjk_name.to_lowercase()) {
            return true;
        }
    }
    // 日本語文字が含まれているかチェック
    // U+3000-9FFF: CJK統合漢字、ひらがな、カタカナ、句読点
    // U+FF00-FFEF: 半角・全角形
    name.chars().any(|c| {
        ('\u{3000}'..='\u{9FFF}').contains(&c) || ('\u{FF00}'..='\u{FFEF}').contains(&c)
    })
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


