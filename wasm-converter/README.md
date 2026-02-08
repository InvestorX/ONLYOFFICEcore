# WASM ドキュメントコンバーター

Rust + WebAssembly で構築されたドキュメント変換ツールです。ブラウザ上で各種ドキュメントフォーマットをPDFや画像（ZIP）に変換できます。

## 対応フォーマット

| フォーマット | 拡張子 | 状態 |
|:---|:---|:---|
| テキスト | .txt | ✅ 完全対応 |
| CSV | .csv | ✅ 完全対応 |
| RTF | .rtf | ✅ テキスト抽出 |
| Microsoft Word | .docx | ✅ レイアウト保持（書式・テーブル・画像対応） |
| Microsoft Word (旧) | .doc | 🔧 開発中 |
| OpenDocument Text | .odt | ✅ テキスト抽出 |
| EPUB | .epub | ✅ テキスト抽出 |
| XPS | .xps | 🔧 開発中 |
| DjVu | .djvu | 🔧 開発中 |
| Microsoft Excel | .xlsx, .xls, .ods | ✅ テーブル表示 |
| Microsoft PowerPoint | .pptx | ✅ レイアウト保持（シェイプ位置・書式・画像対応） |
| Microsoft PowerPoint (旧) | .ppt | 🔧 開発中 |
| OpenDocument Presentation | .odp | ✅ テキスト抽出 |

## 出力形式

- **PDF** — 単一のPDFファイルとして出力
- **画像（ZIP）** — 各ページをPNG画像にレンダリングし、ZIPアーカイブで出力

## 日本語フォント

日本語テキストの表示に対応するため、以下のフォントを内蔵（または外部ロード）できます：

- **Noto Sans JP** (Google Noto Fonts) — SIL Open Font License
- **Noto Serif JP** (Google Noto Fonts) — SIL Open Font License
- **LINE Seed JP** (LY Corporation) — SIL Open Font License

### フォントのダウンロード

```bash
cd fonts
bash download_fonts.sh
```

## ビルド方法

### 前提条件

- [Rust](https://rustup.rs/) (1.70以上)
- [wasm-pack](https://rustwasm.github.io/wasm-pack/installer/)

```bash
# wasm-packのインストール
cargo install wasm-pack

# WASMターゲットの追加
rustup target add wasm32-unknown-unknown
```

### ビルド

```bash
# リリースビルド
./build.sh build

# デバッグビルド
./build.sh dev

# テスト実行
./build.sh test

# または直接cargoコマンド:
cargo test --lib
```

### フォント埋め込みビルド

日本語フォントをWASMバイナリに内蔵する場合：

```bash
# フォントをダウンロード
./build.sh fonts

# フォント埋め込みビルド
wasm-pack build --target web --release --out-dir www/pkg -- --features embed-fonts
```

### ローカルサーバーで動作確認

```bash
./build.sh serve
# http://localhost:8080 にアクセス
```

## 使い方

### JavaScript API

```javascript
import init, { WasmConverter, convertDocument, getVersion } from './pkg/wasm_document_converter.js';

// 初期化
await init();

// バージョン確認
console.log(getVersion());

// コンバーターインスタンスを作成
const converter = new WasmConverter();

// ファイルをPDFに変換
const fileData = new Uint8Array(arrayBuffer);
const pdfBytes = converter.convertToPdf('document.docx', fileData);

// ファイルを画像ZIPに変換（150 DPI）
const zipBytes = converter.convertToImagesZip('document.xlsx', fileData, 150);

// 簡易関数
const result = convertDocument('report.txt', textData, 'pdf');
```

### 外部フォントの追加（実行時読み込み）

コンパイル後のWASMバイナリに対して、実行時に外部フォントを読み込むことができます。
これにより、ドキュメント内で参照されているフォントに近い描画が可能になります。

**フォントの読み込みは、PDFとPNG画像の両方で正確な文字描画を実現するために重要です。**
フォントが読み込まれていない場合、PDF出力ではラテン文字のみ（Helveticaフォールバック）が表示され、
PNG画像出力では簡易矩形描画にフォールバックします。日本語・中国語・韓国語（CJK）テキストの
正確な表示には、対応フォント（例: Noto Sans JP）の読み込みが必要です。

```javascript
const converter = new WasmConverter();

// ファイルから読み込み
const fontResponse = await fetch('NotoSansJP-Regular.ttf');
const fontData = new Uint8Array(await fontResponse.arrayBuffer());
converter.addFont('NotoSansJP', fontData);

// Google Fontsから読み込み
const gfResp = await fetch('https://fonts.gstatic.com/s/notosansjp/v53/...otf');
converter.addFont('NotoSansJP', new Uint8Array(await gfResp.arrayBuffer()));

// フォント一覧の確認
console.log(JSON.parse(converter.listFonts())); // ["NotoSansJP", ...]
console.log(converter.hasAnyFont()); // true
console.log(converter.externalFontCount()); // 1

// フォント削除
converter.removeFont('NotoSansJP');
```

**注意:** フォントはPDFとPNG画像の両方の出力で使用されます。
外部フォントが読み込まれている場合、内蔵フォントより優先されます。
ドキュメント内で参照されるMS明朝、游ゴシック等のCJKフォント名は、
利用可能な最適なフォントに自動的にフォールバックされます。

## アーキテクチャ

```
入力ファイル → [フォーマットコンバーター] → Document モデル → [レンダラー] → 出力
                                                              ├── PDF Writer → PDF
                                                              └── Image Renderer → PNG → ZIP
```

### 主要コンポーネント

| モジュール | 説明 |
|:---|:---|
| `converter.rs` | コアトレイト・型定義（Document, Page, PageElement, GradientRect, Ellipse等） |
| `pdf_writer.rs` | 軽量PDF生成エンジン（外部依存なし、Unicode対応、グラデーション、ベジェ楕円、Helveticaフォールバック） |
| `image_renderer.rs` | ページ画像化（ab_glyphフォントラスタライズ、JPEG/PNGデコード、グラデーション描画、楕円描画） + ZIPバンドル |
| `font_manager.rs` | フォント管理（内蔵フォント + 実行時外部フォント読み込み、CJKフォント名自動解決） |
| `formats/pptx_layout.rs` | PPTXコンバーター（シェイプ/塗り/グラデーション/テーマ/グループ/シャドウ/3D/チャート/SmartArt） |
| `formats/docx_layout.rs` | DOCXコンバーター（段落/ラン書式/テーブル/画像/自動ページ分割） |
| `formats/chart.rs` | チャートレンダリング（棒/円/面/折れ線/散布） |
| `formats/smartart.rs` | SmartArt/ダイアグラムレンダリング |
| `formats/odt.rs` | ODTコンバーター（OpenDocument Text テキスト抽出・メタデータ） |
| `formats/epub.rs` | EPUBコンバーター（OPF/spine解析・XHTML テキスト抽出） |
| `formats/odp.rs` | ODPコンバーター（OpenDocument Presentation スライドテキスト抽出） |
| `formats/` | その他のフォーマットコンバーター（txt, csv, rtf, xlsx） |
| `lib.rs` | WASMエントリーポイント（wasm-bindgen API + フォント管理API） |

## ライセンス

GNU AGPL v3.0 — 詳細は [LICENSE.txt](../LICENSE.txt) を参照してください。

### フォントライセンス

- Noto Sans JP / Noto Serif JP: [SIL Open Font License 1.1](https://scripts.sil.org/OFL)
- LINE Seed JP: [SIL Open Font License 1.1](https://scripts.sil.org/OFL) — [公式サイト](https://seed.line.me/) / [GitHub](https://github.com/line/seed)
