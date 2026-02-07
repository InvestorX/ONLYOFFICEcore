# WASM ドキュメントコンバーター

Rust + WebAssembly で構築されたドキュメント変換ツールです。ブラウザ上で各種ドキュメントフォーマットをPDFや画像（ZIP）に変換できます。

## 対応フォーマット

| フォーマット | 拡張子 | 状態 |
|:---|:---|:---|
| テキスト | .txt | ✅ 完全対応 |
| CSV | .csv | ✅ 完全対応 |
| RTF | .rtf | ✅ テキスト抽出 |
| Microsoft Word | .docx | ✅ テキスト抽出 |
| Microsoft Word (旧) | .doc | 🔧 開発中 |
| OpenDocument Text | .odt | 🔧 開発中 |
| EPUB | .epub | 🔧 開発中 |
| XPS | .xps | 🔧 開発中 |
| DjVu | .djvu | 🔧 開発中 |
| Microsoft Excel | .xlsx, .xls, .ods | ✅ テーブル表示 |
| Microsoft PowerPoint | .pptx | 🔧 開発中 |
| Microsoft PowerPoint (旧) | .ppt | 🔧 開発中 |
| OpenDocument Presentation | .odp | 🔧 開発中 |

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

### 外部フォントの追加

```javascript
const converter = new WasmConverter();

// フォントファイルを読み込んで追加
const fontResponse = await fetch('MyFont.ttf');
const fontData = new Uint8Array(await fontResponse.arrayBuffer());
converter.addFont('MyFont', fontData);
```

## アーキテクチャ

```
入力ファイル → [フォーマットコンバーター] → Document モデル → [レンダラー] → 出力
                                                              ├── PDF Writer → PDF
                                                              └── Image Renderer → PNG → ZIP
```

### 主要コンポーネント

| モジュール | 説明 |
|:---|:---|
| `converter.rs` | コアトレイト・型定義（Document, Page, PageElement等） |
| `pdf_writer.rs` | 軽量PDF生成エンジン（外部依存なし、Unicode対応） |
| `image_renderer.rs` | ページ画像化 + ZIPバンドル |
| `font_manager.rs` | フォント管理（日本語フォント内蔵対応） |
| `formats/` | 各フォーマットのコンバーター実装 |
| `lib.rs` | WASMエントリーポイント（wasm-bindgen API） |

## ライセンス

GNU AGPL v3.0 — 詳細は [LICENSE.txt](../LICENSE.txt) を参照してください。

### フォントライセンス

- Noto Sans JP / Noto Serif JP: [SIL Open Font License 1.1](https://scripts.sil.org/OFL)
- LINE Seed JP: [SIL Open Font License 1.1](https://scripts.sil.org/OFL) — [公式サイト](https://seed.line.me/) / [GitHub](https://github.com/line/seed)
