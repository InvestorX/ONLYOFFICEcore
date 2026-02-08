# WASM ドキュメントコンバーター

Rust + WebAssembly で構築されたドキュメント変換ツールです。ブラウザ上で各種ドキュメントフォーマットをPDFや画像（ZIP）に変換できます。

## ⚠️ よくある質問（FAQ）

### Q: `www/pkg/wasm_document_converter.js` が見つかりません

**A:** このファイルは **Rust側でビルド時に自動生成されます**。

以下のコマンドを実行してビルドしてください：

```bash
# Linux / macOS
./build.sh build

# Windows
build.bat build
```

詳細は [ビルド方法](#ビルド方法) セクションをご覧ください。

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
| Microsoft PowerPoint | .pptx | ✅ 完全対応（シェイプ位置・書式・画像・**86種類の全プリセットジオメトリ**・カスタムジオメトリ・グラデーション・シャドウ・3D効果・チャート・SmartArt） |
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

### ⚠️ 重要: ビルド前の注意事項

**`www/pkg/wasm_document_converter.js` ファイルが見つからない場合:**

このファイルは **Rust側でビルド時に自動生成されます**。リポジトリには含まれていません（`.gitignore`で除外）。

以下のビルドコマンドを実行すると、`www/pkg/` ディレクトリに自動的に生成されます：
- `wasm_document_converter.js` — JavaScriptバインディング
- `wasm_document_converter_bg.wasm` — WebAssemblyバイナリ
- `wasm_document_converter.d.ts` — TypeScript型定義
- `package.json` — npm パッケージ情報

### 前提条件

- [Rust](https://rustup.rs/) (1.70以上)
- [wasm-pack](https://rustwasm.github.io/wasm-pack/installer/)

```bash
# wasm-packのインストール
cargo install wasm-pack

# WASMターゲットの追加
rustup target add wasm32-unknown-unknown
```

### ビルド方法（選択してください）

以下のいずれかの方法でビルドできます：
- **方法1: 自動スクリプトを使用**（推奨・簡単）
- **方法2: 手動で1ステップずつ実行**（詳細を理解したい場合）

---

### 方法1: 自動スクリプトを使用（推奨）

#### Linux / macOS

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

#### Windows

```cmd
REM リリースビルド
build.bat build

REM デバッグビルド
build.bat dev

REM テスト実行
build.bat test

REM または直接cargoコマンド:
cargo test --lib
```

---

### 方法2: 手動で1ステップずつ実行

自動スクリプトを使わず、各ステップを個別に実行したい場合は以下の手順に従ってください。

#### ステップ1: 前提条件の確認

まず、必要なツールがインストールされているか確認します。

```bash
# Rustのバージョン確認
rustc --version
# 出力例: rustc 1.70.0 (...)

# Cargoの確認
cargo --version
# 出力例: cargo 1.70.0 (...)

# wasm-packの確認
wasm-pack --version
# 出力例: wasm-pack 0.12.0

# wasm-packがインストールされていない場合
cargo install wasm-pack
```

#### ステップ2: WASMターゲットの追加

Rustコンパイラにwasm32ターゲットを追加します（初回のみ）。

```bash
# wasm32-unknown-unknownターゲットの追加
rustup target add wasm32-unknown-unknown

# インストール済みターゲットの確認
rustup target list --installed
# wasm32-unknown-unknown が含まれていることを確認
```

#### ステップ3: 日本語フォントのダウンロード（オプション）

外部フォントを埋め込む場合、事前にダウンロードします。

**Linux / macOS:**

```bash
# fontsディレクトリを作成
mkdir -p fonts

# Noto Sans JP フォントをダウンロード
curl -L -o fonts/NotoSansJP-Regular.ttf \
  "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf"

# ダウンロード確認
ls -lh fonts/NotoSansJP-Regular.ttf
```

**Windows (PowerShell):**

```powershell
# fontsディレクトリを作成
New-Item -ItemType Directory -Force -Path fonts

# Noto Sans JP フォントをダウンロード
Invoke-WebRequest -Uri "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf" `
  -OutFile "fonts\NotoSansJP-Regular.ttf"

# ダウンロード確認
Get-Item fonts\NotoSansJP-Regular.ttf
```

**Windows (curl):**

```cmd
REM fontsディレクトリを作成
mkdir fonts

REM Noto Sans JP フォントをダウンロード
curl -L -o fonts\NotoSansJP-Regular.ttf ^
  "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf"

REM ダウンロード確認
dir fonts\NotoSansJP-Regular.ttf
```

**注意:** フォントダウンロードは任意です。フォントをダウンロードしない場合でも、デフォルトで内蔵フォント（Noto Sans JP サブセット版）が使用されます。

#### ステップ4: WASMビルドの実行

wasm-packを使ってWebAssemblyバイナリをビルドします。

**リリースビルド（本番用・最適化あり）:**

```bash
# wasm-converterディレクトリで実行
wasm-pack build --target web --release --out-dir www/pkg
```

**デバッグビルド（開発用・最適化なし）:**

```bash
# デバッグ情報付きでビルド
wasm-pack build --target web --dev --out-dir www/pkg
```

**フォント埋め込みビルド（外部フォントを含める場合）:**

```bash
# embed-fonts フィーチャーを有効化してビルド
wasm-pack build --target web --release --out-dir www/pkg -- --features embed-fonts
```

ビルドが成功すると、`www/pkg/` ディレクトリに以下のファイルが生成されます：
- `wasm_document_converter.js` — JavaScriptバインディング
- `wasm_document_converter_bg.wasm` — WebAssemblyバイナリ
- `wasm_document_converter.d.ts` — TypeScript型定義
- `package.json` — npmパッケージ情報

#### ステップ5: ビルド結果の確認

生成されたファイルが正しく出力されているか確認します。

```bash
# 出力ファイルの確認
ls -lh www/pkg/

# 主要ファイルの存在確認
ls www/pkg/wasm_document_converter.js
ls www/pkg/wasm_document_converter_bg.wasm
```

**Windows:**

```cmd
REM 出力ファイルの確認
dir www\pkg\

REM 主要ファイルの存在確認
dir www\pkg\wasm_document_converter.js
dir www\pkg\wasm_document_converter_bg.wasm
```

#### ステップ6: テストの実行（オプション）

Rustのユニットテストと統合テストを実行して、コードが正しく動作するか確認します。

```bash
# wasm-converterディレクトリで実行
cargo test --lib

# すべてのテストを実行（ネイティブ環境で実行）
cargo test

# 詳細な出力を表示
cargo test --lib -- --nocapture
```

テストが成功すると、以下のような出力が表示されます：
```
running 77 tests
...
test result: ok. 77 passed; 0 failed; 0 ignored; 0 measured; 0 filtered out
```

#### ステップ7: ローカルサーバーで動作確認

ビルドしたWASMを実際にブラウザで動かして確認します。

**Linux / macOS:**

```bash
# wwwディレクトリに移動
cd www

# Python 3でHTTPサーバーを起動
python3 -m http.server 8080

# または Python 2の場合
python -m SimpleHTTPServer 8080
```

**Windows:**

```cmd
REM wwwディレクトリに移動
cd www

REM PythonでHTTPサーバーを起動
python -m http.server 8080
```

ブラウザで `http://localhost:8080` にアクセスして、ドキュメント変換機能をテストできます。

#### ステップ8: ビルドのクリーンアップ（必要に応じて）

ビルド成果物を削除して、クリーンな状態に戻すことができます。

```bash
# 生成されたWASMファイルを削除
rm -rf www/pkg/

# Cargoのビルドキャッシュも削除する場合
cargo clean
```

**Windows:**

```cmd
REM 生成されたWASMファイルを削除
rmdir /s /q www\pkg

REM Cargoのビルドキャッシュも削除する場合
cargo clean
```

---

### 補足: ビルドオプション一覧

| コマンド | 説明 |
|:---|:---|
| `wasm-pack build --target web --release` | 本番用リリースビルド（最適化あり） |
| `wasm-pack build --target web --dev` | 開発用デバッグビルド（デバッグ情報あり） |
| `wasm-pack build ... -- --features embed-fonts` | フォント埋め込みビルド |
| `--out-dir www/pkg` | 出力先ディレクトリを指定 |
| `cargo test --lib` | ユニットテストのみ実行 |
| `cargo test` | すべてのテスト実行 |
| `cargo clean` | ビルドキャッシュをクリア |

---

### スクリプトでのローカルサーバー起動

ビルド後、自動スクリプトを使ってローカルサーバーを起動することもできます。

#### Linux / macOS

```bash
./build.sh serve
# http://localhost:8080 にアクセス
```

#### Windows

```cmd
build.bat serve
REM http://localhost:8080 にアクセス
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

### 内蔵フォント

デフォルトで **Noto Sans CJK JP**（サブセット版、約130KB）が内蔵されています。
このフォントは以下の文字セットをカバーします:
- ASCII（U+0020-007E）、Latin-1 Supplement（U+00A0-00FF）
- ひらがな（U+3040-309F）、カタカナ（U+30A0-30FF）
- CJK記号・句読点（U+3000-303F）
- 基本漢字（約500字：常用漢字の主要部分）
- 全角英数字・カタカナ

フォント未埋め込み時でも基本的なラテン文字・日本語テキストが正しく描画されます。

### 外部フォントの追加（実行時読み込み）

より多くの漢字や他のCJK言語の文字を表示するには、実行時に外部フォントを読み込んでください。
コンパイル後のWASMバイナリに対して、`addFont()`で外部フォントを追加できます。

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
| `converter.rs` | コアトレイト・型定義（Document, Page, PageElement, PathCommand, GradientRect, Ellipse等） |
| `pdf_writer.rs` | 軽量PDF生成エンジン（Unicode対応、グラデーション、ベジェ楕円、パス描画、Helveticaフォールバック） |
| `image_renderer.rs` | ページ画像化（ab_glyphフォントラスタライズ、パススキャンライン塗りつぶし、JPEG/PNGデコード、グラデーション・楕円描画） + ZIPバンドル |
| `font_manager.rs` | フォント管理（NotoSansJP内蔵 + 実行時外部フォント読み込み、CJKフォント名自動解決） |
| `formats/pptx_layout.rs` | PPTXコンバーター（シェイプ/塗り/グラデーション/テーマ/グループ/シャドウ/3D/チャート/SmartArt/**86種類のプリセットジオメトリ**/カスタムジオメトリ） |
| `formats/docx_layout.rs` | DOCXコンバーター（段落/ラン書式/テーブル/画像/自動ページ分割） |
| `formats/chart.rs` | チャートレンダリング（棒/円/面/折れ線/散布） |
| `formats/smartart.rs` | SmartArt/ダイアグラムレンダリング（dsp:drawing解析、テキスト抽出、グリッドレイアウト） |
| `formats/odt.rs` | ODTコンバーター（OpenDocument Text テキスト抽出・メタデータ） |
| `formats/epub.rs` | EPUBコンバーター（OPF/spine解析・XHTML テキスト抽出） |
| `formats/odp.rs` | ODPコンバーター（OpenDocument Presentation スライドテキスト抽出） |
| `formats/` | その他のフォーマットコンバーター（txt, csv, rtf, xlsx） |
| `lib.rs` | WASMエントリーポイント（wasm-bindgen API + フォント管理API） |

## プリセットジオメトリ対応

PPTX変換では**86種類すべて**のプリセットジオメトリに完全対応しています。すべての図形が正確にレンダリングされます。

### 基本図形（10種類）
- 三角形（triangle, isosTriangle, rtTriangle）
- 四角形（diamond, parallelogram, trapezoid）
- 多角形（pentagon, hexagon, octagon）
- 角丸矩形（roundRect）、角切り矩形（snip1Rect、snip2SameRect、snip2DiagRect）

### 矢印（15種類）
- arrow, rightArrow, leftArrow, upArrow, downArrow
- leftRightArrow, upDownArrow, notchedRightArrow
- bentArrow, bentUpArrow, curvedRightArrow, curvedLeftArrow
- stripedRightArrow, chevron, homePlate

### フローチャート（17種類完全対応）
- flowChartProcess, flowChartDecision, flowChartTerminator
- flowChartDocument, flowChartPredefinedProcess
- flowChartInputOutput, flowChartPreparation
- flowChartManualInput, flowChartManualOperation
- flowChartConnector, flowChartOffpageConnector
- flowChartSort, flowChartExtract, flowChartMerge
- flowChartDelay, flowChartDisplay
- flowChartMultidocument, flowChartOnlineStorage
- flowChartAlternateProcess, flowChartMagneticDisk

### 星形（10種類）
- star4, star5, star6, star8, star10, star12, star16, star24, star32
- 角丸星形（roundStar）

### 特殊図形（20種類以上）
- moon（三日月）, smileyFace（スマイル）, sun（太陽）
- noSmoking（禁煙マーク）, heart（ハート）
- lightningBolt（稲妻）, cloud（雲）, cloudCallout
- foldedCorner（折り目付き）, frame（額縁）, bevel（面取り）
- gear6（歯車6）, gear9（歯車9）
- irregularSeal1, irregularSeal2, explosion1, explosion2
- plus（プラス）, cross（クロス）

### その他（14種類）
- arc（弧）, pie（扇形）, donut（ドーナツ）
- wave（波）, doubleWave（二重波）
- ribbon（リボン）, ribbon2, ellipseRibbon, ellipseRibbon2
- アクションボタン（actionButtonBlank, actionButtonHome, actionButtonHelp, actionButtonBack, actionButtonForward）

### カスタムジオメトリ（完全対応）
OOXML `<a:custGeom>` 要素による完全なカスタムパス定義に対応：
- `<a:moveTo>`, `<a:lnTo>` — 移動・直線
- `<a:cubicBezTo>` — 3次ベジェ曲線
- `<a:quadBezTo>` — 2次ベジェ曲線
- `<a:close>` — パスのクローズ
- ビューポート座標系の自動スケーリング

すべてのプリセットジオメトリとカスタムジオメトリはPDF出力とPNG画像出力の両方で正確にレンダリングされます。

## ライセンス

GNU AGPL v3.0 — 詳細は [LICENSE.txt](../LICENSE.txt) を参照してください。

### フォントライセンス

- Noto Sans JP / Noto Serif JP: [SIL Open Font License 1.1](https://scripts.sil.org/OFL)
- LINE Seed JP: [SIL Open Font License 1.1](https://scripts.sil.org/OFL) — [公式サイト](https://seed.line.me/) / [GitHub](https://github.com/line/seed)
