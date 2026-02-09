#!/bin/bash
# build.sh - WASMビルドスクリプト
#
# 使用方法:
#   ./build.sh          # リリースビルド
#   ./build.sh dev      # デバッグビルド
#   ./build.sh test     # テスト実行
#   ./build.sh serve    # ローカルサーバー起動

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# 色付きメッセージ
info() { echo -e "\033[1;34m[INFO]\033[0m $1"; }
success() { echo -e "\033[1;32m[OK]\033[0m $1"; }
error() { echo -e "\033[1;31m[ERROR]\033[0m $1"; }

# wasm-packの確認とインストール
check_wasm_pack() {
    if ! command -v wasm-pack &> /dev/null; then
        info "wasm-pack をインストール中..."
        cargo install wasm-pack
    fi
}

# フォントダウンロード
download_fonts() {
    if [ ! -f "fonts/NotoSansJP-Regular.ttf" ]; then
        info "日本語フォント (Noto Sans JP) をダウンロード中..."
        mkdir -p fonts
        FONT_URL="https://github.com/google/fonts/raw/main/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf"
        if command -v curl &> /dev/null; then
            curl -L -o "fonts/NotoSansJP-Regular.ttf" "$FONT_URL" 2>/dev/null || \
                info "フォントのダウンロードに失敗しました。embed-fontsフィーチャーなしでビルドします。"
        elif command -v wget &> /dev/null; then
            wget -O "fonts/NotoSansJP-Regular.ttf" "$FONT_URL" 2>/dev/null || \
                info "フォントのダウンロードに失敗しました。embed-fontsフィーチャーなしでビルドします。"
        else
            info "curl/wgetが見つかりません。フォントを手動で fonts/ ディレクトリに配置してください。"
        fi
    fi
}

# ネイティブテスト実行
run_tests() {
    info "テストを実行中..."
    cargo test --lib -- --nocapture
    success "全テスト合格！"
}

# WASMビルド
build_wasm() {
    local mode="${1:-release}"

    check_wasm_pack

    if [ "$mode" = "dev" ]; then
        info "デバッグモードでWASMをビルド中..."
        wasm-pack build --target web --dev --out-dir www/pkg
    else
        info "リリースモードでWASMをビルド中..."
        wasm-pack build --target web --release --out-dir www/pkg
    fi

    success "WASMビルド完了！ www/pkg/ に出力されました。"
}

# ローカルサーバー
serve() {
    info "ローカルサーバーを起動中 (http://localhost:8080)..."
    cd www
    if command -v python3 &> /dev/null; then
        python3 -m http.server 8080
    elif command -v python &> /dev/null; then
        python -m SimpleHTTPServer 8080
    else
        error "Python が見つかりません。別の HTTP サーバーを使用してください。"
        exit 1
    fi
}

# メイン
case "${1:-build}" in
    build|release)
        download_fonts
        build_wasm release
        ;;
    dev)
        download_fonts
        build_wasm dev
        ;;
    test)
        run_tests
        ;;
    serve)
        serve
        ;;
    fonts)
        download_fonts
        ;;
    *)
        echo "使用方法: $0 {build|dev|test|serve|fonts}"
        echo ""
        echo "コマンド:"
        echo "  build   - リリースモードでWASMをビルド（デフォルト）"
        echo "  dev     - デバッグモードでWASMをビルド"
        echo "  test    - ネイティブテストを実行"
        echo "  serve   - ローカルHTTPサーバーを起動"
        echo "  fonts   - 日本語フォントをダウンロード"
        exit 1
        ;;
esac
