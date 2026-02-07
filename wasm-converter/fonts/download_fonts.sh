#!/bin/bash
# download_fonts.sh - 日本語フォントダウンロードスクリプト
#
# オープンソースの日本語フォントをダウンロードします。
# ダウンロードしたフォントは fonts/ ディレクトリに配置されます。
#
# 対応フォント:
#   - Noto Sans JP (Google) - SIL Open Font License
#   - Noto Serif JP (Google) - SIL Open Font License

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "日本語フォントをダウンロード中..."

# Noto Sans CJK JP (Regular)
if [ ! -f "NotoSansJP-Regular.ttf" ]; then
    echo "  Noto Sans JP Regular をダウンロード中..."
    curl -L -o "NotoSansJP-Regular.ttf" \
        "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf" \
        2>/dev/null || echo "  ダウンロード失敗: NotoSansJP-Regular.ttf"
fi

# Noto Sans CJK JP (Bold)
if [ ! -f "NotoSansJP-Bold.ttf" ]; then
    echo "  Noto Sans JP Bold をダウンロード中..."
    curl -L -o "NotoSansJP-Bold.ttf" \
        "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Japanese/NotoSansCJKjp-Bold.otf" \
        2>/dev/null || echo "  ダウンロード失敗: NotoSansJP-Bold.ttf"
fi

# Noto Serif CJK JP (Regular)
if [ ! -f "NotoSerifJP-Regular.ttf" ]; then
    echo "  Noto Serif JP Regular をダウンロード中..."
    curl -L -o "NotoSerifJP-Regular.ttf" \
        "https://github.com/googlefonts/noto-cjk/raw/main/Serif/OTF/Japanese/NotoSerifCJKjp-Regular.otf" \
        2>/dev/null || echo "  ダウンロード失敗: NotoSerifJP-Regular.ttf"
fi

echo ""
echo "ダウンロード完了！"
echo ""
echo "フォントファイル一覧:"
ls -lh *.ttf *.otf 2>/dev/null || echo "  (フォントファイルなし)"
echo ""
echo "フォントを埋め込んでビルドするには:"
echo "  cd .. && cargo build --features embed-fonts"
