#!/bin/bash
# download_fonts.sh - 日本語フォントダウンロードスクリプト
#
# オープンソースの日本語フォントをダウンロードします。
# ダウンロードしたフォントは fonts/ ディレクトリに配置されます。
#
# 対応フォント:
#   - Noto Sans JP (Google) - SIL Open Font License
#   - Noto Serif JP (Google) - SIL Open Font License
#   - LINE Seed JP (LY Corporation) - SIL Open Font License

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "日本語フォントをダウンロード中..."

# Noto Sans CJK JP (Regular)
if [ ! -f "NotoSansJP-Regular.ttf" ]; then
    echo "  Noto Sans JP Regular をダウンロード中..."
    curl -L -o "NotoSansJP-Regular.ttf" \
        "https://github.com/google/fonts/raw/main/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf" \
        2>/dev/null || echo "  ダウンロード失敗: NotoSansJP-Regular.ttf"
fi

# Noto Sans CJK JP (Bold)
if [ ! -f "NotoSansJP-Bold.ttf" ]; then
    echo "  Noto Sans JP Bold をダウンロード中..."
    curl -L -o "NotoSansJP-Bold.ttf" \
        "https://github.com/google/fonts/raw/main/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf" \
        2>/dev/null || echo "  ダウンロード失敗: NotoSansJP-Bold.ttf"
fi

# Noto Serif CJK JP (Regular)
if [ ! -f "NotoSerifJP-Regular.ttf" ]; then
    echo "  Noto Serif JP Regular をダウンロード中..."
    curl -L -o "NotoSerifJP-Regular.ttf" \
        "https://github.com/google/fonts/raw/main/ofl/notoserifjp/NotoSerifJP%5Bwght%5D.ttf" \
        2>/dev/null || echo "  ダウンロード失敗: NotoSerifJP-Regular.ttf"
fi

# LINE Seed JP (Regular)
if [ ! -f "LINESeedJP-Regular.ttf" ]; then
    echo "  LINE Seed JP Regular をダウンロード中..."
    # GitHub Releasesからダウンロードし、ZIPから該当ファイルを抽出
    LINESEED_ZIP="$(mktemp)"
    LINESEED_DIR="$(mktemp -d)"
    if curl -L -o "$LINESEED_ZIP" \
        "https://github.com/line/seed/releases/download/v20251119/seed-v20251119.zip" \
        2>/dev/null; then
        if unzip -o "$LINESEED_ZIP" -d "$LINESEED_DIR" 2>/dev/null; then
            # TTFファイルを検索してコピー
            REGULAR_TTF=$(find "$LINESEED_DIR" -name "*Regular*" -name "*.ttf" -path "*LINESeedJP*" | head -1)
            BOLD_TTF=$(find "$LINESEED_DIR" -name "*Bold*" -not -name "*ExtraBold*" -name "*.ttf" -path "*LINESeedJP*" | head -1)
            THIN_TTF=$(find "$LINESEED_DIR" -name "*Thin*" -name "*.ttf" -path "*LINESeedJP*" | head -1)
            if [ -n "$REGULAR_TTF" ]; then
                cp "$REGULAR_TTF" "LINESeedJP-Regular.ttf"
                echo "  ✅ LINE Seed JP Regular をダウンロードしました"
            else
                echo "  ダウンロード失敗: LINESeedJP-Regular.ttf (ZIPにTTFが見つかりません)"
            fi
            if [ -n "$BOLD_TTF" ]; then
                cp "$BOLD_TTF" "LINESeedJP-Bold.ttf"
                echo "  ✅ LINE Seed JP Bold をダウンロードしました"
            fi
            if [ -n "$THIN_TTF" ]; then
                cp "$THIN_TTF" "LINESeedJP-Thin.ttf"
                echo "  ✅ LINE Seed JP Thin をダウンロードしました"
            fi
        else
            echo "  ダウンロード失敗: LINESeedJP-Regular.ttf (ZIP展開エラー)"
        fi
    else
        echo "  ダウンロード失敗: LINESeedJP-Regular.ttf"
    fi
    rm -f "$LINESEED_ZIP"
    rm -rf "$LINESEED_DIR"
fi

echo ""
echo "ダウンロード完了！"
echo ""
echo "フォントファイル一覧:"
ls -lh *.ttf *.otf 2>/dev/null || echo "  (フォントファイルなし)"
echo ""
echo "フォントを埋め込んでビルドするには:"
echo "  cd .. && cargo build --features embed-fonts"
