#!/bin/bash
# html_to_pdf.sh — 用 Chrome headless 把 preview.html 转成 preview.pdf
# 用法: ./html_to_pdf.sh <input.html> [output.pdf]
# 默认输出到桌面"预览稿"目录

set -e

if [ -z "$1" ]; then
    echo "用法: $0 <input.html> [output.pdf]"
    echo "示例: $0 preview.html ~/Desktop/预览稿/项目-v1.pdf"
    exit 1
fi

INPUT="$1"
OUTPUT="${2:-}"

# Chrome 路径（macOS）
CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

if [ ! -x "$CHROME" ]; then
    echo "❌ 未找到 Chrome: $CHROME"
    echo "   请安装 Google Chrome 或修改脚本里的路径"
    exit 1
fi

# 默认输出路径：桌面/预览稿/{输入文件名去后缀}.pdf
if [ -z "$OUTPUT" ]; then
    BASENAME=$(basename "$INPUT" .html)
    OUTPUT="$HOME/Desktop/预览稿/${BASENAME}.pdf"
fi

mkdir -p "$(dirname "$OUTPUT")"

# 转 PDF
"$CHROME" --headless --disable-gpu --no-sandbox \
    --no-pdf-header-footer \
    --print-to-pdf="$OUTPUT" \
    "file://$(realpath "$INPUT")" \
    2>&1 | grep -v "CVDisplayLink\|task_policy_set" | tail -3

if [ -f "$OUTPUT" ]; then
    SIZE=$(du -h "$OUTPUT" | cut -f1)
    echo "✅ PDF 已生成: $OUTPUT ($SIZE)"
else
    echo "❌ PDF 生成失败"
    exit 1
fi
