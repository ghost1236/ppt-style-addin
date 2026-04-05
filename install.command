#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo ""
echo "======================================"
echo "  PPT 디자인 도구 Add-in 설치"
echo "======================================"
echo ""

# ── PowerPoint에 Add-in 등록 ──────────────
echo "📋 PowerPoint에 Add-in 등록 중..."
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF_DIR"
cp "$SCRIPT_DIR/manifest.xml" "$WEF_DIR/ppt-style-addin.xml"

if [ $? -ne 0 ]; then
    echo "❌ Add-in 등록 실패 (PowerPoint가 설치되어 있는지 확인해주세요)"
    read -p "Enter를 눌러 창을 닫으세요..."
    exit 1
fi
echo "✅ Add-in 등록 완료"

# ── 완료 ─────────────────────────────────
echo ""
echo "======================================"
echo "  ✅ 설치 완료!"
echo "======================================"
echo ""
echo "PowerPoint를 열면 홈 탭에 '디자인 도구' 그룹이 생깁니다."
echo "이미 열려 있다면 완전히 종료 후 다시 열어주세요."
echo ""

open -a "Microsoft PowerPoint" 2>/dev/null

read -p "Enter를 눌러 창을 닫으세요..."
