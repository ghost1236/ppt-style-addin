#!/bin/bash

echo ""
echo "======================================"
echo "  PPT 디자인 도구 Add-in 제거"
echo "======================================"
echo ""

MANIFEST_PATH="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/ppt-style-addin.xml"
if [ -f "$MANIFEST_PATH" ]; then
    rm "$MANIFEST_PATH"
    echo "✅ Add-in 등록 해제됨"
else
    echo "ℹ️  등록된 Add-in 없음"
fi

echo ""
echo "======================================"
echo "  ✅ 제거 완료"
echo "======================================"
echo ""
echo "PowerPoint를 재시작하면 Add-in이 사라집니다."
echo ""

read -p "Enter를 눌러 창을 닫으세요..."
