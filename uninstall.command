#!/bin/bash
# PPT 디자인 도구 제거 (macOS)

WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

echo ""
echo "========================================="
echo "  PPT 디자인 도구 제거 (macOS)"
echo "========================================="
echo ""

echo "[1/2] PowerPoint 종료 중..."
osascript -e 'tell application "Microsoft PowerPoint" to quit' 2>/dev/null
sleep 2

echo "[2/2] 애드인 제거 중..."
rm -f "$WEF_DIR"/a1b2c3d4-e5f6-7890-abcd-ef1234567890.manifest.xml
rm -f "$WEF_DIR"/manifest.xml
echo "       제거 완료"

echo ""
echo "========================================="
echo "  제거 완료!"
echo "========================================="
echo ""
echo "  PowerPoint를 다시 열면"
echo "  디자인 도구가 사라집니다."
echo ""
read -n 1 -s -r -p "아무 키나 누르면 종료합니다..."
