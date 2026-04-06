#!/bin/bash
# PPT 디자인 도구 설치 (macOS)

MANIFEST_URL="https://raw.githubusercontent.com/ghost1236/ppt-style-addin/main/manifest.xml"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
MANIFEST_FILE="$WEF_DIR/a1b2c3d4-e5f6-7890-abcd-ef1234567890.manifest.xml"

echo ""
echo "========================================="
echo "  PPT 디자인 도구 설치 (macOS)"
echo "========================================="
echo ""

# 인터넷 연결 확인
echo "[1/3] 인터넷 연결 확인 중..."
if ! curl -s --head "$MANIFEST_URL" | head -1 | grep -q "200"; then
    echo "[오류] 인터넷에 연결할 수 없거나"
    echo "       다운로드 주소에 접근할 수 없습니다."
    echo ""
    read -n 1 -s -r -p "아무 키나 누르면 종료합니다..."
    exit 1
fi
echo "       연결 확인 완료"

# manifest.xml 다운로드 및 등록
echo "[2/3] 애드인 다운로드 및 등록 중..."
mkdir -p "$WEF_DIR"
rm -f "$WEF_DIR"/*.xml "$WEF_DIR"/*.manifest.xml
curl -sL "$MANIFEST_URL" -o "$MANIFEST_FILE"
if [ $? -ne 0 ] || [ ! -s "$MANIFEST_FILE" ]; then
    echo "[오류] 다운로드에 실패했습니다."
    echo ""
    read -n 1 -s -r -p "아무 키나 누르면 종료합니다..."
    exit 1
fi
echo "       등록 완료"

# PowerPoint 재시작
echo "[3/3] PowerPoint 재시작 중..."
osascript -e 'tell application "Microsoft PowerPoint" to quit' 2>/dev/null
sleep 2
open -a "Microsoft PowerPoint"

echo ""
echo "========================================="
echo "  설치 완료!"
echo "========================================="
echo ""
echo "  PowerPoint가 열리면:"
echo "  홈 탭 → '추가 기능' 버튼 클릭"
echo "  → '디자인 도구' 클릭"
echo ""
echo "  처음 한 번만 클릭하면 이후에는"
echo "  자동으로 로드됩니다."
echo ""
echo "  * 인터넷 연결이 필요합니다."
echo ""
read -n 1 -s -r -p "아무 키나 누르면 종료합니다..."
