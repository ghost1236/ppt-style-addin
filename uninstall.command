#!/bin/bash

echo ""
echo "======================================"
echo "  PPT 디자인 도구 Add-in 제거"
echo "======================================"
echo ""

# ── 서버 자동 시작 제거 ───────────────────
PLIST_PATH="$HOME/Library/LaunchAgents/com.pptaddon.server.plist"
if [ -f "$PLIST_PATH" ]; then
    launchctl unload "$PLIST_PATH" 2>/dev/null
    rm "$PLIST_PATH"
    echo "✅ 자동 시작 제거됨"
else
    echo "ℹ️  자동 시작 설정 없음"
fi

# ── PowerPoint Add-in 등록 해제 ──────────
MANIFEST_PATH="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/ppt-style-addin.xml"
if [ -f "$MANIFEST_PATH" ]; then
    rm "$MANIFEST_PATH"
    echo "✅ Add-in 등록 해제됨"
else
    echo "ℹ️  등록된 Add-in 없음"
fi

# ── HTTPS 인증서 제거 ────────────────────
if [ -d "$HOME/.office-addin-dev-certs" ]; then
    echo ""
    echo "🔐 HTTPS 인증서 제거 중..."
    npx office-addin-dev-certs uninstall 2>/dev/null
    echo "✅ 인증서 제거됨"
fi

echo ""
echo "======================================"
echo "  ✅ 제거 완료"
echo "======================================"
echo ""
echo "PowerPoint를 재시작하면 Add-in이 사라집니다."
echo ""

read -p "Enter를 눌러 창을 닫으세요..."
