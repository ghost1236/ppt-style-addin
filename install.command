#!/bin/bash

# 이 스크립트가 있는 폴더 기준으로 실행
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo ""
echo "======================================"
echo "  PPT 디자인 도구 Add-in 설치"
echo "======================================"
echo ""

# ── Node.js 확인 ──────────────────────────
# Apple Silicon / Intel / nvm 등 다양한 경로 지원
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$HOME/.nvm/versions/node/$(ls $HOME/.nvm/versions/node 2>/dev/null | tail -1)/bin:$PATH"

if ! command -v node &> /dev/null; then
    echo "❌ Node.js가 설치되어 있지 않습니다."
    echo ""
    echo "아래 주소에서 Node.js LTS 버전을 설치 후 다시 실행해주세요."
    echo "   https://nodejs.org"
    echo ""
    read -p "Enter를 눌러 창을 닫으세요..."
    exit 1
fi

echo "✅ Node.js $(node -v) 확인"

# ── 패키지 설치 ───────────────────────────
echo ""
echo "📦 패키지 설치 중... (잠시 기다려주세요)"
npm install --silent 2>&1

if [ $? -ne 0 ]; then
    echo "❌ 패키지 설치 실패. 인터넷 연결을 확인해주세요."
    read -p "Enter를 눌러 창을 닫으세요..."
    exit 1
fi
echo "✅ 패키지 설치 완료"

# ── HTTPS 인증서 설치 ─────────────────────
echo ""
echo "🔐 HTTPS 인증서 설치 중..."
echo "   (Mac 비밀번호 입력창이 뜰 수 있습니다)"
npx office-addin-dev-certs install 2>&1

if [ $? -ne 0 ]; then
    echo "❌ 인증서 설치 실패"
    read -p "Enter를 눌러 창을 닫으세요..."
    exit 1
fi
echo "✅ 인증서 설치 완료"

# ── PowerPoint에 Add-in 등록 ──────────────
echo ""
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

# ── Mac 시작 시 서버 자동 실행 설정 ─────────
echo ""
echo "⚙️  자동 시작 설정 중..."
LAUNCH_AGENTS_DIR="$HOME/Library/LaunchAgents"
mkdir -p "$LAUNCH_AGENTS_DIR"
PLIST_PATH="$LAUNCH_AGENTS_DIR/com.pptaddon.server.plist"
NPM_PATH="$(which npm)"

cat > "$PLIST_PATH" << PLIST_EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.pptaddon.server</string>
    <key>ProgramArguments</key>
    <array>
        <string>${NPM_PATH}</string>
        <string>run</string>
        <string>dev</string>
    </array>
    <key>WorkingDirectory</key>
    <string>${SCRIPT_DIR}</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>/tmp/pptaddon.log</string>
    <key>StandardErrorPath</key>
    <string>/tmp/pptaddon-error.log</string>
    <key>EnvironmentVariables</key>
    <dict>
        <key>PATH</key>
        <string>/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin</string>
    </dict>
</dict>
</plist>
PLIST_EOF

# 이미 실행 중이면 중지 후 재등록
launchctl unload "$PLIST_PATH" 2>/dev/null
launchctl load "$PLIST_PATH"

if [ $? -ne 0 ]; then
    echo "⚠️  자동 시작 설정 실패. 수동으로 서버를 실행해야 합니다."
else
    echo "✅ 자동 시작 설정 완료 (Mac 로그인 시 서버가 자동 실행됩니다)"
fi

# ── 완료 ─────────────────────────────────
echo ""
echo "======================================"
echo "  ✅ 설치 완료!"
echo "======================================"
echo ""
echo "PowerPoint를 열면 홈 탭에 '디자인 도구' 그룹이 생깁니다."
echo "이미 열려 있다면 완전히 종료 후 다시 열어주세요."
echo ""

# PowerPoint 자동 실행
open -a "Microsoft PowerPoint" 2>/dev/null

read -p "Enter를 눌러 창을 닫으세요..."
