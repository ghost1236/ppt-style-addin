#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$ROOT_DIR"

echo "Building..."
npm run build

REMOTE_URL=$(git remote get-url origin)
COMMIT_MSG="Deploy $(date '+%Y-%m-%d %H:%M:%S')"

echo "Deploying to gh-pages branch..."
cd dist

# dist 안에 임시 git 저장소 생성 후 force push
git init -b gh-pages 2>/dev/null || { git init && git checkout -b gh-pages 2>/dev/null || git checkout -B gh-pages; }
git add -A
git commit -m "$COMMIT_MSG"
git remote add origin "$REMOTE_URL" 2>/dev/null || git remote set-url origin "$REMOTE_URL"
git push -f origin gh-pages

cd "$ROOT_DIR"

echo ""
echo "Done! GitHub Pages에 배포되었습니다."
echo "URL: $(echo $REMOTE_URL | sed 's/git@github.com:/https:\/\//' | sed 's/\.git$//' | sed 's/github.com\///' | awk -F'/' '{print "https://"$1".github.io/"$2"/taskpane/index.html"}')"
