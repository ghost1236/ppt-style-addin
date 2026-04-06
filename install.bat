@echo off
chcp 65001 >nul 2>&1
setlocal

set "MANIFEST_URL=https://raw.githubusercontent.com/ghost1236/ppt-style-addin/main/manifest.xml"
set "WEF_DIR=%APPDATA%\Microsoft\PowerPoint\wef"
set "MANIFEST_FILE=%WEF_DIR%\a1b2c3d4-e5f6-7890-abcd-ef1234567890.manifest.xml"

echo.
echo =========================================
echo   PPT 디자인 도구 설치 (Windows)
echo =========================================
echo.

:: 인터넷 연결 확인
echo [1/3] 인터넷 연결 확인 중...
curl -s --head "%MANIFEST_URL%" | findstr "200" >nul 2>&1
if errorlevel 1 (
    :: curl이 없으면 powershell로 시도
    powershell -Command "(Invoke-WebRequest -Uri '%MANIFEST_URL%' -Method Head -UseBasicParsing).StatusCode" >nul 2>&1
    if errorlevel 1 (
        echo [오류] 인터넷에 연결할 수 없습니다.
        echo.
        pause
        exit /b 1
    )
)
echo        연결 확인 완료

:: manifest.xml 다운로드 및 등록
echo [2/3] 애드인 다운로드 및 등록 중...
if not exist "%WEF_DIR%" mkdir "%WEF_DIR%"
del /q "%WEF_DIR%\*.xml" >nul 2>&1

curl -sL "%MANIFEST_URL%" -o "%MANIFEST_FILE%" 2>nul
if not exist "%MANIFEST_FILE%" (
    :: curl 실패 시 powershell로 다운로드
    powershell -Command "Invoke-WebRequest -Uri '%MANIFEST_URL%' -OutFile '%MANIFEST_FILE%' -UseBasicParsing" 2>nul
)

if not exist "%MANIFEST_FILE%" (
    echo [오류] 다운로드에 실패했습니다.
    echo.
    pause
    exit /b 1
)
echo        등록 완료

:: PowerPoint 재실행
echo [3/3] PowerPoint 재시작 중...
taskkill /im POWERPNT.EXE /f >nul 2>&1
timeout /t 2 /nobreak >nul
start "" "POWERPNT.EXE" 2>nul

echo.
echo =========================================
echo   설치 완료!
echo =========================================
echo.
echo   PowerPoint가 열리면:
echo   홈 탭 → '추가 기능' 버튼 클릭
echo   → '디자인 도구' 클릭
echo.
echo   처음 한 번만 클릭하면 이후에는
echo   자동으로 로드됩니다.
echo.
echo   * 인터넷 연결이 필요합니다.
echo.
pause
