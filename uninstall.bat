@echo off
chcp 65001 >nul 2>&1
setlocal

set "WEF_DIR=%APPDATA%\Microsoft\PowerPoint\wef"

echo.
echo =========================================
echo   PPT 디자인 도구 제거 (Windows)
echo =========================================
echo.

echo [1/2] PowerPoint 종료 중...
taskkill /im POWERPNT.EXE /f >nul 2>&1
timeout /t 2 /nobreak >nul

echo [2/2] 애드인 제거 중...
del /q "%WEF_DIR%\a1b2c3d4-e5f6-7890-abcd-ef1234567890.manifest.xml" >nul 2>&1
del /q "%WEF_DIR%\manifest.xml" >nul 2>&1
echo        제거 완료

echo.
echo =========================================
echo   제거 완료!
echo =========================================
echo.
echo   PowerPoint를 다시 열면
echo   디자인 도구가 사라집니다.
echo.
pause
