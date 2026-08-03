@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 서버를 시작합니다...
start "AI 안전일지 서버" powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"
timeout /t 3 /nobreak >nul
start "" "http://localhost:8787"
echo.
echo 브라우저가 자동으로 열립니다. 안 열리면 http://localhost:8787 로 직접 접속하세요.
echo 서버를 끄려면 새로 뜬 "AI 안전일지 서버" 창을 닫으세요.
pause >nul
