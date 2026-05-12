@echo off
cd /d "%~dp0"
title WMS

:: .env에서 PORT 읽기 (없으면 기본 3000)
set PORT=3000
for /f "tokens=1,2 delims==" %%A in ('type .env 2^>nul ^| findstr /i "^PORT"') do set PORT=%%B

title WMS - localhost:%PORT%
echo Open in browser: http://localhost:%PORT%
echo Keep this window open. Close it to stop the server.
node server.js
pause
