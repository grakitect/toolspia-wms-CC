@echo off
cd /d "%~dp0"

:: .env에서 PORT 읽기 (없으면 기본 4000)
set PORT=4000
for /f "tokens=1,2 delims==" %%A in ('type .env 2^>nul ^| findstr /i "^PORT"') do set PORT=%%B

title WMS - localhost:%PORT%
echo ========================================
echo  TOOLSPIA WMS (CC)
echo  http://localhost:%PORT%
echo  자동 업데이트: 30초마다 git pull
echo  이 창을 닫으면 서버가 종료됩니다.
echo ========================================
echo.

:: 시작 시 최신 코드 받기
echo [업데이트 확인 중...]
git pull origin claude/redesign-wms-ui-VRvSQ
echo.

:: 패키지 설치 확인
if not exist "node_modules\puppeteer-core" (
  echo [패키지 설치 중... 잠시 기다려주세요]
  npm install
  echo.
)

:: 30초마다 git pull 백그라운드 실행
start /min powershell -WindowStyle Hidden -Command ^
  "while($true) { Start-Sleep 30; Set-Location '%~dp0'; git pull origin claude/redesign-wms-ui-VRvSQ 2>$null }"

node server.js
pause
