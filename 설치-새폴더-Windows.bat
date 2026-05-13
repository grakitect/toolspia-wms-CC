@echo off
chcp 65001 >nul
title TOOLSPIA WMS (CC) - 새 폴더 설치

echo ============================================
echo  TOOLSPIA WMS (CC) - 새 폴더 설치
echo  설치 경로: Desktop\WMS\Claude_Code\wms-cc
echo  포트: 4000
echo ============================================
echo.

:: Node.js 확인
where node >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [오류] Node.js가 설치되어 있지 않습니다.
    echo Node.js 18 이상을 설치해주세요: https://nodejs.org
    pause
    exit /b 1
)
for /f "tokens=1" %%v in ('node --version') do set NODE_VER=%%v
echo [OK] Node.js %NODE_VER%

:: Git 확인
where git >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [오류] Git이 설치되어 있지 않습니다.
    echo https://git-scm.com 에서 설치해주세요.
    pause
    exit /b 1
)
echo [OK] Git 확인 완료

:: 설치 경로 설정
set TARGET=C:\Users\home\Desktop\WMS\Claude_Code\wms-cc

echo.
echo 설치 경로: %TARGET%
echo.

:: 기존 폴더 있으면 건너뜀
if exist "%TARGET%\.git" (
    echo [OK] 이미 설치되어 있습니다. 최신 코드로 업데이트합니다...
    cd /d "%TARGET%"
    git pull origin claude/redesign-wms-ui-VRvSQ
    goto :install_deps
)

:: 새 폴더에 클론
echo [1/4] 코드 다운로드 중...
git clone -b claude/redesign-wms-ui-VRvSQ https://github.com/grakitect/toolspia-wms-CC.git "%TARGET%"
if %ERRORLEVEL% neq 0 (
    echo [오류] 다운로드 실패. 인터넷 연결을 확인해주세요.
    pause
    exit /b 1
)
echo [OK] 다운로드 완료

:install_deps
cd /d "%TARGET%"

:: .env 생성 (포트 4000)
echo.
echo [2/4] 설정 파일 생성 중...
if not exist "%TARGET%\.env" (
    echo PORT=4000> "%TARGET%\.env"
    echo HOST=localhost>> "%TARGET%\.env"
    echo.>> "%TARGET%\.env"
    echo # 이카운트 연동 설정 (선택)>> "%TARGET%\.env"
    echo # ECOUNT_COM_CODE=회사코드>> "%TARGET%\.env"
    echo # ECOUNT_USER_ID=아이디>> "%TARGET%\.env"
    echo # ECOUNT_API_CERT_KEY=인증키>> "%TARGET%\.env"
    echo [OK] .env 생성 완료 (포트: 4000)
) else (
    echo [OK] 기존 .env 유지
)

:: npm 설치
echo.
echo [3/4] 패키지 설치 중...
call npm install --silent
if %ERRORLEVEL% neq 0 (
    echo [오류] npm install 실패
    pause
    exit /b 1
)
echo [OK] 패키지 설치 완료

:: 바탕화면 바로가기 생성
echo.
echo [4/4] 바탕화면 바로가기 생성 중...
set SHORTCUT=%USERPROFILE%\Desktop\TOOLSPIA WMS (CC).lnk
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%SHORTCUT%'); $s.TargetPath = '%TARGET%\WMS-실행.bat'; $s.WorkingDirectory = '%TARGET%'; $s.Description = 'TOOLSPIA WMS CC - port 4000'; $s.Save()" >nul 2>&1
if exist "%SHORTCUT%" (
    echo [OK] 바탕화면에 "TOOLSPIA WMS (CC)" 바로가기 생성 완료
) else (
    echo [경고] 바로가기 생성 실패 - %TARGET%\WMS-실행.bat 을 직접 실행하세요
)

echo.
echo ============================================
echo  설치 완료!
echo ============================================
echo.
echo  실행: 바탕화면 "TOOLSPIA WMS (CC)" 클릭
echo  접속: http://localhost:4000
echo  (기존 WMS는 포트 3001로 그대로 유지됩니다)
echo.
echo 지금 바로 실행하시겠습니까? (Y/N)
set /p RUN_NOW=
if /i "%RUN_NOW%"=="Y" (
    start "" "%TARGET%\WMS-실행.bat"
)

pause
