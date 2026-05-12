@echo off
chcp 65001 >nul
title TOOLSPIA WMS - Windows 설치

echo ============================================
echo  TOOLSPIA WMS - Windows 설치 프로그램
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
echo [OK] Node.js 버전: %NODE_VER%

:: 대상 폴더 생성
set TARGET=C:\Users\home\Desktop\WMS\Claude_Code
echo.
echo 설치 경로: %TARGET%
echo.

if not exist "%TARGET%" (
    mkdir "%TARGET%"
    echo [OK] 폴더 생성 완료
) else (
    echo [OK] 폴더가 이미 존재합니다
)

:: 파일 복사
echo.
echo 파일 복사 중...

xcopy /E /I /Y /Q "%~dp0public" "%TARGET%\public\" >nul
xcopy /E /I /Y /Q "%~dp0scripts" "%TARGET%\scripts\" >nul
xcopy /E /I /Y /Q "%~dp0tools" "%TARGET%\tools\" >nul
xcopy /Y /Q "%~dp0server.js" "%TARGET%\" >nul
xcopy /Y /Q "%~dp0package.json" "%TARGET%\" >nul
xcopy /Y /Q "%~dp0package-lock.json" "%TARGET%\" >nul
xcopy /Y /Q "%~dp0README.md" "%TARGET%\" >nul
xcopy /Y /Q "%~dp0WMS-실행.bat" "%TARGET%\" >nul
xcopy /Y /Q "%~dp0.gitignore" "%TARGET%\" >nul

:: data 폴더 (db.json이 없을 경우에만 복사)
if not exist "%TARGET%\data\db.json" (
    if not exist "%TARGET%\data" mkdir "%TARGET%\data"
    xcopy /Y /Q "%~dp0data\db.json" "%TARGET%\data\" >nul
    echo [OK] 초기 데이터베이스 생성
) else (
    echo [OK] 기존 데이터베이스 유지
)

echo [OK] 파일 복사 완료

:: .env 파일 생성 (없는 경우)
if not exist "%TARGET%\.env" (
    echo PORT=3000> "%TARGET%\.env"
    echo HOST=localhost>> "%TARGET%\.env"
    echo.>> "%TARGET%\.env"
    echo # 이카운트 연동 설정 (선택)>> "%TARGET%\.env"
    echo # ECOUNT_COM_CODE=회사코드>> "%TARGET%\.env"
    echo # ECOUNT_USER_ID=아이디>> "%TARGET%\.env"
    echo # ECOUNT_API_CERT_KEY=인증키>> "%TARGET%\.env"
    echo [OK] .env 설정파일 생성
)

:: npm 의존성 설치
echo.
echo npm 패키지 설치 중...
cd /d "%TARGET%"
call npm install --silent
if %ERRORLEVEL% neq 0 (
    echo [오류] npm install 실패
    pause
    exit /b 1
)
echo [OK] npm 패키지 설치 완료

:: 바탕화면 바로가기 생성
echo.
echo 바탕화면 바로가기 생성 중...
set SHORTCUT=%USERPROFILE%\Desktop\TOOLSPIA WMS.lnk
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%SHORTCUT%'); $s.TargetPath = '%TARGET%\WMS-실행.bat'; $s.WorkingDirectory = '%TARGET%'; $s.Description = 'TOOLSPIA WMS'; $s.Save()" >nul 2>&1
if exist "%SHORTCUT%" (
    echo [OK] 바탕화면 바로가기 생성 완료
) else (
    echo [경고] 바로가기 생성 실패 - 수동으로 WMS-실행.bat을 실행하세요
)

echo.
echo ============================================
echo  설치 완료!
echo ============================================
echo.
echo  실행 방법:
echo   1. 바탕화면의 "TOOLSPIA WMS" 바로가기 클릭
echo   2. 또는 %TARGET%\WMS-실행.bat 실행
echo   3. 브라우저에서 http://localhost:3000 접속
echo.
echo 지금 바로 실행하시겠습니까? (Y/N)
set /p RUN_NOW=
if /i "%RUN_NOW%"=="Y" (
    start "" "%TARGET%\WMS-실행.bat"
)

pause
