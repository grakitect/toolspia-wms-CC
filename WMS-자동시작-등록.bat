@echo off
cd /d "%~dp0"

echo ========================================
echo  WMS 자동시작 등록
echo ========================================
echo.

:: 시작 프로그램 폴더 경로
set STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set SHORTCUT=%STARTUP%\TOOLSPIA-WMS.vbs
set BAT_PATH=%~dp0WMS-실행.bat

:: VBScript로 백그라운드 창 없이 시작되는 래퍼 생성
echo Set WshShell = CreateObject("WScript.Shell") > "%SHORTCUT%"
echo WshShell.Run Chr(34) ^& "%BAT_PATH%" ^& Chr(34), 1, False >> "%SHORTCUT%"

if exist "%SHORTCUT%" (
    echo [완료] 자동시작 등록 성공!
    echo 경로: %SHORTCUT%
    echo.
    echo 다음 Windows 로그인부터 WMS가 자동으로 실행됩니다.
    echo 브라우저에서 http://localhost:4000 으로 접속하세요.
) else (
    echo [오류] 등록 실패. 관리자 권한으로 실행해보세요.
)

echo.
pause
