@echo off
setlocal enabledelayedexpansion

chcp 65001 >nul

set "PS_SCRIPT=tools\Build-VbaProject.ps1"
set "CONFIG_PATH=config\build-config.json"

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set YMD=%%I
set "LOG_DIR=logs"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
set "LOG_FILE=%LOG_DIR%\build-%YMD%.log"

echo ========================================
echo   VBA 自動ビルドツール v1.1
echo   %date% %time%
echo ========================================

echo [INFO] ビルド処理を開始します...

powershell -NoProfile -ExecutionPolicy Bypass -Command "& { & '%PS_SCRIPT%' -ConfigPath '%CONFIG_PATH%' 2>&1 | Tee-Object -FilePath '%LOG_FILE%' -Append; exit $LASTEXITCODE }"
set ERR=%ERRORLEVEL%
if %ERR%==0 (
  echo [SUCCESS] ビルドが正常に完了しました。出力先: build\
) else (
  echo [ERROR] ビルドに失敗しました。詳細はログをご確認ください: %LOG_FILE%
)

echo.
pause

endlocal
