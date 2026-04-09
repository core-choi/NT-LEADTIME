@echo off
chcp 65001 >nul 2>&1
title NT Dashboard - Daily Schedule

echo.
echo =========================================================
echo   Register daily auto-upload (9:00 AM)
echo =========================================================
echo.

set "REPO_DIR=%~dp0"
set "AUTO_BAT=%~dp0auto_upload.bat"

(
echo @echo off
echo cd /d "%REPO_DIR%"
echo if not exist ".git" exit /b 0
echo git pull origin main --quiet 2^>nul
echo git add data/ -A
echo git diff --cached --quiet ^|^| ^( git commit -m "auto update %%date%% %%time:~0,5%%" --quiet ^& git push origin main --quiet ^)
) > "%AUTO_BAT%"

schtasks /create /tn "NT_Dashboard_Upload" /tr "\"%AUTO_BAT%\"" /sc daily /st 09:00 /f

if errorlevel 1 (
    echo [ERROR] Failed! Right-click and Run as Administrator.
) else (
    echo [OK] Daily 9:00 AM auto-upload registered!
    echo.
    echo   To remove: schtasks /delete /tn "NT_Dashboard_Upload" /f
)

echo.
pause
