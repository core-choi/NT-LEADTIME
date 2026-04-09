@echo off
cd /d "%~dp0"

echo.
echo =========================================================
echo   Register daily auto-upload at 9:00 AM
echo =========================================================
echo.

set "AUTOBAT=%~dp0auto_run.bat"

echo @echo off > "%AUTOBAT%"
echo cd /d "%~dp0" >> "%AUTOBAT%"
echo if not exist ".git" exit /b 0 >> "%AUTOBAT%"
echo git add -A >> "%AUTOBAT%"
echo git diff --cached --quiet ^|^| ( git commit -m "auto %%date%% %%time:~0,5%%" --quiet ^& git push origin main --quiet ) >> "%AUTOBAT%"

schtasks /create /tn "NT_Dashboard" /tr "\"%AUTOBAT%\"" /sc daily /st 09:00 /f

if %errorlevel%==0 (
    echo [OK] Daily 9:00 AM upload registered!
    echo.
    echo   To remove: schtasks /delete /tn "NT_Dashboard" /f
) else (
    echo [ERROR] Failed. Right-click, Run as Administrator.
)

echo.
pause
