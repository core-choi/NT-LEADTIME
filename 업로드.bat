@echo off
cd /d "%~dp0"

echo.
echo =========================================================
echo   NT Dashboard - Auto Upload
echo =========================================================
echo.

if not exist ".git" (
    echo [ERROR] .git folder not found!
    pause
    exit /b 1
)

echo [1/4] OK - Repository found
echo.
echo [2/4] Checking files...
dir /b "data\*" 2>nul
echo.

echo [3/4] Adding changes...
git add -A

git diff --cached --quiet
if %errorlevel%==0 (
    echo.
    echo =========================================================
    echo   No changes. Already up to date.
    echo =========================================================
    echo.
    pause
    exit /b 0
)

echo [4/4] Uploading to GitHub...
git commit -m "update %date% %time:~0,5%" --quiet
git push origin main

if %errorlevel%==0 (
    echo.
    echo =========================================================
    echo   Done! Page will update in 1-2 minutes.
    echo   https://core-choi.github.io/NT-LEADTIME/
    echo =========================================================
) else (
    echo.
    echo [ERROR] Upload failed. Check internet.
)

echo.
pause
