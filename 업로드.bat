@echo off
chcp 65001 >nul 2>&1
title NT Dashboard Upload

echo.
echo =========================================================
echo   NT Dashboard - Auto Upload
echo =========================================================
echo.

cd /d "%~dp0"

REM Check .git
if not exist ".git" (
    echo [ERROR] .git folder not found!
    pause
    exit /b 1
)
echo [1/4] Repository OK

REM Check data files
echo.
echo [2/4] Checking data files...
echo.
dir /b "data\현월\*.xlsx" 2>nul && echo.
dir /b "data\전월\*.xlsx" 2>nul && echo.

REM Git sync
echo [3/4] Syncing with GitHub...
git pull origin main --quiet 2>nul
git add data/ -A

git diff --cached --quiet
if errorlevel 1 (
    git commit -m "update %date% %time:~0,5%" --quiet

    echo [4/4] Pushing to GitHub...
    git push origin main --quiet

    if errorlevel 1 (
        echo [ERROR] Push failed! Check internet connection.
        pause
        exit /b 1
    )

    echo.
    echo =========================================================
    echo   Upload complete! Web page will update in 1-2 minutes.
    echo.
    echo   https://core-choi.github.io/NT-LEADTIME/
    echo =========================================================
) else (
    echo.
    echo =========================================================
    echo   No changes detected. Already up to date.
    echo =========================================================
)

echo.
pause
