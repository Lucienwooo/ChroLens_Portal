@echo off
chcp 65001 >nul
title ChroLens_Portal 2.3 啟動器

echo ╔════════════════════════════════════════╗
echo ║   ChroLens_Portal 2.3 模組化版         ║
echo ║   正在啟動...                           ║
echo ╚════════════════════════════════════════╝
echo.

REM 檢查 Python 是否安裝
python --version >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先安裝 Python 3.7 或更新版本
    echo.
    pause
    exit /b 1
)

echo [提示] 程式需要管理員權限才能正常運作
echo [提示] 請在 UAC 提示時點擊「是」
echo.

REM 啟動主程式
python main.py

if errorlevel 1 (
    echo.
    echo [錯誤] 程式執行失敗
    echo.
    pause
)
