@echo off
chcp 65001 >nul
echo ========================================
echo ChroLens_Portal 打包工具
echo ========================================
echo.

REM 檢查 Python 是否安裝
python --version >nul 2>&1
if errorlevel 1 (
    echo 錯誤: 找不到 Python
    echo 請先安裝 Python 3.8 或更新版本
    pause
    exit /b 1
)

REM 檢查 PyInstaller 是否安裝
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo 正在安裝 PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo 錯誤: 無法安裝 PyInstaller
        pause
        exit /b 1
    )
)

echo 開始打包...
echo.
python build_simple.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo 打包失敗
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo ========================================
pause
