@echo off
chcp 65001 >nul
echo ========================================
echo ChroLens_Portal 自動打包與發布工具
echo ========================================
echo.

REM 檢查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 錯誤: 找不到 Python
    pause
    exit /b 1
)

REM 檢查並安裝必要套件
echo 檢查必要套件...
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo 正在安裝 PyInstaller...
    pip install pyinstaller
)

python -c "import github" >nul 2>&1
if errorlevel 1 (
    echo 正在安裝 PyGithub...
    pip install PyGithub
)

echo.
echo 開始打包與發布...
echo.

REM 執行打包腳本
python build_and_release.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo 執行失敗
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo 完成！
echo ========================================
pause
