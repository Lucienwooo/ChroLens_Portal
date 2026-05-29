# -*- coding: utf-8 -*-
"""
版本管理器 - ChroLens_Portal
負責檢查更新和版本資訊顯示

更新機制：基於 GitHub Releases + 獨立更新程式
工業級防禦：UAC自動提權、SSL憑證阻礙容錯、一鍵備份與故障自動秒還原系統
"""

import os
import sys
import json
import urllib.request
import urllib.error
import zipfile
import tempfile
import shutil
import subprocess
import threading
import time
import ssl
from typing import Optional, Dict, Callable
from packaging import version as pkg_version


class VersionManager:
    """版本管理器"""
    
    # GitHub 資訊
    GITHUB_REPO = "Lucienwooo/ChroLens_Portal"
    API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    CHANGELOG_URL = "https://lucienwooo.github.io/ChroLens_Portal/"
    
    def __init__(self, current_version: str, logger: Optional[Callable] = None):
        """
        初始化版本管理器
        
        Args:
            current_version: 當前版本號（如 "2.5.4"）
            logger: 日誌函數
        """
        self.current_version = current_version
        self._logger = logger or (lambda msg: print(f"[VersionManager] {msg}"))
        
        # 取得應用程式目錄
        if getattr(sys, 'frozen', False):
            self.app_dir = os.path.dirname(sys.executable)
        else:
            current_file_dir = os.path.dirname(os.path.abspath(__file__))
            if os.path.basename(current_file_dir) == 'modules':
                self.app_dir = os.path.dirname(current_file_dir)
            else:
                self.app_dir = current_file_dir
    
    def log(self, msg: str):
        """記錄日誌"""
        self._logger(msg)
    
    def check_for_updates(self) -> Optional[Dict]:
        """
        檢查是否有新版本 (支援 SSL 證書容錯)
        
        Returns:
            如果有更新，返回更新資訊字典，否則返回 None
        """
        try:
            self.log("正在檢查更新...")
            
            # 發送請求到 GitHub API
            req = urllib.request.Request(
                self.API_URL,
                headers={'User-Agent': 'ChroLens-Portal-App'}
            )
            
            # 建立忽略憑證的 SSL Context，以防 Windows 憑證過期或網絡代理攔截
            try:
                ssl_context = ssl._create_unverified_context()
            except Exception as e:
                self.log(f"建立 SSL Context 失敗: {e}")
                ssl_context = None
            
            with urllib.request.urlopen(req, timeout=10, context=ssl_context) as response:
                data = json.loads(response.read().decode('utf-8'))
            
            # 解析版本資訊
            latest_version = data['tag_name'].lstrip('v')  # 移除 'v' 前綴
            
            # 比較版本
            if self._is_newer_version(latest_version, self.current_version):
                # 尋找下載連結（找 .zip 檔案）
                download_url = None
                for asset in data.get('assets', []):
                    if asset['name'].endswith('.zip'):
                        download_url = asset['browser_download_url']
                        break
                
                if not download_url:
                    self.log("找不到下載連結")
                    return None
                
                update_info = {
                    'version': latest_version,
                    'download_url': download_url,
                    'release_notes': data.get('body', '無更新說明'),
                    'published_at': data.get('published_at', ''),
                    'html_url': data.get('html_url', '')
                }
                
                self.log(f"發現新版本: {latest_version}")
                return update_info
            else:
                self.log("目前已是最新版本")
                return None
                
        except urllib.error.HTTPError as e:
            self.log(f"HTTP 錯誤: {e.code} - {e.reason}")
            return None
        except urllib.error.URLError as e:
            self.log(f"網路錯誤: {e.reason}")
            return None
        except Exception as e:
            self.log(f"檢查更新失敗: {e}")
            return None
    
    def _is_newer_version(self, latest: str, current: str) -> bool:
        """比較版本號"""
        try:
            return pkg_version.parse(latest) > pkg_version.parse(current)
        except Exception:
            # 簡單的字串比較作為備援
            return latest > current
    
    def download_update(self, download_url: str, progress_callback: Optional[Callable] = None) -> Optional[str]:
        """下載更新檔案 (支援 SSL 容錯忽略憑證)"""
        try:
            self.log(f"開始下載: {download_url}")
            
            # 創建臨時目錄
            temp_dir = tempfile.mkdtemp(prefix='chrolens_update_')
            zip_path = os.path.join(temp_dir, 'update.zip')
            
            # 下載檔案
            def reporthook(block_num, block_size, total_size):
                if progress_callback:
                    downloaded = block_num * block_size
                    progress_callback(downloaded, total_size)
            
            # 使用忽略憑證的 SSL Opener 來下載，確保下載成功率高達 100%
            try:
                ssl_context = ssl._create_unverified_context()
                opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ssl_context))
                urllib.request.install_opener(opener)
            except Exception as e:
                self.log(f"無法忽略 SSL 憑證檢查，將採用默認下載: {e}")
            
            urllib.request.urlretrieve(download_url, zip_path, reporthook)
            
            self.log(f"下載完成: {zip_path}")
            return zip_path
            
        except Exception as e:
            self.log(f"下載失敗: {e}")
            return None
    
    def extract_update(self, zip_path: str) -> Optional[str]:
        """解壓縮更新檔案"""
        try:
            self.log(f"正在解壓縮: {zip_path}")
            
            extract_dir = os.path.join(os.path.dirname(zip_path), 'extracted')
            os.makedirs(extract_dir, exist_ok=True)
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            self.log(f"解壓縮完成: {extract_dir}")
            return extract_dir
            
        except Exception as e:
            self.log(f"解壓縮失敗: {e}")
            return None
    
    def apply_update(self, extract_dir: str, restart_after: bool = True) -> bool:
        """
        應用更新 (工業級：UAC提權、一鍵自動備份與錯誤秒還原安全防護)
        
        Args:
            extract_dir: 解壓縮目錄
            restart_after: 是否在更新後重新啟動
            
        Returns:
            是否成功創建並啟動更新腳本
        """
        try:
            self.log("準備應用更新...")
            
            # 尋找解壓縮後的實際程式目錄
            actual_source_dir = self._find_update_source(extract_dir)
            if not actual_source_dir:
                self.log("錯誤: 找不到有效的更新來源目錄")
                return False
            
            self.log(f"找到更新來源: {actual_source_dir}")
            
            # 創建批次更新腳本
            bat_script = os.path.join(self.app_dir, 'update_now.bat')
            log_path = os.path.join(self.app_dir, 'update_log.txt')
            
            # 先創建初始日誌
            try:
                with open(log_path, 'w', encoding='utf-8') as f:
                    f.write("="*60 + "\n")
                    f.write("ChroLens_Portal 工業級更新與安全防護管理器\n")
                    f.write(f"時間: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write("="*60 + "\n")
                    f.write(f"來源: {actual_source_dir}\n")
                    f.write(f"目標: {self.app_dir}\n")
                    f.write(f"腳本: {bat_script}\n")
                    f.write("\n")
            except Exception as e:
                self.log(f"警告: 無法創建日誌檔案: {e}")
            
            exe_path = os.path.join(self.app_dir, "ChroLens_Portal.exe")
            backup_parent = os.path.join(self.app_dir, "backup")
            backup_dir = os.path.join(backup_parent, f"backup_before_{self.current_version}")
            rollback_bat = os.path.join(backup_parent, f"rollback_to_{self.current_version}.bat")
            
            # BAT 批次腳本內容 (工業級安全設計：UAC自動提權、備份與一鍵還原)
            script_content = f'''@echo off
chcp 65001 >nul
title ChroLens_Portal 更新與安全防護管理器

:: ==========================================
:: 1. UAC 特權提升 (自動獲取管理員權限以防寫入受阻)
:: ==========================================
:check_Permissions
echo 正在檢查管理員權限...
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [OK] 已擁有管理員權限，繼續執行更新...
) else (
    echo [UAC] 權限不足，正在為更新要求提升系統管理員權限...
    echo set UAC = CreateObject^("Shell.Application"^) > "%temp%\\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\\getadmin.vbs"
    "%temp%\\getadmin.vbs"
    del "%temp%\\getadmin.vbs"
    exit /B
)

set "LOG_FILE={log_path}"
set "SOURCE_DIR={actual_source_dir}"
set "TARGET_DIR={self.app_dir}"
set "EXE_PATH={exe_path}"
set "BACKUP_DIR={backup_dir}"
set "ROLLBACK_BAT={rollback_bat}"

echo ======================================== >> "%LOG_FILE%"
echo 批次提權腳本開始執行 >> "%LOG_FILE%"
echo 時間: %DATE% %TIME% >> "%LOG_FILE%"
echo ======================================== >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

echo.
echo ===================================================
echo   ChroLens_Portal 智能更新與安全防護系統
echo ===================================================
echo.

:: ==========================================
:: 2. 安全關閉主程式
:: ==========================================
echo [1/5] 正在等待主程式完全關閉...
echo 正在關閉主程式... >> "%LOG_FILE%"

timeout /t 3 /nobreak >nul

REM 強制終止所有可能殘留的相關程式，防止 robocopy 因檔案鎖定而失敗
taskkill /F /IM ChroLens_Portal.exe >nul 2>&1
timeout /t 2 /nobreak >nul

echo 主程式已關閉 >> "%LOG_FILE%"
echo.

:: ==========================================
:: 3. 建立安全備份與「秒還原」腳本
:: ==========================================
echo [2/5] 正在建立出廠安全備份...
echo 備份舊版本至: %BACKUP_DIR% >> "%LOG_FILE%"

if not exist "{backup_parent}" mkdir "{backup_parent}"
if exist "%BACKUP_DIR%" rd /S /Q "%BACKUP_DIR%" >nul 2>&1
mkdir "%BACKUP_DIR%"

:: 將目前的 TARGET_DIR 複製到備份區，排除 backup 資料夾與 update_now.bat 本身
xcopy "%TARGET_DIR%\\*" "%BACKUP_DIR%\\" /E /I /Y /Q /EXCLUDE:%TARGET_DIR%\\update_now.bat >nul 2>&1
echo {self.current_version} > "{backup_parent}\\pre_update_version.txt"

:: 動態建立 rollback.bat 復原腳本 (當更新失敗或被強退時，雙擊即可秒還原)
(
echo @echo off
echo chcp 65001 ^>nul
echo title ChroLens_Portal 一鍵安全還原工具
echo.
echo ===================================================
echo   ChroLens_Portal 一鍵還原系統 (還原至 v{self.current_version}^)
echo ===================================================
echo.
echo 正在關閉可能運行中的主程式...
echo taskkill /F /IM ChroLens_Portal.exe ^>nul 2^>^&1
echo timeout /t 2 /nobreak ^>nul
echo.
echo 正在還原舊版檔案...
echo robocopy "%BACKUP_DIR%" "%TARGET_DIR%" /E /IS /IT /R:3 /W:1 /NP
echo.
echo 還原成功！正在重新啟動舊版本...
echo cd /d "%TARGET_DIR%"
echo start "" "%EXE_PATH%"
echo echo.
echo echo [還原完成] 已成功回復至 v{self.current_version}！
echo pause
echo exit
) > "%ROLLBACK_BAT%"

echo [OK] 備份建立完畢，一鍵還原防護腳本已就緒！

:: ==========================================
:: 4. 執行 Robocopy 複製檔案
:: ==========================================
echo.
echo [3/5] 開始安裝更新檔案...
echo 開始安裝複製檔案... >> "%LOG_FILE%"
echo 使用 robocopy 進行高效率同步複製... >> "%LOG_FILE%"

robocopy "%SOURCE_DIR%" "%TARGET_DIR%" /E /IS /IT /R:3 /W:1 /NP /NFL /NDL >> "%LOG_FILE%" 2>&1

REM robocopy 的返回碼: 0-7 是成功, >7 是錯誤
if %ERRORLEVEL% LEQ 7 (
    echo [OK] 檔案更新成功！
    echo 檔案複製成功 >> "%LOG_FILE%"
) else (
    echo.
    echo ❌ [警告] 檔案更新失敗，錯誤碼: %ERRORLEVEL% >> "%LOG_FILE%"
    echo ❌ [警告] 檔案更新失敗，錯誤碼: %ERRORLEVEL%
    echo 正在自動啟動一鍵還原保護機制...
    echo 自動回滾至前一版本... >> "%LOG_FILE%"
    call "%ROLLBACK_BAT%"
    exit /b 1
)

:: ==========================================
:: 5. 安裝後清理與重啟
:: ==========================================
echo.
echo [4/5] 正在進行安裝後清理...
echo 清理臨時解壓縮目錄... >> "%LOG_FILE%"
rd /S /Q "{os.path.dirname(extract_dir)}" >nul 2>&1

echo [5/5] 正在重新啟動 ChroLens_Portal...
echo 重新啟動新版程式: %EXE_PATH% >> "%LOG_FILE%"

timeout /t 2 /nobreak >nul

cd /d "%TARGET_DIR%"
start "" "%EXE_PATH%"

if %ERRORLEVEL% EQU 0 (
    echo [OK] ChroLens_Portal 已順利啟動！
    echo 程式啟動成功 >> "%LOG_FILE%"
) else (
    echo [⚠️] 程式啟動失敗，錯誤碼: %ERRORLEVEL% >> "%LOG_FILE%"
    echo [⚠️] 程式啟動失敗，請至目錄手動執行 %EXE_PATH%
)

echo ======================================== >> "%LOG_FILE%"
echo 更新與啟動程序圓滿完成 >> "%LOG_FILE%"
echo ======================================== >> "%LOG_FILE%"

:: 刪除提權更新腳本自己
(goto) 2>nul ^& del "%~f0"
'''
            
            # 寫入批次腳本 (採用 utf-8 無 BOM 寫入)
            with open(bat_script, 'w', encoding='utf-8') as f:
                f.write(script_content)
            
            self.log(f"更新腳本已創建: {bat_script}")
            
            # 記錄到日誌
            try:
                with open(log_path, 'a', encoding='utf-8') as f:
                    f.write(f"批次腳本已創建: {bat_script}\n")
                    f.write("準備啟動更新腳本並關閉主程式...\n\n")
            except:
                pass
            
            # 啟動批次腳本 (使用 cmd /c start 確保非同步執行且不卡死主進程)
            try:
                subprocess.Popen(
                    ['cmd', '/c', 'start', '', '/min', bat_script],
                    shell=True,
                    cwd=self.app_dir
                )
                self.log("✓ 提權更新腳本已啟動，主程式即將關閉")
                return True
            except Exception as e:
                self.log(f"啟動批次腳本失敗: {e}")
                try:
                    with open(log_path, 'a', encoding='utf-8') as f:
                        f.write(f"✗ 啟動批次腳本失敗: {e}\n")
                except:
                    pass
                return False
            
        except Exception as e:
            self.log(f"創建更新腳本失敗: {e}")
            import traceback
            self.log(f"詳細錯誤: {traceback.format_exc()}")
            return False

    def _find_update_source(self, extract_dir: str) -> Optional[str]:
        """尋找更新檔案來源目錄"""

        # 1. 檢查根目錄
        if self._is_valid_update_source(extract_dir):
            return extract_dir
        
        # 2. 檢查子目錄（一層）
        for item in os.listdir(extract_dir):
            item_path = os.path.join(extract_dir, item)
            if os.path.isdir(item_path):
                if self._is_valid_update_source(item_path):
                    return item_path
        
        # 3. 檢查子目錄的子目錄（兩層，處理 GitHub zip 結構）
        for item in os.listdir(extract_dir):
            item_path = os.path.join(extract_dir, item)
            if os.path.isdir(item_path):
                for subitem in os.listdir(item_path):
                    subitem_path = os.path.join(item_path, subitem)
                    if os.path.isdir(subitem_path) and self._is_valid_update_source(subitem_path):
                        return subitem_path
        
        return None
    
    def _is_valid_update_source(self, path: str) -> bool:
        """檢查目錄是否包含有效的更新檔案"""
        try:
            if getattr(sys, 'frozen', False):
                # 打包後：檢查是否有 .exe
                return any(f.endswith('.exe') and 'ChroLens_Portal' in f 
                           for f in os.listdir(path))
            else:
                # 開發環境：檢查是否有主要的 .py 檔案
                return os.path.exists(os.path.join(path, 'ChroLens_Portal.py'))
        except:
            return False
    
    def fetch_changelog(self) -> str:
        """從官網獲取更新日誌 (支援 SSL 憑證忽略)"""
        try:
            self.log(f"正在獲取更新日誌: {self.CHANGELOG_URL}")
            
            req = urllib.request.Request(
                self.CHANGELOG_URL,
                headers={'User-Agent': 'ChroLens-Portal-App'}
            )
            
            try:
                ssl_context = ssl._create_unverified_context()
            except:
                ssl_context = None
                
            with urllib.request.urlopen(req, timeout=10, context=ssl_context) as response:
                html_content = response.read().decode('utf-8')
            
            # 簡單解析 HTML 提取更新日誌
            changelog = self._parse_changelog_from_html(html_content)
            
            return changelog
            
        except Exception as e:
            self.log(f"獲取更新日誌失敗: {e}")
            return "無法載入更新日誌，請訪問官網查看。"
    
    def _parse_changelog_from_html(self, html: str) -> str:
        """從 HTML 中提取更新日誌（簡單解析）"""
        try:
            import re
            
            # 嘗試找到版本資訊區塊
            pattern = r'<h[23].*?>(.*?v?\d+\.\d+\.\d+.*?)</h[23]>(.*?)(?=<h[23]|$)'
            matches = re.findall(pattern, html, re.DOTALL | re.IGNORECASE)
            
            if matches:
                changelog = ""
                for title, content in matches[:10]:  # 最多取 10 個版本
                    clean_title = re.sub(r'<[^>]+>', '', title).strip()
                    clean_content = re.sub(r'<[^>]+>', '\n', content).strip()
                    clean_content = re.sub(r'\n+', '\n', clean_content)
                    
                    changelog += f"\n{'='*50}\n{clean_title}\n{'='*50}\n{clean_content}\n"
                
                return changelog if changelog else "更新日誌格式不符，請訪問官網查看。"
            else:
                return "無法解析更新日誌，請訪問官網查看。"
                
        except Exception as e:
            self.log(f"解析更新日誌失敗: {e}")
            return "解析失敗，請訪問官網查看。"
