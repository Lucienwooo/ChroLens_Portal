"""
ChroLens_Portal 簡化打包工具
打包主程式為單一執行檔
"""

import os
import sys
import json
import shutil
import subprocess
import zipfile
from pathlib import Path
from datetime import datetime

class PortalBuilder:
    """Portal 打包工具"""
    
    def __init__(self):
        # 專案目錄
        self.project_dir = Path(__file__).parent
        self.main_file = self.project_dir / "ChroLens_Portal.py"
        self.icon_file = self.project_dir / "冥想貓貓.ico"
        
        # 輸出目錄
        self.build_dir = self.project_dir / "build"
        self.dist_dir = self.project_dir / "dist"
        self.output_dir = self.dist_dir / "ChroLens_Portal"
        
        # 讀取版本號
        self.version = self._read_version()
        
        print(f"\n{'='*50}")
        print(f"ChroLens_Portal 簡化打包工具")
        print(f"版本: {self.version}")
        print(f"{'='*50}\n")
    
    def _read_version(self) -> str:
        """從主程式讀取版本號"""
        try:
            with open(self.main_file, 'r', encoding='utf-8') as f:
                for line in f:
                    if line.strip().startswith('CURRENT_VERSION ='):
                        version = line.split('=')[1].strip().strip('"\'')
                        return version
        except Exception as e:
            print(f"警告: 無法讀取版本號: {e}")
            return "2.4"
    
    def clean(self):
        """清理舊檔案"""
        print("\n[1/5] 清理舊檔案...")
        
        for dir_path in [self.build_dir, self.dist_dir]:
            if dir_path.exists():
                print(f"  - 嘗試刪除 {dir_path.name}/")
                try:
                    shutil.rmtree(dir_path, ignore_errors=False)
                    print(f"  ✓ 已刪除 {dir_path.name}/")
                except Exception as e:
                    print(f"  ⚠ 無法完全刪除 {dir_path.name}/ - 將使用現有目錄")
                    print(f"    原因: 某些檔案可能正在使用中")
        
        print("  清理完成\n")
    
    def build_main(self):
        """打包主程式"""
        print("\n[2/5] 打包主程式...")
        
        # 檢查並終止可能佔用檔案的程序
        try:
            import psutil
            for proc in psutil.process_iter(['name']):
                if 'ChroLens_Portal.exe' in proc.info['name']:
                    print(f"  - 發現運行中的程序，正在關閉...")
                    proc.kill()
                    import time
                    time.sleep(2)
        except ImportError:
            print("  - 提示：安裝 psutil 可以自動關閉佔用程序")
            print("    pip install psutil")
        except Exception as e:
            print(f"  - 無法自動關閉程序: {e}")
        
        # PyInstaller 命令
        cmd = [
            'pyinstaller',
            '--clean',
            '--noconfirm',
            '--onedir',
            '--windowed',
            '--name=ChroLens_Portal',
        ]
        
        # 添加圖示
        if self.icon_file.exists():
            cmd.append(f'--icon={self.icon_file}')
        
        # 添加數據文件
        data_files = [
            ('update_manager.py', '.'),
            ('update_dialog.py', '.'),
        ]
        
        for src, dest in data_files:
            src_path = self.project_dir / src
            if src_path.exists():
                cmd.append(f'--add-data={src_path};{dest}')
                print(f"  - 添加數據: {src}")
        
        # 添加圖示到打包檔案中
        if self.icon_file.exists():
            cmd.append(f'--add-data={self.icon_file};.')
        
        # 隱藏導入模組
        hidden_imports = [
            # 快捷鍵與輸入控制
            'keyboard', 'keyboard._winkeyboard', 'keyboard._canonical_names',
            # Windows API
            'win32gui', 'win32con', 'win32api', 'win32process',
            'win32com', 'win32com.shell', 'win32com.shell.shell',
            'pythoncom', 'pywintypes',
            # GUI
            'ttkbootstrap', 'tkinter', 'tkinter.filedialog', 
            'tkinter.messagebox', 'tkinter.font',
            # 更新系統
            'update_manager', 'update_dialog',
            # 網路與系統
            'urllib', 'urllib.request', 'urllib.error',
            'json', 'zipfile', 'tempfile', 'shutil',
            # 其他必需模組
            'threading', 'subprocess', 'atexit', 'functools',
            'time', 'os', 'sys', 'pathlib'
        ]
        for module in hidden_imports:
            cmd.append(f'--hidden-import={module}')
        
        # 主文件
        cmd.append(str(self.main_file))
        
        # 執行打包
        print(f"  執行 PyInstaller...")
        try:
            result = subprocess.run(cmd, cwd=str(self.project_dir), 
                                   capture_output=False, text=True)
            
            if result.returncode != 0:
                raise Exception("主程式打包失敗")
        except Exception as e:
            print(f"\n錯誤: {e}")
            print("\n可能的原因：")
            print("  1. dist 目錄中的檔案被其他程序佔用")
            print("  2. 請關閉 ChroLens_Portal 後重試")
            raise
        
        print("  主程式打包完成\n")
    
    def copy_files(self):
        """複製必要文件到輸出目錄"""
        print("\n[3/5] 複製必要文件...")
        
        # 創建必要的目錄
        backup_dir = self.output_dir / "backup"
        backup_dir.mkdir(exist_ok=True)
        
        # 創建空的配置文件
        config_file = self.output_dir / "chrolens_portal.json"
        if not config_file.exists():
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump({
                    "folder": "",
                    "interval": "1.0",
                    "group_display_names": {},
                    "group_hotkeys": [],
                    "checkbox_entries": [],
                    "schedule_tasks": [],
                    "window_layouts": {}
                }, f, ensure_ascii=False, indent=2)
            print(f"  - 創建 chrolens_portal.json")
        
        print("  必要文件複製完成\n")
    
    def create_version_file(self):
        """創建版本文件"""
        print("\n[4/5] 創建版本文件...")
        
        version_file = self.output_dir / f"version{self.version}.txt"
        
        with open(version_file, 'w', encoding='utf-8') as f:
            f.write(f"ChroLens_Portal v{self.version}\n")
            f.write(f"打包時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("\n")
            f.write("=== 功能特色 ===\n")
            f.write("- 分組啟動/關閉程式\n")
            f.write("- 分組視窗置頂切換\n")
            f.write("- 視窗佈局記憶（類似 FancyZones）\n")
            f.write("- 全域快捷鍵支援\n")
            f.write("- 排程任務功能\n")
            f.write("- Mini 模式\n")
            f.write("- 自動更新系統\n")
        
        print(f"  - 創建 version{self.version}.txt")
        print("  版本文件創建完成\n")
    
    def create_zip(self):
        """創建 ZIP 壓縮包"""
        print("\n[5/5] 創建 ZIP 壓縮包...")
        
        zip_filename = f"ChroLens_Portal_v{self.version}.zip"
        zip_path = self.dist_dir / zip_filename
        
        # 刪除舊的 ZIP（如果存在）
        if zip_path.exists():
            zip_path.unlink()
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # 遍歷輸出目錄
            for root, dirs, files in os.walk(self.output_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(self.output_dir.parent)
                    zipf.write(file_path, arcname)
                    
        file_size = zip_path.stat().st_size / (1024 * 1024)
        print(f"  - 創建 {zip_filename} ({file_size:.2f} MB)")
        print("  ZIP 壓縮包創建完成\n")
        
        return zip_path
    
    def build(self):
        """執行完整打包流程"""
        try:
            self.clean()
            self.build_main()
            self.copy_files()
            self.create_version_file()
            zip_path = self.create_zip()
            
            print("\n" + "="*50)
            print("✓ 打包完成！")
            print("="*50)
            print(f"\n輸出目錄: {self.output_dir}")
            print(f"ZIP 檔案: {zip_path}")
            print(f"\n可以將 ZIP 檔案上傳至 GitHub Releases")
            print("\n")
            
        except Exception as e:
            print(f"\n✗ 打包失敗: {e}")
            sys.exit(1)

if __name__ == "__main__":
    builder = PortalBuilder()
    builder.build()
    
    input("\n按 Enter 鍵退出...")
