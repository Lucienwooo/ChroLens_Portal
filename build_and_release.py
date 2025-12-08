"""
ChroLens_Portal 自動打包並發布到 GitHub
整合打包、壓縮、上傳到 GitHub Releases 的完整流程

使用方法:
1. 首次使用需要設定 GitHub Token (一次性設定)
2. 更新 CHANGELOG.md 中的版本紀錄
3. 執行此腳本會自動完成打包並上傳到 GitHub
   - 從 ChroLens_Portal.py 讀取 CURRENT_VERSION
   - 從 CHANGELOG.md 讀取對應版本的更新說明
   - 自動生成 Release Notes 並上傳

需要安裝:
pip install PyGithub
"""

import os
import sys
import json
import shutil
import subprocess
import zipfile
from pathlib import Path
from datetime import datetime
import getpass

try:
    from github import Github, GithubException
except ImportError:
    print("錯誤: 需要安裝 PyGithub")
    print("請執行: pip install PyGithub")
    sys.exit(1)


class PortalReleaseBuilder:
    """Portal 打包與發布工具"""
    
    def __init__(self):
        # 專案目錄
        self.project_dir = Path(__file__).parent
        self.main_file = self.project_dir / "ChroLens_Portal.py"
        self.icon_file = self.project_dir / "冥想貓貓.ico"
        
        # 輸出目錄
        self.build_dir = self.project_dir / "build"
        self.dist_dir = self.project_dir / "dist"
        self.output_dir = self.dist_dir / "ChroLens_Portal"
        
        # GitHub 設定
        self.github_repo = "Lucienwooo/ChroLens_Portal"
        self.token_file = self.project_dir / ".github_token"
        
        # 讀取版本號
        self.version = self._read_version()
        
        print(f"\n{'='*60}")
        print(f"ChroLens_Portal 自動打包與發布工具")
        print(f"版本: {self.version}")
        print(f"{'='*60}\n")
    
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
            return "2.5.1"
    
    def _get_github_token(self) -> str:
        """獲取 GitHub Token"""
        # 檢查是否已存在 token
        if self.token_file.exists():
            try:
                with open(self.token_file, 'r') as f:
                    token = f.read().strip()
                    if token:
                        return token
            except:
                pass
        
        # 直接使用預設 token
        token = "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"
        
        # 保存 token
        try:
            with open(self.token_file, 'w') as f:
                f.write(token)
            # 設定檔案為只讀（安全性）
            os.chmod(self.token_file, 0o600)
        except:
            pass
        
        return token
    
    def _extract_changelog(self) -> str:
        """從 CHANGELOG.md 提取當前版本的更新說明"""
        changelog_file = self.project_dir / "CHANGELOG.md"
        
        if not changelog_file.exists():
            print("  警告: 找不到 CHANGELOG.md")
            return "本次更新包含功能改進與錯誤修復"
        
        try:
            with open(changelog_file, 'r', encoding='utf-8') as f:
                content = f.read()
                lines = content.split('\n')
                
                # 尋找當前版本的區段
                in_version_section = False
                changelog_lines = []
                
                for line in lines:
                    # 找到當前版本標題
                    if line.startswith(f"## [{self.version}]"):
                        in_version_section = True
                        continue
                    
                    # 如果遇到下一個版本標題，停止
                    if in_version_section and line.startswith("## ["):
                        break
                    
                    # 收集版本內容
                    if in_version_section:
                        line = line.strip()
                        # 保留所有內容，包括小標題和列表項目
                        if line:
                            # 保留 ### 標題
                            if line.startswith('### '):
                                changelog_lines.append('')  # 空行分隔
                                changelog_lines.append('**' + line[4:] + '**')  # 轉換為粗體
                            # 保留列表項目
                            elif line.startswith('- '):
                                changelog_lines.append(line)
                            # 保留其他文字
                            elif not line.startswith('#'):
                                changelog_lines.append(line)
                
                if changelog_lines:
                    return '\n'.join(changelog_lines)
                else:
                    print(f"  警告: 在 CHANGELOG.md 中找不到版本 {self.version} 的記錄")
                    return "本次更新包含功能改進與錯誤修復"
        
        except Exception as e:
            print(f"  警告: 無法讀取 CHANGELOG.md: {e}")
            return "本次更新包含功能改進與錯誤修復"
    
    def _format_release_notes(self, version_description: str) -> str:
        """格式化 Release Notes"""
        notes = f"# ChroLens Portal v{self.version}\n\n"
        
        # 版本更新說明（從代碼中提取）
        notes += f"## 📝 更新內容\n\n"
        notes += f"{version_description}\n\n"
        
        # 安裝說明
        notes += "## 📦 安裝方式\n\n"
        notes += "### 方式一：自動更新（推薦）\n"
        notes += "1. 開啟 ChroLens Portal\n"
        notes += "2. 點擊「檢查更新」按鈕\n"
        notes += "3. 程式會自動下載並安裝更新\n\n"
        
        notes += "### 方式二：手動安裝\n"
        notes += f"1. 下載 `ChroLens_Portal_v{self.version}.zip`\n"
        notes += "2. 解壓縮到任意位置\n"
        notes += "3. 執行 `ChroLens_Portal.exe`\n\n"
        
        notes += "---\n\n"
        notes += f"📅 發布時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        notes += f"💻 適用系統: Windows 10/11\n"
        
        return notes
    
    def clean(self):
        """清理舊檔案"""
        print("\n[1/6] 清理舊檔案...")
        
        for dir_path in [self.build_dir, self.dist_dir]:
            if dir_path.exists():
                print(f"  - 刪除 {dir_path.name}/")
                try:
                    shutil.rmtree(dir_path, ignore_errors=False)
                except Exception as e:
                    print(f"  ⚠ 警告: {e}")
        
        print("  ✓ 清理完成\n")
    
    def build_main(self):
        """打包主程式"""
        print("\n[2/6] 打包主程式...")
        
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
        
        # 添加圖示到打包檔案中
        if self.icon_file.exists():
            cmd.append(f'--add-data={self.icon_file};.')
        
        # 隱藏導入模組
        hidden_imports = [
            'keyboard', 'win32gui', 'win32con', 'win32api', 
            'win32process', 'win32com', 'win32com.shell', 
            'ttkbootstrap', 'update_manager', 'update_dialog',
        ]
        for module in hidden_imports:
            cmd.append(f'--hidden-import={module}')
        
        # 主文件
        cmd.append(str(self.main_file))
        
        # 執行打包
        print(f"  執行 PyInstaller...")
        result = subprocess.run(cmd, cwd=str(self.project_dir), 
                               capture_output=True, text=True)
        
        if result.returncode != 0:
            print(f"  錯誤: {result.stderr}")
            raise Exception("主程式打包失敗")
        
        print("  ✓ 主程式打包完成\n")
    
    def copy_files(self):
        """複製必要文件"""
        print("\n[3/6] 複製必要文件...")
        
        # 創建配置文件
        backup_dir = self.output_dir / "backup"
        backup_dir.mkdir(exist_ok=True)
        
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
        
        print("  ✓ 必要文件複製完成\n")
    
    def create_version_file(self):
        """創建版本文件"""
        print("\n[4/6] 創建版本文件...")
        
        version_file = self.output_dir / f"version{self.version}.txt"
        
        with open(version_file, 'w', encoding='utf-8') as f:
            f.write(f"ChroLens_Portal v{self.version}\n")
            f.write(f"打包時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        
        print(f"  ✓ version{self.version}.txt 已創建\n")
    
    def create_zip(self) -> Path:
        """創建 ZIP 壓縮包"""
        print("\n[5/6] 創建 ZIP 壓縮包...")
        
        zip_filename = f"ChroLens_Portal_v{self.version}.zip"
        zip_path = self.dist_dir / zip_filename
        
        if zip_path.exists():
            zip_path.unlink()
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(self.output_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(self.output_dir.parent)
                    zipf.write(file_path, arcname)
        
        file_size = zip_path.stat().st_size / (1024 * 1024)
        print(f"  ✓ {zip_filename} ({file_size:.2f} MB)\n")
        
        return zip_path
    
    def create_github_release(self, zip_path: Path):
        """創建 GitHub Release 並上傳檔案"""
        print("\n[6/6] 發布到 GitHub...")
        
        # 獲取 Token
        token = self._get_github_token()
        
        # 連接 GitHub
        try:
            g = Github(token)
            repo = g.get_repo(self.github_repo)
            print(f"  ✓ 已連接到 {self.github_repo}")
        except GithubException as e:
            print(f"  ✗ GitHub 認證失敗: {e}")
            print("  請檢查 Token 權限或重新設定")
            return False
        
        # 檢查 Release 是否已存在
        tag_name = f"v{self.version}"
        try:
            existing_release = repo.get_release(tag_name)
            print(f"  ⚠ Release {tag_name} 已存在，自動刪除並重新創建...")
            existing_release.delete_release()
            print(f"  ✓ 已刪除舊的 Release")
        except GithubException:
            pass  # Release 不存在，繼續
        
        # 提取更新說明
        version_description = self._extract_changelog()
        release_notes = self._format_release_notes(version_description)
        
        # 創建 Release
        try:
            print(f"  正在創建 Release {tag_name}...")
            release = repo.create_git_release(
                tag=tag_name,
                name=f"ChroLens Portal v{self.version}",
                message=release_notes,
                draft=False,
                prerelease=False
            )
            print(f"  ✓ Release 已創建")
        except GithubException as e:
            print(f"  ✗ 創建 Release 失敗: {e}")
            return False
        
        # 上傳 ZIP 檔案
        try:
            print(f"  正在上傳 {zip_path.name}...")
            # 使用 upload_asset 而不是 upload_asset_from_memory
            release.upload_asset(
                str(zip_path),
                label=zip_path.name,
                content_type='application/zip'
            )
            print(f"  ✓ 檔案已上傳")
        except GithubException as e:
            print(f"  ✗ 上傳失敗: {e}")
            return False
        
        print(f"\n  🎉 發布成功!")
        print(f"  🔗 查看 Release: https://github.com/{self.github_repo}/releases/tag/{tag_name}")
        
        return True
    
    def _validate_before_build(self):
        """打包前驗證"""
        print("\n[0/6] 打包前驗證...")
        
        # 檢查 CHANGELOG.md 是否包含當前版本
        changelog_file = self.project_dir / "CHANGELOG.md"
        if changelog_file.exists():
            with open(changelog_file, 'r', encoding='utf-8') as f:
                content = f.read()
                if f"## [{self.version}]" not in content:
                    print(f"  ⚠ 警告: CHANGELOG.md 中找不到版本 {self.version}")
                    print(f"  請先更新 CHANGELOG.md")
                    return False
                else:
                    print(f"  ✓ CHANGELOG.md 包含版本 {self.version}")
        else:
            print(f"  ⚠ 警告: 找不到 CHANGELOG.md")
        
        # 檢查版本號格式
        import re
        if not re.match(r'^\d+\.\d+(\.\d+)?$', self.version):
            print(f"  ⚠ 警告: 版本號格式不正確: {self.version}")
            return False
        else:
            print(f"  ✓ 版本號格式正確: {self.version}")
        
        print("  ✓ 驗證通過\n")
        return True
    
    def build_and_release(self):
        """執行完整流程"""
        try:
            # 驗證
            if not self._validate_before_build():
                print("\n驗證失敗，已取消打包")
                sys.exit(1)
            
            self.clean()
            self.build_main()
            self.copy_files()
            self.create_version_file()
            zip_path = self.create_zip()
            
            # 自動上傳到 GitHub（不詢問）
            print("\n" + "="*60)
            print("正在自動上傳到 GitHub Releases...")
            print("="*60)
            
            success = self.create_github_release(zip_path)
            
            if success:
                print("\n" + "="*60)
                print("✅ 打包與發布完成！")
                print("="*60)
            else:
                print("\n" + "="*60)
                print("⚠ 打包完成，但發布失敗")
                print(f"ZIP 檔案: {zip_path}")
                print("請手動上傳到 GitHub")
                print("="*60)
            
            print()
            
        except Exception as e:
            print(f"\n✗ 錯誤: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)


if __name__ == "__main__":
    builder = PortalReleaseBuilder()
    builder.build_and_release()
    
    input("\n按 Enter 鍵退出...")
