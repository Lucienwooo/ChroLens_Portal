"""檔案開啟模組"""
import os
import time
import subprocess
import pythoncom
from win32com.shell import shell
from typing import List, Tuple, Optional

class FileOpener:
    """檔案開啟管理器"""
    
    def __init__(self, logger=None):
        self.logger = logger
    
    def log(self, msg: str):
        """記錄日誌"""
        if self.logger:
            self.logger.log(msg)
        else:
            print(msg)
    
    def open_lnk_target(self, lnk_path: str) -> Tuple[Optional[str], Optional[str]]:
        """解析 .lnk 捷徑檔案
        
        Args:
            lnk_path: 捷徑檔案路徑
            
        Returns:
            (目標路徑, 參數字串) 的元組
        """
        try:
            pythoncom.CoInitialize()
            shell_link = pythoncom.CoCreateInstance(
                shell.CLSID_ShellLink, None,
                pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink
            )
            persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
            persist_file.Load(lnk_path)
            target_path, _ = shell_link.GetPath(shell.SLGP_UNCPRIORITY)
            arguments = shell_link.GetArguments()
            return target_path, arguments
        except Exception as e:
            self.log(f"解析捷徑失敗 {lnk_path}: {e}")
            return None, None
    
    def open_file(self, file_path: str) -> bool:
        """開啟檔案
        
        Args:
            file_path: 檔案路徑
            
        Returns:
            是否成功開啟
        """
        if not os.path.exists(file_path):
            self.log(f"找不到檔案: {file_path}")
            return False
        
        try:
            if file_path.lower().endswith('.lnk'):
                target, args = self.open_lnk_target(file_path)
                if target and os.path.exists(target):
                    self.log(f"開啟捷徑目標: {target} {args}")
                    subprocess.Popen(
                        f'"{target}" {args}',
                        shell=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    return True
                else:
                    self.log(f"無法解析捷徑或目標不存在: {file_path}")
                    return False
            else:
                self.log(f"開啟檔案: {file_path}")
                os.startfile(file_path)
                return True
        except Exception as e:
            self.log(f"無法開啟檔案: {file_path}，錯誤：{e}")
            return False
    
    def open_files_with_interval(self, folder: str, files: List[str], interval: float = 1.0):
        """依序開啟多個檔案，每個檔案之間間隔指定秒數
        
        Args:
            folder: 資料夾路徑
            files: 檔案名稱列表
            interval: 間隔秒數
        """
        success_count = 0
        for file in files:
            file_path = os.path.join(folder, file)
            if self.open_file(file_path):
                success_count += 1
            time.sleep(interval)
        
        self.log(f"已開啟 {success_count}/{len(files)} 個檔案")
