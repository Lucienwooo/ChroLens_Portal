"""日誌系統模組"""
import time
from typing import Optional, Callable

class Logger:
    """日誌管理器"""
    
    def __init__(self):
        self.history = []
        self.callback: Optional[Callable[[str], None]] = None
    
    def set_callback(self, callback: Callable[[str], None]):
        """設定日誌輸出回調函數
        
        Args:
            callback: 接收日誌訊息的函數
        """
        self.callback = callback
    
    def log(self, msg: str):
        """記錄日誌
        
        Args:
            msg: 日誌訊息
        """
        timestamp = time.strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {msg}"
        self.history.append(full_msg)
        
        # 限制歷史記錄大小
        if len(self.history) > 1000:
            self.history = self.history[-500:]
        
        if self.callback:
            try:
                self.callback(full_msg)
            except Exception as e:
                print(f"Logger callback error: {e}")
    
    def get_history(self) -> list:
        """取得日誌歷史記錄
        
        Returns:
            日誌歷史列表
        """
        return self.history.copy()
    
    def clear(self):
        """清空日誌歷史"""
        self.history.clear()
