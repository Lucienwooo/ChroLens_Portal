"""設定檔管理模組"""
import json
import os
from typing import Dict, Any, Optional

class ConfigManager:
    """設定檔管理器"""
    
    SETTINGS_FILE = "chrolens_portal.json"
    
    def __init__(self):
        self.data: Dict[str, Any] = {}
        self._pending_save = False
        self._save_timer = None
    
    def load(self) -> Dict[str, Any]:
        """載入設定檔
        
        Returns:
            設定資料字典
        """
        if not os.path.exists(self.SETTINGS_FILE):
            return {}
        
        try:
            with open(self.SETTINGS_FILE, "r", encoding="utf-8") as f:
                self.data = json.load(f)
            return self.data
        except Exception as e:
            print(f"設定檔讀取失敗: {e}")
            return {}
    
    def save(self, data: Optional[Dict[str, Any]] = None):
        """儲存設定檔
        
        Args:
            data: 要儲存的設定資料，若為 None 則使用內部資料
        """
        if data is not None:
            self.data = data
        
        try:
            with open(self.SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"儲存設定檔失敗: {e}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """取得設定值
        
        Args:
            key: 設定鍵名
            default: 預設值
            
        Returns:
            設定值
        """
        return self.data.get(key, default)
    
    def set(self, key: str, value: Any):
        """設定值
        
        Args:
            key: 設定鍵名
            value: 設定值
        """
        self.data[key] = value
    
    def update(self, data: Dict[str, Any]):
        """批次更新設定
        
        Args:
            data: 要更新的設定資料
        """
        self.data.update(data)
    
    def debounce_save(self, data: Dict[str, Any], delay: float = 1.0, callback=None):
        """延遲儲存（防抖動）
        
        Args:
            data: 要儲存的資料
            delay: 延遲秒數
            callback: 儲存後的回調函數
        """
        import threading
        
        if self._save_timer:
            self._save_timer.cancel()
        
        def _save():
            self.save(data)
            if callback:
                callback()
        
        self._save_timer = threading.Timer(delay, _save)
        self._save_timer.daemon = True
        self._save_timer.start()
