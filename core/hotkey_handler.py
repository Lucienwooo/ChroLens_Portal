"""熱鍵處理模組"""
import keyboard
import functools
from typing import Dict, Callable, Optional

class HotkeyHandler:
    """熱鍵管理器"""
    
    def __init__(self, logger=None):
        self.logger = logger
        self.handlers: Dict[str, any] = {}
    
    def log(self, msg: str):
        """記錄日誌"""
        if self.logger:
            self.logger.log(msg)
    
    def register_hotkey(self, hotkey: str, callback: Callable, group_code: str = None):
        """註冊全域熱鍵
        
        Args:
            hotkey: 熱鍵組合（如 "Alt+1"）
            callback: 觸發時的回調函數
            group_code: 分組代碼（用於識別）
        """
        key = group_code if group_code else hotkey
        
        # 移除舊的熱鍵
        if key in self.handlers:
            try:
                keyboard.remove_hotkey(self.handlers[key])
            except Exception:
                pass
        
        try:
            handler = keyboard.add_hotkey(
                hotkey,
                callback,
                suppress=False,
                trigger_on_release=False
            )
            self.handlers[key] = handler
            if self.logger:
                self.log(f"已註冊熱鍵：{hotkey} ({key})")
        except Exception as e:
            if self.logger:
                self.log(f"註冊熱鍵失敗 {hotkey}: {e}")
    
    def unregister_hotkey(self, key: str):
        """取消註冊熱鍵
        
        Args:
            key: 熱鍵識別碼或分組代碼
        """
        if key in self.handlers:
            try:
                keyboard.remove_hotkey(self.handlers[key])
                del self.handlers[key]
                if self.logger:
                    self.log(f"已取消熱鍵：{key}")
            except Exception as e:
                if self.logger:
                    self.log(f"取消熱鍵失敗 {key}: {e}")
    
    def unregister_all(self):
        """取消所有熱鍵"""
        for key in list(self.handlers.keys()):
            self.unregister_hotkey(key)
    
    def update_hotkeys(self, hotkey_map: Dict[str, tuple]):
        """批次更新熱鍵
        
        Args:
            hotkey_map: {group_code: (hotkey, callback)} 的字典
        """
        # 先移除所有舊熱鍵
        self.unregister_all()
        
        # 註冊新熱鍵
        for group_code, (hotkey, callback) in hotkey_map.items():
            self.register_hotkey(hotkey, callback, group_code)
