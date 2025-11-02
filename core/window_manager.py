"""視窗管理模組"""
import win32gui
import win32con
from typing import List, Dict, Tuple, Optional, Callable
import json
import os

class WindowLayout:
    """視窗布局資料類別"""
    
    def __init__(self, title: str, x: int, y: int, width: int, height: int, is_maximized: bool = False):
        self.title = title
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.is_maximized = is_maximized
    
    def to_dict(self) -> dict:
        return {
            'title': self.title,
            'x': self.x,
            'y': self.y,
            'width': self.width,
            'height': self.height,
            'is_maximized': self.is_maximized
        }
    
    @staticmethod
    def from_dict(data: dict) -> 'WindowLayout':
        return WindowLayout(
            data['title'],
            data['x'],
            data['y'],
            data['width'],
            data['height'],
            data.get('is_maximized', False)
        )


class WindowManager:
    """視窗管理器 - 負責視窗操作與布局管理"""
    
    LAYOUT_FILE = "window_layouts.json"
    
    def __init__(self, logger=None):
        self.logger = logger
        self.group_focus_indexes: Dict[str, int] = {}
        self.group_hwnd_lists: Dict[str, List[int]] = {}
        self.layouts: Dict[str, List[WindowLayout]] = {}  # 儲存每個分組的布局
        self.load_layouts()
    
    def log(self, msg: str):
        """記錄日誌"""
        if self.logger:
            self.logger.log(msg)
        else:
            print(msg)
    
    def get_taskbar_window_titles(self) -> List[str]:
        """取得所有任務列視窗標題
        
        Returns:
            視窗標題列表
        """
        exclude_keywords = [
            "設定", "windows 輸入體驗", "windows input experience", "searchui", 
            "cortana", "開始功能表", "start menu", "工作管理員", "task manager", 
            "lockapp", "shell experience host", "runtimebroker", "searchapp"
        ]
        titles = []
        
        def enum_handler(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title and title.strip():
                    title_lower = title.strip().lower()
                    if not any(keyword in title_lower for keyword in exclude_keywords):
                        titles.append(title)
        
        win32gui.EnumWindows(enum_handler, None)
        return titles
    
    def find_windows_by_titles(self, titles: List[str], my_hwnd: int = 0) -> List[int]:
        """根據標題關鍵字查找視窗
        
        Args:
            titles: 標題關鍵字列表
            my_hwnd: 要排除的視窗句柄（通常是主程式視窗）
            
        Returns:
            視窗句柄列表
        """
        target_titles = [title.lower() for title in titles if title]
        hwnds = []
        
        def enum_handler(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            if hwnd == my_hwnd:
                return
            window_text = win32gui.GetWindowText(hwnd).lower().strip()
            if any(title and title in window_text for title in target_titles):
                hwnds.append(hwnd)
        
        win32gui.EnumWindows(enum_handler, None)
        return hwnds
    
    def focus_window(self, hwnd: int) -> bool:
        """聚焦視窗
        
        Args:
            hwnd: 視窗句柄
            
        Returns:
            是否成功
        """
        try:
            # 如果視窗最小化，先還原
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            # 置頂視窗
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_TOP, 0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
            )
            
            # 嘗試設為前景視窗
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            
            return True
        except Exception as e:
            self.log(f"聚焦視窗失敗: {e}")
            return False
    
    def close_windows_by_titles(self, titles: List[str]) -> int:
        """根據標題關鍵字關閉視窗
        
        Args:
            titles: 標題關鍵字列表
            
        Returns:
            關閉的視窗數量
        """
        keywords = []
        for filename in titles:
            filename = filename.strip().lower()
            if filename:
                keywords.append(filename)
                if "." in filename:
                    keywords.append(os.path.splitext(filename)[0])
        keywords = list(set(keywords))
        
        closed_count = 0
        
        def enum_handler(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                window_text_lower = window_text.lower()
                if any(kw and kw in window_text_lower for kw in keywords):
                    try:
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                        self.log(f"已關閉視窗：{window_text}")
                        nonlocal closed_count
                        closed_count += 1
                    except Exception as e:
                        self.log(f"關閉視窗失敗：{window_text} ({e})")
        
        win32gui.EnumWindows(enum_handler, None)
        return closed_count
    
    def update_group_windows(self, group_code: str, file_titles: List[str], my_hwnd: int = 0):
        """更新分組視窗列表
        
        Args:
            group_code: 分組代碼
            file_titles: 檔案標題列表
            my_hwnd: 主程式視窗句柄
        """
        if group_code not in self.group_focus_indexes:
            self.group_focus_indexes[group_code] = 0
        
        self.group_hwnd_lists[group_code] = self.find_windows_by_titles(file_titles, my_hwnd)
    
    def focus_next_in_group(self, group_code: str) -> Optional[str]:
        """切換到分組中的下一個視窗
        
        Args:
            group_code: 分組代碼
            
        Returns:
            切換到的視窗標題，若無視窗則返回 None
        """
        hwnds = self.group_hwnd_lists.get(group_code, [])
        if not hwnds:
            return None
        
        idx = self.group_focus_indexes.get(group_code, 0)
        hwnd = hwnds[idx % len(hwnds)]
        
        if self.focus_window(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            self.group_focus_indexes[group_code] = (idx + 1) % len(hwnds)
            return window_title
        
        return None
    
    # ===== 視窗布局管理功能 =====
    
    def get_window_rect(self, hwnd: int) -> Optional[Tuple[int, int, int, int]]:
        """取得視窗位置與大小
        
        Args:
            hwnd: 視窗句柄
            
        Returns:
            (x, y, width, height) 或 None
        """
        try:
            rect = win32gui.GetWindowRect(hwnd)
            x, y, right, bottom = rect
            return (x, y, right - x, bottom - y)
        except Exception:
            return None
    
    def set_window_rect(self, hwnd: int, x: int, y: int, width: int, height: int) -> bool:
        """設定視窗位置與大小
        
        Args:
            hwnd: 視窗句柄
            x, y: 位置座標
            width, height: 寬度與高度
            
        Returns:
            是否成功
        """
        try:
            # 如果視窗最大化，先還原
            if win32gui.IsIconic(hwnd) or win32gui.IsZoomed(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_TOP,
                x, y, width, height,
                win32con.SWP_NOZORDER
            )
            return True
        except Exception as e:
            self.log(f"設定視窗位置失敗: {e}")
            return False
    
    def save_group_layout(self, group_code: str, file_titles: List[str], my_hwnd: int = 0):
        """儲存分組的視窗布局
        
        Args:
            group_code: 分組代碼
            file_titles: 檔案標題列表
            my_hwnd: 主程式視窗句柄
        """
        hwnds = self.find_windows_by_titles(file_titles, my_hwnd)
        if not hwnds:
            self.log(f"分組 {group_code} 沒有可儲存的視窗")
            return
        
        layouts = []
        for hwnd in hwnds:
            try:
                title = win32gui.GetWindowText(hwnd)
                is_maximized = win32gui.IsZoomed(hwnd)
                
                if is_maximized:
                    # 如果最大化，先還原以取得原始大小
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    rect = self.get_window_rect(hwnd)
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                else:
                    rect = self.get_window_rect(hwnd)
                
                if rect:
                    x, y, width, height = rect
                    layout = WindowLayout(title, x, y, width, height, is_maximized)
                    layouts.append(layout)
                    self.log(f"已儲存視窗布局：{title}")
            except Exception as e:
                self.log(f"儲存視窗布局失敗: {e}")
        
        if layouts:
            self.layouts[group_code] = layouts
            self.save_layouts()
            self.log(f"已儲存分組 {group_code} 的 {len(layouts)} 個視窗布局")
    
    def restore_group_layout(self, group_code: str, file_titles: List[str], my_hwnd: int = 0):
        """還原分組的視窗布局
        
        Args:
            group_code: 分組代碼
            file_titles: 檔案標題列表
            my_hwnd: 主程式視窗句柄
        """
        if group_code not in self.layouts:
            self.log(f"分組 {group_code} 沒有儲存的布局")
            return
        
        hwnds = self.find_windows_by_titles(file_titles, my_hwnd)
        if not hwnds:
            self.log(f"分組 {group_code} 沒有開啟的視窗")
            return
        
        layouts = self.layouts[group_code]
        matched = 0
        
        # 建立視窗標題到句柄的映射
        hwnd_map = {}
        for hwnd in hwnds:
            title = win32gui.GetWindowText(hwnd).lower()
            hwnd_map[title] = hwnd
        
        for layout in layouts:
            # 嘗試匹配視窗
            target_hwnd = None
            layout_title_lower = layout.title.lower()
            
            for title, hwnd in hwnd_map.items():
                if layout_title_lower in title or title in layout_title_lower:
                    target_hwnd = hwnd
                    break
            
            if target_hwnd:
                if self.set_window_rect(target_hwnd, layout.x, layout.y, layout.width, layout.height):
                    if layout.is_maximized:
                        win32gui.ShowWindow(target_hwnd, win32con.SW_MAXIMIZE)
                    matched += 1
                    self.log(f"已還原視窗：{layout.title}")
        
        if matched > 0:
            self.log(f"已還原分組 {group_code} 的 {matched} 個視窗布局")
        else:
            self.log(f"無法還原分組 {group_code} 的視窗布局（無匹配視窗）")
    
    def load_layouts(self):
        """載入視窗布局設定"""
        if not os.path.exists(self.LAYOUT_FILE):
            return
        
        try:
            with open(self.LAYOUT_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            self.layouts = {}
            for group_code, layouts_data in data.items():
                self.layouts[group_code] = [
                    WindowLayout.from_dict(layout_data) 
                    for layout_data in layouts_data
                ]
        except Exception as e:
            self.log(f"載入視窗布局失敗: {e}")
    
    def save_layouts(self):
        """儲存視窗布局設定"""
        try:
            data = {}
            for group_code, layouts in self.layouts.items():
                data[group_code] = [layout.to_dict() for layout in layouts]
            
            with open(self.LAYOUT_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"儲存視窗布局失敗: {e}")
    
    def has_saved_layout(self, group_code: str) -> bool:
        """檢查分組是否有儲存的布局
        
        Args:
            group_code: 分組代碼
            
        Returns:
            是否有儲存的布局
        """
        return group_code in self.layouts and len(self.layouts[group_code]) > 0
