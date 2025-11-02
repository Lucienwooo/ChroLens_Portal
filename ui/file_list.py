"""檔案與視窗列表 UI 元件"""
import tkinter as tk
import ttkbootstrap as tb
from tkinter import font as tkfont
from typing import List, Callable, Optional

class DraggableListFrame:
    """可拖曳的列表框架"""
    
    def __init__(self, parent, title: str, on_drag_start: Optional[Callable] = None):
        self.parent = parent
        self.title = title
        self.on_drag_start = on_drag_start
        
        self.outer_frame = tb.Frame(parent)
        self.outer_frame.grid_propagate(True)
        self.outer_frame.grid_rowconfigure(0, weight=1)
        self.outer_frame.grid_columnconfigure(0, weight=1)
        
        self.canvas = tk.Canvas(self.outer_frame, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        
        self.inner_frame = tb.Frame(self.canvas)
        self.inner_frame_id = self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.inner_frame.grid_columnconfigure(0, weight=1)
        
        self.scrollbar = tb.Scrollbar(self.outer_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self._on_vsb)
        
        self.inner_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))
    
    def _on_vsb(self, *args):
        self.scrollbar.set(*args)
        if float(args[0]) <= 0.0 and float(args[1]) >= 1.0:
            self.scrollbar.grid_remove()
        else:
            self.scrollbar.grid(row=0, column=1, sticky="ns")
    
    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.itemconfig(self.inner_frame_id, width=self.canvas.winfo_width())
    
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def grid(self, **kwargs):
        """將框架放置到父容器"""
        self.outer_frame.grid(**kwargs)
    
    def update_items(self, items: List[str]):
        """更新列表項目
        
        Args:
            items: 項目列表
        """
        # 清空現有項目
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        
        # 新增項目
        for row, item in enumerate(items):
            lbl = tb.Label(
                self.inner_frame,
                text=item,
                anchor="w",
                font=tkfont.Font(family="Microsoft JhengHei", size=10)
            )
            lbl.grid(row=row, column=0, sticky="ew", padx=2, pady=1)
            
            if self.on_drag_start:
                lbl.bind("<ButtonPress-1>", lambda e, t=item: self.on_drag_start(e, t))
        
        self.inner_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))


class FileListPanel:
    """檔案列表面板"""
    
    def __init__(self, parent, on_drag_start: Optional[Callable] = None):
        self.list_frame = DraggableListFrame(parent, "檔案列表", on_drag_start)
    
    def grid(self, **kwargs):
        self.list_frame.grid(**kwargs)
    
    def update(self, files: List[str]):
        """更新檔案列表
        
        Args:
            files: 檔案名稱列表
        """
        self.list_frame.update_items(files)


class WindowListPanel:
    """視窗列表面板"""
    
    def __init__(self, parent, on_drag_start: Optional[Callable] = None):
        self.list_frame = DraggableListFrame(parent, "視窗列表", on_drag_start)
    
    def grid(self, **kwargs):
        self.list_frame.grid(**kwargs)
    
    def update(self, windows: List[str]):
        """更新視窗列表
        
        Args:
            windows: 視窗標題列表
        """
        self.list_frame.update_items(windows)
