"""主視窗模組"""
import os
import sys
import threading
import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, font as tkfont
import atexit
import functools

# 匯入自定義模組
from utils import ConfigManager, Logger
from core import WindowManager, HotkeyHandler, FileOpener
from ui.file_list import FileListPanel, WindowListPanel


def resource_path(relative_path):
    """取得資源檔案路徑（支援打包後的執行檔）"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class MainWindow:
    """ChroLens_Portal 主視窗"""
    
    def __init__(self):
        # 初始化核心元件
        self.logger = Logger()
        self.config = ConfigManager()
        self.window_manager = WindowManager(self.logger)
        self.hotkey_handler = HotkeyHandler(self.logger)
        self.file_opener = FileOpener(self.logger)
        
        # 分組設定
        self.group_codes = ["A", "B", "C", "D", "E", "F"]
        self.default_hotkeys = ["Alt+1", "Alt+2", "Alt+3", "Alt+4", "Alt+Q", "Alt+W"]
        
        # 建立主視窗
        self.app = tb.Window(themename="darkly")
        self.app.title("ChroLens_Portal 2.3 (模組化版)")
        self.app.geometry("1200x900")
        
        # 設定圖示
        try:
            ico_path = resource_path("冥想貓貓.ico")
            self.app.iconbitmap(ico_path)
        except Exception as e:
            print(f"無法設定 icon: {e}")
        
        # 介面變數
        self.folder_var = tk.StringVar(value="")
        self.interval_var = tk.StringVar(value="4")
        self.group_display_names = {c: tk.StringVar(value=c) for c in self.group_codes}
        self.group_hotkeys = [tk.StringVar(value=self.default_hotkeys[i]) for i in range(6)]
        
        # UI 元件容器
        self.group_buttons = {}
        self.close_buttons = {}
        self.show_label_frames = []
        self.checkbox_vars_entries = []
        
        # 拖曳相關
        self.dragged_window_title = {"title": None}
        self.drag_label_popup = {"win": None}
        
        # 建立 UI
        self._create_ui()
        self._setup_callbacks()
        
        # 載入設定
        self.load_settings()
        
        # 註冊熱鍵
        self._register_hotkeys()
        
        # 設定日誌回調
        self.logger.set_callback(self._on_log)
        
        # 關閉時儲存設定
        atexit.register(self.save_settings)
        self.app.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _create_ui(self):
        """建立使用者介面"""
        # 主框架
        frm = tb.Frame(self.app, padding=2)
        frm.pack(fill="both", expand=True)
        self.main_frame = frm
        
        # Row 0: 頂部工具列
        self._create_top_toolbar(frm)
        
        # Row 1: 分組置頂切換區
        self._create_group_switcher(frm)
        
        # Row 2: 分組檔案列
        self._create_group_panels(frm)
        
        # Row 8-9: 啟動/關閉按鈕
        self._create_control_buttons(frm)
        
        # Row 8-10: 日誌區
        self._create_log_panel(frm)
        
        # Row 10: 檔案/視窗列表
        self._create_file_window_lists(frm)
    
    def _create_top_toolbar(self, parent):
        """創建頂部工具列"""
        toolbar = tb.Frame(parent, padding=2)
        toolbar.grid(row=0, column=0, columnspan=8, sticky="ew", pady=(2, 2))
        
        # 資料夾選擇
        folder_frame = tb.Frame(toolbar, padding=(2,2))
        folder_frame.grid(row=0, column=0, sticky="w", padx=(0, 4))
        tb.Entry(folder_frame, textvariable=self.folder_var, width=25).grid(row=0, column=0, padx=(2,2))
        tb.Button(folder_frame, text="選擇路徑", command=self._choose_folder, bootstyle=SECONDARY).grid(row=0, column=1, padx=(2,0))
        
        # 間隔秒數
        interval_frame = tb.Frame(toolbar, padding=(2,2))
        interval_frame.grid(row=0, column=1, sticky="w", padx=(0, 4))
        tb.Label(interval_frame, text="間隔秒數:").grid(row=0, column=0)
        tb.Entry(interval_frame, textvariable=self.interval_var, width=3).grid(row=0, column=1, padx=(2,0))
        
        # 分組名稱編輯
        self.group_name_combo_var = tk.StringVar(value=self.group_codes[0])
        group_combo = tb.Combobox(toolbar, values=self.group_codes, textvariable=self.group_name_combo_var, width=3, state="readonly")
        group_combo.grid(row=0, column=3, padx=(8,2))
        
        self.group_name_entry = tb.Entry(toolbar, width=12)
        self.group_name_entry.grid(row=0, column=4, padx=(2,2))
        self.group_name_entry.insert(0, "修改分組名稱")
        self.group_name_entry.config(foreground="#888")
        
        # 存檔按鈕
        tb.Button(toolbar, text="存檔", command=self._manual_save, bootstyle="info", width=5).grid(row=0, column=5, padx=(8,2))
        
        # 關於按鈕
        tb.Button(toolbar, text="關於", command=self._show_about, bootstyle=SECONDARY, width=6).grid(row=0, column=9, padx=(8,2))
        
        # 刷新視窗按鈕
        tb.Button(toolbar, text="刷新視窗", command=self._refresh_windows, bootstyle=SECONDARY, width=8).grid(row=0, column=10, padx=(2,2))
    
    def _create_group_switcher(self, parent):
        """創建分組切換區"""
        switcher = tb.Frame(parent)
        switcher.grid(row=1, column=0, columnspan=8, sticky="ew", pady=(2,2))
        for i in range(7):
            switcher.grid_columnconfigure(i, weight=1)
        
        font = tkfont.Font(family="Microsoft JhengHei", size=12)
        tb.Label(switcher, text="置頂切換", width=12, anchor="center", font=font).grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        for idx, code in enumerate(self.group_codes):
            frame = tb.Frame(switcher, borderwidth=2, relief="groove")
            frame.grid(row=0, column=idx+1, padx=2, pady=2, sticky="ew")
            
            label = tb.Label(frame, text=f"{self.group_display_names[code].get()} ", width=6, font=font, cursor="hand2")
            label.pack(side="left")
            label.bind("<Button-1>", lambda e, c=code: self._focus_group(c))
            
            hotkey_entry = tb.Entry(frame, textvariable=self.group_hotkeys[idx], width=8, state="readonly", justify="center", font=font)
            hotkey_entry.pack(side="left", padx=(2,5))
            
            self.show_label_frames.append((label, hotkey_entry))
    
    def _create_group_panels(self, parent):
        """創建分組檔案面板"""
        for col in range(3):
            group_frame = tb.Frame(parent, borderwidth=1, relief="solid", padding=2)
            group_frame.grid(row=2, column=col, padx=2, pady=2, sticky="nsew")
            parent.grid_columnconfigure(col, weight=1)
            group_frame.grid_columnconfigure(1, weight=2)
            
            for i in range(5):  # 每欄5行
                idx = col * 5 + i
                if idx >= 15:
                    break
                
                # 序號按鈕
                num_btn = tb.Button(group_frame, text=str(idx+1), width=2, bootstyle="secondary",
                                   command=lambda i=idx: self._open_entry_file(i))
                num_btn.grid(row=i, column=0, sticky="w", padx=0)
                
                # 檔案名稱輸入框
                entry = tb.Entry(group_frame, state="readonly", width=10)
                entry.grid(row=i, column=1, padx=0, pady=1, sticky="ew")
                entry.bind("<Button-3>", lambda e, ent=entry: self._clear_entry(ent))
                
                # 分組下拉選單
                combos = []
                for j in range(4):
                    var = tk.StringVar(value="")
                    combo = tb.Combobox(group_frame, textvariable=var, width=3, state="readonly",
                                       values=[""] + [self.group_display_names[c].get() for c in self.group_codes])
                    combo.grid(row=i, column=2+j, padx=0, pady=1)
                    combos.append((var, combo))
                
                self.checkbox_vars_entries.append((entry, *[c[0] for c in combos], *[c[1] for c in combos]))
    
    def _create_control_buttons(self, parent):
        """創建控制按鈕區"""
        btns_frame = tb.Frame(parent)
        btns_frame.grid(row=8, column=1, rowspan=2, columnspan=6, sticky="ew", padx=(0, 4), pady=(8, 4))
        for i in range(6):
            btns_frame.grid_columnconfigure(i, weight=1)
        
        # 啟動按鈕
        for idx, code in enumerate(self.group_codes):
            col = idx
            # 啟動按鈕
            btn = tb.Button(btns_frame, text=f"啟動 {self.group_display_names[code].get()}", bootstyle="success-outline",
                           command=lambda c=code: self._start_group(c), width=8)
            btn.grid(row=0, column=col, padx=(4, 4), pady=(2, 2), sticky="nsew")
            self.group_buttons[code] = btn
            
            # 關閉按鈕
            btn = tb.Button(btns_frame, text=f"關閉 {self.group_display_names[code].get()}", bootstyle="danger-outline",
                           command=lambda c=code: self._close_group(c), width=8)
            btn.grid(row=1, column=col, padx=(4, 4), pady=(2, 2), sticky="nsew")
            self.close_buttons[code] = btn
    
    def _create_log_panel(self, parent):
        """創建日誌面板"""
        self.log_text = tb.Text(parent, height=20, width=18, state="disabled", wrap="word",
                               font=tkfont.Font(family="Microsoft JhengHei", size=10))
        self.log_text.grid(row=8, column=0, rowspan=3, sticky="nsew", padx=(0, 8), pady=(0, 0))
        parent.grid_rowconfigure(8, weight=1)
        parent.grid_rowconfigure(9, weight=1)
        parent.grid_rowconfigure(10, weight=1)
    
    def _create_file_window_lists(self, parent):
        """創建檔案與視窗列表"""
        bottom_frame = tb.Frame(parent)
        bottom_frame.grid(row=10, column=1, columnspan=2, sticky="ew", pady=(8, 2))
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_columnconfigure(1, weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)
        
        # 檔案列表
        self.file_list = FileListPanel(bottom_frame, on_drag_start=self._on_drag_start)
        self.file_list.grid(row=0, column=0, sticky="nsew")
        
        # 視窗列表
        self.window_list = WindowListPanel(bottom_frame, on_drag_start=self._on_drag_start)
        self.window_list.grid(row=0, column=1, sticky="nsew")
    
    def _setup_callbacks(self):
        """設定回調函數"""
        # 監聽變數變化
        self.folder_var.trace_add("write", lambda *a: self.save_settings())
        self.interval_var.trace_add("write", lambda *a: self.save_settings())
        
        for code in self.group_codes:
            self.group_display_names[code].trace_add("write", self._on_group_name_change)
        
        for var in self.group_hotkeys:
            var.trace_add("write", lambda *a: self._register_hotkeys())
    
    # ===== 事件處理 =====
    
    def _choose_folder(self):
        """選擇資料夾"""
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)
            self._update_file_list()
    
    def _manual_save(self):
        """手動儲存"""
        self.save_settings()
        self.logger.log("已手動儲存設定檔")
    
    def _focus_group(self, code: str):
        """聚焦分組"""
        files = self._get_group_files(code)
        if not files:
            self.logger.log(f"分組 {self.group_display_names[code].get()} 沒有檔案")
            return
        
        file_titles = [os.path.splitext(f)[0] for f in files]
        self.window_manager.update_group_windows(code, file_titles, self.app.winfo_id())
        title = self.window_manager.focus_next_in_group(code)
        
        if title:
            self.logger.log(f"切換到分組 {self.group_display_names[code].get()} 的視窗：{title}")
        else:
            self.logger.log(f"分組 {self.group_display_names[code].get()} 沒有視窗")
    
    def _start_group(self, code: str):
        """啟動分組"""
        folder = self.folder_var.get()
        try:
            interval = float(self.interval_var.get())
        except ValueError:
            self.logger.log("請輸入正確的間隔秒數")
            return
        
        if not os.path.isdir(folder):
            self.logger.log("請選擇正確的資料夾")
            return
        
        files = self._get_group_files(code)
        if not files:
            self.logger.log(f"分組 {self.group_display_names[code].get()} 沒有檔案")
            return
        
        self.logger.log(f"開始開啟分組 {self.group_display_names[code].get()} 的 {len(files)} 個檔案")
        
        def open_files():
            self.file_opener.open_files_with_interval(folder, files, interval)
        
        threading.Thread(target=open_files, daemon=True).start()
    
    def _close_group(self, code: str):
        """關閉分組"""
        files = self._get_group_files(code)
        if not files:
            self.logger.log(f"分組 {self.group_display_names[code].get()} 沒有檔案")
            return
        
        count = self.window_manager.close_windows_by_titles(files)
        
        if count > 0:
            self.logger.log(f"已關閉分組 {self.group_display_names[code].get()} 的 {count} 個視窗")
        else:
            self.logger.log(f"找不到分組 {self.group_display_names[code].get()} 的視窗可關閉")
    
    def _open_entry_file(self, idx: int):
        """開啟欄位中的檔案"""
        if idx >= len(self.checkbox_vars_entries):
            return
        
        entry = self.checkbox_vars_entries[idx][0]
        filename = entry.get().strip()
        
        if not filename:
            self.logger.log("此欄位無檔案名稱")
            return
        
        folder = self.folder_var.get()
        filepath = os.path.join(folder, filename)
        self.file_opener.open_file(filepath)
    
    def _clear_entry(self, entry):
        """清空欄位"""
        old = entry.get()
        entry.config(state="normal")
        entry.delete(0, tk.END)
        entry.config(state="readonly")
        if old:
            self.logger.log(f"清空分組欄位（原：{old}）")
    
    def _update_file_list(self):
        """更新檔案列表"""
        folder = self.folder_var.get()
        if not os.path.isdir(folder):
            self.file_list.update([])
            return
        
        files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
        self.file_list.update(files)
    
    def _refresh_windows(self):
        """刷新視窗列表"""
        titles = self.window_manager.get_taskbar_window_titles()
        self.window_list.update(titles)
        self.logger.log(f"已刷新視窗列表（共 {len(titles)} 個視窗）")
    
    def _on_drag_start(self, event, title: str):
        """開始拖曳"""
        self.dragged_window_title["title"] = title
        
        # 建立浮動標籤
        if self.drag_label_popup["win"]:
            self.drag_label_popup["win"].destroy()
        
        popup = tk.Toplevel(self.app)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        label = tb.Label(popup, text=title, background="#222", foreground="#fff")
        label.pack()
        self.drag_label_popup["win"] = popup
        
        def follow_mouse(ev):
            popup.geometry(f"+{ev.x_root + 10}+{ev.y_root + 10}")
        
        def on_drop(ev):
            # 檢查是否拖到分組欄位上
            for entry, *_ in self.checkbox_vars_entries:
                x1, y1 = entry.winfo_rootx(), entry.winfo_rooty()
                x2, y2 = x1 + entry.winfo_width(), y1 + entry.winfo_height()
                if x1 <= ev.x_root <= x2 and y1 <= ev.y_root <= y2:
                    entry.config(state="normal")
                    entry.delete(0, tk.END)
                    entry.insert(0, title)
                    entry.config(state="readonly")
                    self.logger.log(f"拖移「{title}」到分組欄位")
            
            if self.drag_label_popup["win"]:
                self.drag_label_popup["win"].destroy()
                self.drag_label_popup["win"] = None
            
            self.app.unbind("<Motion>")
            self.app.unbind("<ButtonRelease-1>")
        
        self.app.bind("<Motion>", follow_mouse)
        self.app.bind("<ButtonRelease-1>", on_drop)
    
    def _on_group_name_change(self, *args):
        """分組名稱變更"""
        # 更新顯示標籤
        for idx, code in enumerate(self.group_codes):
            self.show_label_frames[idx][0].config(text=f"{self.group_display_names[code].get()} ")
        
        # 更新按鈕文字
        for code in self.group_codes:
            if code in self.group_buttons:
                self.group_buttons[code].config(text=f"啟動 {self.group_display_names[code].get()}")
            if code in self.close_buttons:
                self.close_buttons[code].config(text=f"關閉 {self.group_display_names[code].get()}")
        
        # 更新下拉選單
        new_values = [""] + [self.group_display_names[c].get() for c in self.group_codes]
        for _, var1, var2, var3, var4, combo1, combo2, combo3, combo4 in self.checkbox_vars_entries:
            combo1.config(values=new_values)
            combo2.config(values=new_values)
            combo3.config(values=new_values)
            combo4.config(values=new_values)
        
        self.save_settings()
    
    def _on_log(self, msg: str):
        """日誌回調"""
        try:
            if self.log_text.winfo_exists():
                self.log_text.config(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.config(state="disabled")
        except Exception:
            pass
    
    def _show_about(self):
        """顯示關於對話框"""
        about_win = tb.Toplevel(self.app)
        about_win.title("關於 ChroLens_Portal")
        about_win.geometry("450x300")
        about_win.resizable(False, False)
        about_win.grab_set()
        
        # 置中
        x = self.app.winfo_x() + (self.app.winfo_width() // 2) - 225
        y = self.app.winfo_y() + 80
        about_win.geometry(f"+{x}+{y}")
        
        frm = tb.Frame(about_win, padding=20)
        frm.pack(fill="both", expand=True)
        
        tb.Label(frm, text="ChroLens_Portal 2.3\n分組開啟/關閉程式\n分組視窗置頂顯示",
                font=("Microsoft JhengHei", 11)).pack(anchor="w", pady=(0, 10))
        
        link = tk.Label(frm, text="ChroLens 討論區", font=("Microsoft JhengHei", 10, "underline"),
                       fg="#5865F2", cursor="hand2")
        link.pack(anchor="w")
        link.bind("<Button-1>", lambda e: os.startfile("https://discord.gg/72Kbs4WPPn"))
        
        tb.Label(frm, text="Created By Lucienwooo\n模組化版本 2.3",
                font=("Microsoft JhengHei", 10)).pack(anchor="w", pady=(10, 0))
        
        tb.Button(frm, text="關閉", command=about_win.destroy, width=8, bootstyle=SECONDARY).pack(anchor="e", pady=(16, 0))
    
    # ===== 輔助方法 =====
    
    def _get_group_files(self, code: str) -> list:
        """取得分組的檔案列表"""
        files = []
        for entry, var1, var2, var3, var4, *_ in self.checkbox_vars_entries:
            if (var1.get() == self.group_display_names[code].get() or
                var2.get() == self.group_display_names[code].get() or
                var3.get() == self.group_display_names[code].get() or
                var4.get() == self.group_display_names[code].get()):
                val = entry.get().strip()
                if val:
                    files.append(val)
        return files
    
    def _register_hotkeys(self):
        """註冊全域熱鍵"""
        hotkey_map = {}
        for idx, code in enumerate(self.group_codes):
            hotkey = self.group_hotkeys[idx].get()
            callback = functools.partial(self._focus_group, code)
            hotkey_map[code] = (hotkey, callback)
        
        self.hotkey_handler.update_hotkeys(hotkey_map)
    
    def save_settings(self):
        """儲存設定"""
        data = {
            "folder": self.folder_var.get(),
            "interval": self.interval_var.get(),
            "group_display_names": {c: self.group_display_names[c].get() for c in self.group_codes},
            "group_hotkeys": [v.get() for v in self.group_hotkeys],
            "checkbox_entries": [entry.get() for entry, *_ in self.checkbox_vars_entries],
            "group_var1": [var1.get() for _, var1, *_ in self.checkbox_vars_entries],
            "group_var2": [var2.get() for _, _, var2, *_ in self.checkbox_vars_entries],
            "group_var3": [var3.get() for _, _, _, var3, *_ in self.checkbox_vars_entries],
            "group_var4": [var4.get() for _, _, _, _, var4, *_ in self.checkbox_vars_entries],
        }
        self.config.save(data)
    
    def load_settings(self):
        """載入設定"""
        data = self.config.load()
        if not data:
            return
        
        self.folder_var.set(data.get("folder", ""))
        self.interval_var.set(data.get("interval", "4"))
        
        for c in self.group_codes:
            self.group_display_names[c].set(data.get("group_display_names", {}).get(c, c))
        
        for i, v in enumerate(data.get("group_hotkeys", [])):
            if i < len(self.group_hotkeys):
                self.group_hotkeys[i].set(v)
        
        entries = data.get("checkbox_entries", [])
        for i, entry in enumerate(entries):
            if i < len(self.checkbox_vars_entries):
                ent = self.checkbox_vars_entries[i][0]
                ent.config(state="normal")
                ent.delete(0, tk.END)
                ent.insert(0, entry)
                ent.config(state="readonly")
        
        for idx, key in enumerate(["group_var1", "group_var2", "group_var3", "group_var4"]):
            vars_list = data.get(key, [])
            for i, v in enumerate(vars_list):
                if i < len(self.checkbox_vars_entries):
                    self.checkbox_vars_entries[i][1+idx].set(v)
        
        # 更新檔案列表
        self._update_file_list()
        
        # 延遲刷新視窗列表
        self.app.after(500, self._refresh_windows)
    
    def _on_closing(self):
        """視窗關閉時"""
        self.save_settings()
        self.hotkey_handler.unregister_all()
        self.app.destroy()
    
    def run(self):
        """執行主程式"""
        self.logger.log("ChroLens_Portal 2.3 已啟動（模組化版，含管理員權限）")
        self.app.mainloop()
