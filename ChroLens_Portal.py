### ChroLens_Portal 2.4
### 2025/11/19 By Lucienwooo
### 視窗管理工具 - 分組啟動、快捷切換、智能布局記憶
### 需要管理者權限（會自動要求提升）
CURRENT_VERSION = "2.4"
import os
import time
import win32gui
import win32con
import threading
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, END, messagebox
import tkinter as tk
import pythoncom
from win32com.shell import shell, shellcon
import subprocess
import sys
import json
import keyboard
from tkinter import font as tkfont
import functools
import atexit
import win32process
from update_manager import UpdateManager
from update_dialog import UpdateDialog, NoUpdateDialog

SETTINGS_FILE = "chrolens_portal.json"

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# === 檢查並要求管理者權限 ===
def is_admin():
    """檢查是否以管理者身份運行"""
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """以管理者身份重新啟動程式"""
    try:
        if sys.argv[0].endswith('.py'):
            # 如果是 .py 檔案
            script = os.path.abspath(sys.argv[0])
            params = ' '.join([script] + sys.argv[1:])
            shell.ShellExecuteEx(
                lpVerb='runas',
                lpFile=sys.executable,
                lpParameters=params,
                nShow=1
            )
        else:
            # 如果是 .exe 檔案
            shell.ShellExecuteEx(
                lpVerb='runas',
                lpFile=sys.executable,
                lpParameters=' '.join(sys.argv[1:]),
                nShow=1
            )
        sys.exit(0)
    except Exception as e:
        print(f"無法以管理者身份啟動: {e}")
        sys.exit(1)

# 檢查管理者權限，如果沒有則要求提升
if not is_admin():
    print("需要管理者權限，正在重新啟動...")
    run_as_admin()

# === 介面區塊 ===
app = tb.Window(themename="darkly")
app.title(f"ChroLens_Portal {CURRENT_VERSION}")
try:
    ico_path = resource_path("冥想貓貓.ico")
    app.iconbitmap(ico_path)
except Exception as e:
    print(f"無法設定 icon: {e}")

# 在這裡建立字型物件
num_font = tkfont.Font(family="Microsoft JhengHei", size=10, weight="bold")

# 主視窗寬度（優化為更緊湊的大小）
app.geometry("1200x720")  # 減少高度

# --- 主 Frame ---
frm = tb.Frame(app, padding=2)
frm.pack(fill="both", expand=True)

# 設定響應式行配置
frm.grid_rowconfigure(0, weight=0)  # 頂部工具列 - 固定
frm.grid_rowconfigure(1, weight=0)  # 置頂切換區 - 固定
frm.grid_rowconfigure(2, weight=1)  # 分組檔案列 - 可擴展
frm.grid_rowconfigure(8, weight=0)  # 按鈕區 - 固定
frm.grid_rowconfigure(9, weight=0)  # 按鈕區 - 固定
frm.grid_rowconfigure(10, weight=1)  # 檔案/視窗列表 - 可擴展

# --- 分組與快捷鍵 ---
group_codes = ["A", "B", "C", "D", "E", "F"]
group_display_names = {c: tk.StringVar(value=c) for c in group_codes}
# 正確的預設快捷鍵
default_hotkeys = ["Alt+1", "Alt+2", "Alt+3", "Alt+q", "Alt+w", "Alt+e"]
group_hotkeys = [tk.StringVar(value=default_hotkeys[i]) for i in range(6)]
group_buttons = {}
close_buttons = {}

# --- Mini 模式狀態追踪（用於資源優化）---
mini_mode_active = False

# --- 排程任務（需要在程式開頭定義以避免 save_settings 錯誤）---
schedule_tasks = []

# --- 視窗位置和大小記憶 (FancyZones 功能) ---
# 儲存格式: { "group_code": { "file_name": {"x": int, "y": int, "width": int, "height": int, "state": str} } }
window_layouts = {}

# --- row 0：頂部工具列 ---
top_row_frame = tb.Frame(frm, padding=2)
top_row_frame.grid(row=0, column=0, columnspan=8, sticky="ew", pady=(2, 2))
top_row_frame.grid_columnconfigure(0, weight=0)
top_row_frame.grid_columnconfigure(1, weight=0)
top_row_frame.grid_columnconfigure(2, weight=0)

folder_var = tb.StringVar(value="")
interval_var = tb.StringVar(value="4")

def choose_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_var.set(folder_selected)
        if not mini_mode_active:
            update_file_list()  # 只在非 mini 模式下刷新檔案列表

folder_frame = tb.Frame(top_row_frame, padding=(2,2))
folder_frame.grid(row=0, column=0, sticky="w", padx=(0, 4))
tb.Entry(folder_frame, textvariable=folder_var, width=25).grid(row=0, column=0, padx=(2,2), sticky="ew")
tb.Button(folder_frame, text="選擇開啟路徑", command=lambda: choose_folder(), bootstyle=SECONDARY).grid(row=0, column=1, padx=(2,0), sticky="ew")

interval_frame = tb.Frame(top_row_frame, padding=(2,2))
interval_frame.grid(row=0, column=1, sticky="w", padx=(0, 4))
tb.Label(interval_frame, text="間隔秒數:").grid(row=0, column=0, sticky="w")
tb.Entry(interval_frame, textvariable=interval_var, width=3).grid(row=0, column=1, padx=(2,0), sticky="w")

# === 視窗佈局記憶功能 (FancyZones) ===
def capture_window_layout(group_code):
    """捕獲指定分組所有視窗的位置和大小"""
    files = get_group_files(group_code)
    if not files:
        log(f"分組 {group_display_names[group_code].get()} 沒有檔案，無法捕獲佈局")
        return
    
    # 取得分組視窗標題關鍵字
    target_titles = {}
    for f in files:
        if f:
            file_key = os.path.splitext(os.path.basename(f))[0].lower()
            target_titles[file_key] = f
    
    my_hwnd = app.winfo_id()
    captured = {}
    
    def enum_handler(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        if hwnd == my_hwnd:
            return
        
        window_text = win32gui.GetWindowText(hwnd)
        window_text_lower = window_text.lower().strip()
        if not window_text_lower:
            return
        
        # 找到匹配的檔案
        for file_key, filename in target_titles.items():
            if file_key in window_text_lower:
                try:
                    # 獲取視窗位置和大小
                    rect = win32gui.GetWindowRect(hwnd)
                    placement = win32gui.GetWindowPlacement(hwnd)
                    
                    # 儲存視窗資訊
                    captured[filename] = {
                        "x": rect[0],
                        "y": rect[1],
                        "width": rect[2] - rect[0],
                        "height": rect[3] - rect[1],
                        "state": placement[1]  # SW_SHOWNORMAL=1, SW_SHOWMINIMIZED=2, SW_SHOWMAXIMIZED=3
                    }
                    log(f"捕獲視窗佈局：{filename} ({rect[2]-rect[0]}x{rect[3]-rect[1]} at {rect[0]},{rect[1]})")
                except Exception as e:
                    log(f"捕獲視窗佈局失敗：{window_text} ({e})")
                break
    
    win32gui.EnumWindows(enum_handler, None)
    
    if captured:
        if group_code not in window_layouts:
            window_layouts[group_code] = {}
        window_layouts[group_code] = captured
        save_settings()
        log(f"已捕獲分組 {group_display_names[group_code].get()} 的 {len(captured)} 個視窗佈局")
    else:
        log(f"未找到分組 {group_display_names[group_code].get()} 的任何視窗")

def restore_window_layout(group_code, hwnd, window_text):
    """恢復指定視窗的位置和大小"""
    if group_code not in window_layouts:
        log(f"[佈局] 分組 {group_code} 沒有保存的佈局資料")
        return False
    
    layout = window_layouts[group_code]
    if not layout:
        log(f"[佈局] 分組 {group_code} 的佈局資料為空")
        return False
    
    window_text_lower = window_text.lower()
    
    # 找到匹配的佈局
    for filename, pos_data in layout.items():
        file_key = os.path.splitext(os.path.basename(filename))[0].lower()
        if file_key in window_text_lower:
            try:
                x = pos_data["x"]
                y = pos_data["y"]
                width = pos_data["width"]
                height = pos_data["height"]
                state = pos_data.get("state", win32con.SW_SHOWNORMAL)
                
                # 恢復視窗狀態和位置
                if state == win32con.SW_SHOWMAXIMIZED:
                    # 最大化視窗
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)
                    log(f"[佈局] 恢復最大化：{window_text}")
                elif state == win32con.SW_SHOWMINIMIZED:
                    # 最小化視窗
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)
                    log(f"[佈局] 恢復最小化：{window_text}")
                else:
                    # 正常視窗：直接設置位置和大小，不調用 ShowWindow
                    # 使用 SWP_NOZORDER 避免改變 Z-order
                    win32gui.SetWindowPos(
                        hwnd, 
                        win32con.HWND_TOP,
                        x, y, width, height,
                        win32con.SWP_SHOWWINDOW | win32con.SWP_NOZORDER
                    )
                    log(f"[佈局] 恢復位置：{window_text} -> {width}x{height} at ({x},{y})")
                return True
            except Exception as e:
                log(f"[佈局] 恢復失敗：{window_text} ({e})")
                return False
    
    log(f"[佈局] 未找到匹配的佈局：{window_text}")
    return False

# 新增「存檔」按鈕（捕獲當前所有分組視窗佈局）
def manual_save():
    log("=" * 50)
    log("開始捕獲所有分組的視窗佈局...")
    total_captured = 0
    
    # 捕獲所有當前活躍分組的視窗佈局
    for group_code in group_codes:
        files = get_group_files(group_code)
        if files:
            log(f"分組 {group_display_names[group_code].get()} 包含檔案: {files}")
            before_count = len(window_layouts.get(group_code, {}))
            capture_window_layout(group_code)
            after_count = len(window_layouts.get(group_code, {}))
            captured = after_count - before_count
            if captured > 0:
                total_captured += captured
        else:
            log(f"分組 {group_display_names[group_code].get()} 沒有檔案，跳過")
    
    save_settings()
    log(f"已手動儲存設定檔，共捕獲 {total_captured} 個新視窗佈局")
    log("=" * 50)

save_btn = tb.Button(top_row_frame, text="存檔", command=manual_save, bootstyle="info", width=5)
save_btn.grid(row=0, column=5, padx=(8,2), sticky="e")

# --- 新增：分組名稱修改區 ---
group_name_edit_var = tk.StringVar()
group_name_edit_combo_var = tk.StringVar(value=group_codes[0])
group_name_placeholder = "修改分組名稱"
group_name_edit_entry = None  # 先宣告

def get_default_group_name(code):
    return code  # 預設就是 A、B、C、D、E、F

def show_placeholder():
    if group_name_edit_entry.get() == "":
        group_name_edit_entry.insert(0, group_name_placeholder)
        group_name_edit_entry.config(foreground="#888")
        group_name_edit_entry.placeholder = True

def hide_placeholder():
    if getattr(group_name_edit_entry, "placeholder", False):
        group_name_edit_entry.delete(0, tk.END)
        group_name_edit_entry.config(foreground="#fff")
        group_name_edit_entry.placeholder = False

def on_group_name_combo_change(event=None):
    code = group_name_edit_combo_var.get()
    val = group_display_names[code].get()
    group_name_edit_entry.delete(0, tk.END)
    if not val or val == get_default_group_name(code):
        show_placeholder()
    else:
        group_name_edit_entry.insert(0, val)
        hide_placeholder()

def on_group_name_edit_submit(event=None):
    code = group_name_edit_combo_var.get()
    entry_val = group_name_edit_entry.get().strip()
    # 若是 placeholder 或空白，恢復預設
    if not entry_val or entry_val == group_name_placeholder:
        group_display_names[code].set(get_default_group_name(code))
        group_name_edit_entry.delete(0, tk.END)
        show_placeholder()
    else:
        group_display_names[code].set(entry_val)
        hide_placeholder()

group_name_edit_combo = tb.Combobox(
    top_row_frame, values=group_codes, textvariable=group_name_edit_combo_var, width=3, state="readonly"
)
group_name_edit_entry = tb.Entry(
    top_row_frame, textvariable=group_name_edit_var, width=12
)
group_name_edit_entry.placeholder = False

group_name_edit_combo.grid(row=0, column=3, padx=(8,2), sticky="w")
group_name_edit_entry.grid(row=0, column=4, padx=(2,2), sticky="w")
save_btn.grid(row=0, column=5, padx=(8,2), sticky="e")

# 綁定事件
group_name_edit_combo.bind("<<ComboboxSelected>>", on_group_name_combo_change)
group_name_edit_entry.bind("<FocusIn>", lambda e: (hide_placeholder(), group_name_edit_entry.config(foreground="#fff")))
group_name_edit_entry.bind("<FocusOut>", lambda e: (show_placeholder() if not group_name_edit_entry.get().strip() else None, on_group_name_edit_submit()))
group_name_edit_entry.bind("<Return>", on_group_name_edit_submit)

# 初始化顯示 placeholder
show_placeholder()

# --- row 1：分組置頂顯示切換區 ---
show_label_frames = []
second_row_frame = tb.Frame(frm)
second_row_frame.grid(row=1, column=0, columnspan=8, sticky="ew")
for i in range(7):
    second_row_frame.grid_columnconfigure(i, weight=1)
show_label_font = tkfont.Font(family="Microsoft JhengHei", size=12)
desc_label = tb.Label(second_row_frame, text="置頂切換", width=12, anchor="center", font=show_label_font)
desc_label.grid(row=0, column=0, padx=2, pady=2, sticky="ew")

# 建立 mini 模式的還原按鈕（初始隱藏）
# 使用 Frame + Label 來顯示較大的 emoji 箭頭
mini_restore_frame = tb.Frame(second_row_frame)
mini_restore_font = tkfont.Font(family="Segoe UI Emoji", size=16, weight="bold")
mini_restore_label = tb.Label(
    mini_restore_frame,
    text="↩️",
    font=mini_restore_font,
    cursor="hand2",
    anchor="center"
)
mini_restore_label.pack()
# 暫時的佔位命令，稍後會被更新
mini_restore_label.bind("<Button-1>", lambda e: None)
# 初始時不顯示這個按鈕
# mini_restore_btn.grid(row=0, column=0, padx=2, pady=2, sticky="ew")

for idx, code in enumerate(group_codes):
    frame = tb.Frame(second_row_frame, borderwidth=2, relief="groove")
    frame.grid(row=0, column=idx+1, padx=2, pady=2, sticky="ew")
    # 將 Label 改為可點擊的按鈕樣式，但保持 Label 外觀
    show_label = tb.Label(frame, text=f"{group_display_names[code].get()} ", width=6, font=show_label_font, cursor="hand2")
    show_label.pack(side="left")
    # 綁定點擊事件，觸發對應分組的快捷鍵功能
    show_label.bind("<Button-1>", lambda e, c=code: focus_next_in_group(c))
    hotkey_entry = tb.Entry(frame, textvariable=group_hotkeys[idx], width=8, state="readonly", justify="center", font=show_label_font)
    hotkey_entry.pack(side="left", padx=(2,5))
    def make_on_key(idx):
        return lambda event, i=idx: on_hotkey_entry_key(event, i)
    hotkey_entry.bind("<Key>", make_on_key(idx))
    hotkey_entry.bind("<Button-1>", lambda e, entry=hotkey_entry: entry.focus_set())
    show_label_frames.append((show_label, hotkey_entry))

def update_show_labels(*args):
    for idx, code in enumerate(group_codes):
        show_label_frames[idx][0].config(text=f"{group_display_names[code].get()} ")  
for c in group_codes:
    group_display_names[c].trace_add("write", update_show_labels)

# --- 熱鍵輸入框的按鍵處理函數 ---
def on_hotkey_entry_key(event, idx):
    """處理熱鍵輸入框的按鍵事件"""
    # 組合按鍵字串
    modifiers = []
    if event.state & 0x4:  # Control
        modifiers.append("Ctrl")
    if event.state & 0x1:  # Shift
        modifiers.append("Shift")
    if event.state & 0x20000:  # Alt
        modifiers.append("Alt")
    
    # 取得按鍵名稱
    key = event.keysym
    if key in ["Control_L", "Control_R", "Shift_L", "Shift_R", "Alt_L", "Alt_R"]:
        return  # 忽略單獨的修飾鍵
    
    # 組合完整的快捷鍵字串
    if modifiers and key:
        hotkey = "+".join(modifiers + [key])
        group_hotkeys[idx].set(hotkey)
        log(f"已設定分組 {group_codes[idx]} 的快捷鍵為：{hotkey}")

# --- 編號標籤點擊處理函數 ---
def on_num_label_click(event, entry):
    """點擊編號標籤時開啟對應的檔案"""
    open_entry_file(entry)

# --- row 2：分組檔案列 ---
group_frames = []
for col in range(3):
    group_frame = tb.Frame(frm, borderwidth=1, relief="solid", padding=1)
    group_frame.grid(row=2, column=col, padx=2, pady=1, sticky="nsew")
    frm.grid_columnconfigure(col, weight=1)
    group_frame.grid_columnconfigure(1, weight=2)  # 讓檔案名稱欄自動展開且多吃空間
    for i in range(2, 6):  # combo欄不自動展開
        group_frame.grid_columnconfigure(i, weight=0)
    group_frames.append(group_frame)

combo_width = 3  # 原本是4，縮短1/3

checkbox_vars_entries = []

for i in range(15):  # 15行
    row = i % 5
    col = i // 5
    entry = tb.Entry(group_frames[col], state="readonly", width=10)
    entry.grid(row=row, column=1, padx=0, pady=1, sticky="ew")  # sticky="ew" 讓欄位自動展開
    group_var1 = tk.StringVar(value="")
    group_var2 = tk.StringVar(value="")
    group_var3 = tk.StringVar(value="")
    group_var4 = tk.StringVar(value="")
    group_combo1 = tb.Combobox(
        group_frames[col], textvariable=group_var1,
        values=[""] + [group_display_names[c].get() for c in group_codes], width=combo_width, state="readonly"
    )
    group_combo2 = tb.Combobox(
        group_frames[col], textvariable=group_var2,
        values=[""] + [group_display_names[c].get() for c in group_codes], width=combo_width, state="readonly"
    )
    group_combo3 = tb.Combobox(
        group_frames[col], textvariable=group_var3,
        values=[""] + [group_display_names[c].get() for c in group_codes], width=combo_width, state="readonly"
    )
    group_combo4 = tb.Combobox(
        group_frames[col], textvariable=group_var4,
        values=[""] + [group_display_names[c].get() for c in group_codes], width=combo_width, state="readonly"
    )
    num_label = tb.Label(group_frames[col], text=str(i+1), width=2, font=("Microsoft JhengHei", 10, "bold"), background="#444", foreground="#fff", anchor="center", cursor="hand2")
    num_label.grid(row=row, column=0, sticky="w", padx=0)
    num_label.bind("<Button-1>", lambda e, ent=entry: on_num_label_click(e, ent))
    entry.grid(row=row, column=1, padx=0, pady=1, sticky="ew")
    group_combo1.grid(row=row, column=2, padx=0, pady=1)
    group_combo2.grid(row=row, column=3, padx=0, pady=1)
    group_combo3.grid(row=row, column=4, padx=0, pady=1)
    group_combo4.grid(row=row, column=5, padx=0, pady=1)
    checkbox_vars_entries.append((entry, group_var1, group_var2, group_var3, group_var4, group_combo1, group_combo2, group_combo3, group_combo4))
    num_btn = tb.Button(
        group_frames[col],
        text=str(i+1),
        width=2,
        bootstyle="secondary",
        style="Num.TButton",
        command=lambda ent=entry: open_entry_file(ent)
    )
    num_btn.grid(row=row, column=0, sticky="w", padx=0)

# 先在初始化時建立 style
style = tb.Style()
style.configure("Num.TButton", font=("Microsoft JhengHei", 10, "bold"))

# --- row 8~10 動態日誌區塊 ---
log_text = tb.Text(frm, height=12, width=18, state="disabled", wrap="word", font=tkfont.Font(family="Microsoft JhengHei", size=10))
log_text.grid(row=8, column=0, rowspan=3, sticky="nsew", padx=(0, 8), pady=(0, 0))

# --- row 8~9 啟動/關閉分組按鈕區域 ---
btns_outer_frame = tb.Frame(frm)
btns_outer_frame.grid(row=8, column=1, rowspan=2, columnspan=6, sticky="new", padx=(0, 4), pady=(0, 0))
for i in range(6):
    btns_outer_frame.grid_columnconfigure(i, weight=1)
group_btn_grid = [
    (8, 0, "啟動", "A", "success-outline", lambda: start_group_opening("A")),
    (8, 1, "啟動", "B", "success-outline", lambda: start_group_opening("B")),
    (8, 2, "啟動", "C", "success-outline", lambda: start_group_opening("C")),
    (8, 3, "關閉", "A", "danger-outline", lambda: close_group_windows("A")),
    (8, 4, "關閉", "B", "danger-outline", lambda: close_group_windows("B")),
    (8, 5, "關閉", "C", "danger-outline", lambda: close_group_windows("C")),
    (9, 0, "啟動", "D", "success-outline", lambda: start_group_opening("D")),
    (9, 1, "啟動", "E", "success-outline", lambda: start_group_opening("E")),
    (9, 2, "啟動", "F", "success-outline", lambda: start_group_opening("F")),
    (9, 3, "關閉", "D", "danger-outline", lambda: close_group_windows("D")),
    (9, 4, "關閉", "E", "danger-outline", lambda: close_group_windows("E")),
    (9, 5, "關閉", "F", "danger-outline", lambda: close_group_windows("F")),
]
for row, col, text, code, bootstyle, cmd in group_btn_grid:
    btn = tb.Button(
        btns_outer_frame,
        text=f"{group_display_names[code].get()}",
        bootstyle=bootstyle,
        command=cmd,
        width=8
    )
    btn.grid(row=row-8, column=col, padx=(2, 2), pady=(2, 2), sticky="ew")
    if text == "啟動":
        group_buttons[code] = btn
    else:
        close_buttons[code] = btn

# --- row 10 檔案名稱/視窗名稱列表 ---
bottom_frame = tb.Frame(frm)
bottom_frame.grid(row=10, column=1, columnspan=2, sticky="nsew", pady=(2, 2))
bottom_frame.grid_columnconfigure(0, weight=1)  # 兩欄等寬
bottom_frame.grid_columnconfigure(1, weight=1)
bottom_frame.grid_rowconfigure(0, weight=1)

# ===== 交換：左側顯示檔案名稱列表，右側顯示視窗名稱列表 =====

# 檔案名稱列表（左側，寬度自動展開）
file_list_outer = tb.Frame(bottom_frame)
file_list_outer.grid(row=0, column=0, sticky="nsew")
file_list_outer.grid_propagate(True)
file_list_outer.grid_rowconfigure(0, weight=1)
file_list_outer.grid_columnconfigure(0, weight=1)
file_list_canvas = tk.Canvas(file_list_outer, highlightthickness=0)
file_list_canvas.grid(row=0, column=0, sticky="nsew")
file_list_inner_frame = tb.Frame(file_list_canvas)
file_list_inner_frame_id = file_list_canvas.create_window((0, 0), window=file_list_inner_frame, anchor="nw")
file_list_inner_frame.grid_columnconfigure(0, weight=1)

file_list_vsb = tb.Scrollbar(file_list_outer, orient="vertical", command=file_list_canvas.yview)
file_list_canvas.configure(yscrollcommand=lambda *args: _on_file_vsb(*args))

def _on_file_vsb(*args):
    file_list_vsb.set(*args)
    if float(args[0]) <= 0.0 and float(args[1]) >= 1.0:
        file_list_vsb.grid_remove()
    else:
        file_list_vsb.grid(row=0, column=1, sticky="ns")

def _on_file_frame_configure(event):
    file_list_canvas.configure(scrollregion=file_list_canvas.bbox("all"))
    file_list_canvas.itemconfig(file_list_inner_frame_id, width=file_list_canvas.winfo_width())
file_list_inner_frame.bind("<Configure>", _on_file_frame_configure)

def _on_file_mousewheel(event):
    file_list_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
file_list_canvas.bind("<Enter>", lambda e: file_list_canvas.bind_all("<MouseWheel>", _on_file_mousewheel))
file_list_canvas.bind("<Leave>", lambda e: file_list_canvas.unbind_all("<MouseWheel>"))

def update_file_list():
    # Mini 模式下跳過更新以節省資源
    if mini_mode_active:
        return
    
    # 清空現有檔案列表
    for widget in file_list_inner_frame.winfo_children():
        widget.destroy()
    folder = folder_var.get()
    if not os.path.isdir(folder):
        return
    files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    for row, filename in enumerate(files):
        lbl = tb.Label(
            file_list_inner_frame,
            text=filename,
            anchor="w",
            font=tkfont.Font(family="Microsoft JhengHei", size=10)
        )
        lbl.grid(row=row, column=0, sticky="ew", padx=2, pady=1)
        # 修正：lambda 預設參數正確傳遞 filename
        lbl.bind("<ButtonPress-1>", lambda e, t=filename: on_label_drag_start(e, t))
    file_list_inner_frame.update_idletasks()
    file_list_canvas.config(scrollregion=file_list_canvas.bbox("all"))

# 視窗名稱列表（右側，寬度自動展開）
window_list_outer = tb.Frame(bottom_frame)
window_list_outer.grid(row=0, column=1, sticky="nsew")
window_list_outer.grid_propagate(True)
window_list_outer.grid_rowconfigure(0, weight=1)
window_list_outer.grid_columnconfigure(0, weight=1)
window_list_canvas = tk.Canvas(window_list_outer, highlightthickness=0)
window_list_canvas.grid(row=0, column=0, sticky="nsew")
window_list_inner_frame = tb.Frame(window_list_canvas)
window_list_inner_frame_id = window_list_canvas.create_window((0, 0), window=window_list_inner_frame, anchor="nw")
window_list_inner_frame.grid_columnconfigure(0, weight=1)

window_list_vsb = tb.Scrollbar(window_list_outer, orient="vertical", command=window_list_canvas.yview)  # ←提前到這裡

window_list_vsb.grid(row=0, column=1, sticky="ns")
def _on_window_vsb(*args):
    window_list_vsb.set(*args)
    # 自動隱藏/顯示
    if float(args[0]) <= 0.0 and float(args[1]) >= 1.0:
        window_list_vsb.grid_remove()
    else:
        window_list_vsb.grid(row=0, column=1, sticky="ns")
window_list_canvas.configure(yscrollcommand=_on_window_vsb)

def _on_window_frame_configure(event):
    window_list_canvas.configure(scrollregion=window_list_canvas.bbox("all"))
    window_list_canvas.itemconfig(window_list_inner_frame_id, width=window_list_canvas.winfo_width())
window_list_inner_frame.bind("<Configure>", _on_window_frame_configure)

def _on_window_mousewheel(event):
    window_list_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
window_list_canvas.bind("<Enter>", lambda e: window_list_canvas.bind_all("<MouseWheel>", _on_window_mousewheel))
window_list_canvas.bind("<Leave>", lambda e: window_list_canvas.unbind_all("<MouseWheel>"))

def update_window_list():
    # 清空現有視窗列表
    for widget in window_list_inner_frame.winfo_children():
        widget.destroy()
    titles = get_taskbar_window_titles()
    for row, title in enumerate(titles):
        lbl = tb.Label(
            window_list_inner_frame,
            text=title,
            anchor="w",
            font=tkfont.Font(family="Microsoft JhengHei", size=10)
        )
        lbl.grid(row=row, column=0, sticky="ew", padx=2, pady=1)
        lbl.bind("<ButtonPress-1>", lambda e, t=title: on_label_drag_start(e, t))
    window_list_inner_frame.update_idletasks()
    window_list_canvas.config(scrollregion=window_list_canvas.bbox("all"))

def get_taskbar_window_titles():
    # 捕捉所有可見視窗標題，排除系統/背景視窗
    exclude_keywords = [
        "設定", "windows 輸入體驗", "windows input experience", "searchui", "cortana", "開始功能表", "start menu",
        "工作管理員", "task manager", "lockapp", "shell experience host", "runtimebroker", "searchapp"
    ]
    titles = []
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and title.strip():
                title_lower = title.strip().lower()
                # 過濾系統視窗
                if not any(keyword in title_lower for keyword in exclude_keywords):
                    titles.append(title)
    win32gui.EnumWindows(enum_handler, None)
    return titles

def update_window_list():
    # Mini 模式下跳過更新以節省資源
    if mini_mode_active:
        return
    
    # 清空現有視窗列表
    for widget in window_list_inner_frame.winfo_children():
        widget.destroy()
    titles = get_taskbar_window_titles()
    for row, title in enumerate(titles):
        lbl = tb.Label(
            window_list_inner_frame,
            text=title,
            anchor="w",
            font=tkfont.Font(family="Microsoft JhengHei", size=10)
        )
        lbl.grid(row=row, column=0, sticky="ew", padx=2, pady=1)
        lbl.bind("<ButtonPress-1>", lambda e, t=title: on_label_drag_start(e, t))
    window_list_inner_frame.update_idletasks()
    window_list_canvas.config(scrollregion=window_list_canvas.bbox("all"))

# 啟動時呼叫
update_file_list()
update_window_list()


window_list_canvas.bind("<Enter>", lambda e: window_list_canvas.bind_all("<MouseWheel>", _on_window_mousewheel))
window_list_canvas.bind("<Leave>", lambda e: window_list_canvas.unbind_all("<MouseWheel>"))

# --- 拖移功能 ---
dragged_window_title = {"title": None}
drag_label_popup = {"win": None}

def on_label_drag_start(event, title):
    dragged_window_title["title"] = title
    # 建立浮動標籤
    if drag_label_popup["win"]:
        drag_label_popup["win"].destroy()
    drag_label_popup["win"] = tk.Toplevel(app)
    drag_label_popup["win"].overrideredirect(True)
    drag_label_popup["win"].attributes("-topmost", True)
    label = tb.Label(drag_label_popup["win"], text=title, background="#222", foreground="#fff", font=tkfont.Font(family="Microsoft JhengHei", size=10))
    label.pack()
    def follow_mouse(ev):
        x = ev.x_root + 10
        y = ev.y_root + 10
        drag_label_popup["win"].geometry(f"+{x}+{y}")
    app.bind("<Motion>", follow_mouse)
    def on_drop(ev):
        # 檢查滑鼠下方是否為分組欄位
        for entry, *_ in checkbox_vars_entries:
            x1 = entry.winfo_rootx()
            y1 = entry.winfo_rooty()
            x2 = x1 + entry.winfo_width()
            y2 = y1 + entry.winfo_height()
            if x1 <= ev.x_root <= x2 and y1 <= ev.y_root <= y2:
                entry.config(state="normal")
                entry.delete(0, tk.END)
                entry.insert(0, title)
                entry.config(state="readonly")
                log(f"拖移「{title}」到分組欄位")
        if drag_label_popup["win"]:
            drag_label_popup["win"].destroy()
            drag_label_popup["win"] = None
        app.unbind("<Motion>")
        app.unbind("<ButtonRelease-1>")
        dragged_window_title["title"] = None
    app.bind("<ButtonRelease-1>", on_drop)

# --- 讓 15 組分組框的檔案名稱框支援右鍵清空 ---
for entry, *_ in checkbox_vars_entries:
    def clear_entry(event, ent=entry):
        old = ent.get()
        ent.config(state="normal")
        ent.delete(0, tk.END)
        ent.config(state="readonly")
        if old:
            log(f"清空分組欄位內容（原內容：{old}）")
    entry.bind("<Button-3>", clear_entry)  # 右鍵點擊清空內容

def set_group_windows_topmost(group_code):
    """將指定分組的所有視窗設為最上層並恢復佈局（FancyZones 功能）"""
    files = get_group_files(group_code)
    if not files:
        return

    # 取得分組視窗標題關鍵字
    target_titles = [os.path.splitext(os.path.basename(f))[0].lower() for f in files if f]
    my_hwnd = app.winfo_id()
    group_hwnds = []

    # 找出分組視窗
    def enum_handler(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        if hwnd == my_hwnd:
            return
        window_text = win32gui.GetWindowText(hwnd)
        window_text_lower = window_text.lower().strip()
        if not window_text_lower:
            return
        if any(title and title in window_text_lower for title in target_titles):
            group_hwnds.append((hwnd, window_text))
    win32gui.EnumWindows(enum_handler, None)

    if not group_hwnds:
        return

    # 處理每個分組視窗
    for hwnd, window_text in group_hwnds:
        try:
            # 嘗試恢復佈局
            layout_restored = restore_window_layout(group_code, hwnd, window_text)
            
            if not layout_restored:
                # 如果沒有保存的佈局，只置頂視窗
                win32gui.SetWindowPos(
                    hwnd, win32con.HWND_TOP, 0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
                )
        except Exception as e:
            log(f"處理視窗失敗：{window_text} ({e})")

def start_group_opening(group_code):
    folder = folder_var.get()
    try:
        interval = float(interval_var.get())
    except ValueError:
        log("請輸入正確的間隔秒數")
        return
    if not os.path.isdir(folder):
        log("請選擇正確的資料夾")
        return
    files = get_group_files(group_code)
    if not files:
        log(f"分組 {group_display_names[group_code].get()} 沒有檔案")
        return
    log(f"開始開啟分組 {group_display_names[group_code].get()} 的檔案於 {folder}")
    
    def wait_for_window_and_restore(file_key, max_wait=10):
        """等待視窗出現並恢復佈局"""
        start_time = time.time()
        my_hwnd = app.winfo_id()
        
        while time.time() - start_time < max_wait:
            found = False
            def enum_handler(hwnd, _):
                nonlocal found
                if not win32gui.IsWindowVisible(hwnd):
                    return
                if hwnd == my_hwnd:
                    return
                
                window_text = win32gui.GetWindowText(hwnd)
                if file_key.lower() in window_text.lower():
                    # 找到視窗，嘗試恢復佈局
                    if restore_window_layout(group_code, hwnd, window_text):
                        found = True
                        return
                    else:
                        # 如果沒有保存的佈局，至少置頂視窗
                        try:
                            win32gui.SetWindowPos(
                                hwnd, win32con.HWND_TOP, 0, 0, 0, 0,
                                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
                            )
                        except:
                            pass
                        found = True
            
            win32gui.EnumWindows(enum_handler, None)
            if found:
                break
            time.sleep(0.3)
    
    def open_files():
        for file in files:
            file_path = os.path.join(folder, file)
            if not os.path.exists(file_path):
                log(f"找不到檔案: {file_path}")
                continue
            
            file_key = os.path.splitext(os.path.basename(file))[0]
            
            try:
                if file.lower().endswith('.lnk'):
                    target, args = open_lnk_target(file_path)
                    if target and os.path.exists(target):
                        log(f"Opening shortcut target: {target} {args}")
                        subprocess.Popen(
                            f'"{target}" {args}',
                            shell=True,
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                    else:
                        log(f"無法解析捷徑或目標不存在: {file_path}")
                        continue
                else:
                    log(f"Opening: {file_path}")
                    os.startfile(file_path)
                
                # 等待視窗出現並恢復佈局
                wait_for_window_and_restore(file_key)
                
            except Exception as e:
                log(f"無法開啟: {file_path}，錯誤：{e}")
            
            time.sleep(interval)
    
    threading.Thread(target=open_files, daemon=True).start()

def close_group_windows(group_code):
    files = get_group_files(group_code)
    if not files:
        log(f"分組 {group_display_names[group_code].get()} 沒有檔案")
        return

    keywords = []
    for filename in files:
        filename = filename.strip().lower()
        if filename:
            keywords.append(filename)
            if "." in filename:
                keywords.append(os.path.splitext(filename)[0])
    keywords = list(set(keywords))
    closed_any = False

    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            window_text_lower = window_text.lower()
            if any(kw and kw in window_text_lower for kw in keywords):
                try:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    log(f"已關閉視窗：{window_text}")
                    nonlocal closed_any
                    closed_any = True
                except Exception as e:
                    log(f"關閉視窗失敗：{window_text} ({e})")
    win32gui.EnumWindows(enum_handler, None)
    if not closed_any:
        log(f"找不到分組 {group_display_names[group_code].get()} 的視窗可關閉")

def save_settings():
    try:
        data = {
            "folder": folder_var.get(),
            "interval": interval_var.get(),
            "group_display_names": {c: group_display_names[c].get() for c in group_codes},
            "group_hotkeys": [v.get() for v in group_hotkeys],
            "checkbox_entries": [entry.get() for entry, *_ in checkbox_vars_entries],
            "group_var1": [var1.get() for _, var1, _, _, _, *_ in checkbox_vars_entries],
            "group_var2": [var2.get() for _, _, var2, _, _, *_ in checkbox_vars_entries],
            "group_var3": [var3.get() for _, _, _, var3, _, *_ in checkbox_vars_entries],
            "group_var4": [var4.get() for _, _, _, _, var4, *_ in checkbox_vars_entries],
            "schedule_tasks": schedule_tasks,  # 儲存排程任務
            "window_layouts": window_layouts,  # 儲存視窗佈局
        }
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except NameError:
        # 忽略變數未定義的錯誤（通常發生在程式初始化時）
        pass
    except Exception as e:
        # 只在非初始化錯誤時記錄
        if "is not defined" not in str(e):
            log(f"儲存設定檔失敗: {e}")

        
# 1. 先定義 log
log_history = []
def log(msg):
    timestamp = time.strftime("%H:%M:%S")
    full_msg = f"[{timestamp}] {msg}"
    log_history.append(full_msg)
    try:
        if log_text.winfo_exists():
            log_text.config(state="normal")
            log_text.insert("end", full_msg + "\n")
            log_text.see("end")
            log_text.config(state="disabled")
    except Exception:
        # 若視窗已關閉則忽略
        pass

def update_group_name(*args):
    # 更新所有 row2 下拉選單的顯示名稱
    new_values = [""] + [group_display_names[c].get() for c in group_codes]
    for _, var1, var2, var3, var4, combo1, combo2, combo3, combo4 in checkbox_vars_entries:
        # 暫存原本的值
        v1, v2, v3, v4 = var1.get(), var2.get(), var3.get(), var4.get()
        combo1.config(values=new_values)
        combo2.config(values=new_values)
        combo3.config(values=new_values)
        combo4.config(values=new_values)
        # 如果原值還在新 values 裡，還原
        if v1 in new_values:
            var1.set(v1)
        if v2 in new_values:
            var2.set(v2)
        if v3 in new_values:
            var3.set(v3)
        if v4 in new_values:
            var4.set(v4)
    # 更新所有啟動/關閉按鈕的顯示名稱
    for code in group_codes:
        if code in group_buttons:
            group_buttons[code].config(text=f"{group_display_names[code].get()}")
        if code in close_buttons:
            close_buttons[code].config(text=f"{group_display_names[code].get()}")

def load_settings():
    global schedule_tasks, window_layouts
    if not os.path.exists(SETTINGS_FILE):
        return
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        folder_var.set(data.get("folder", folder_var.get()))
        interval_var.set(data.get("interval", interval_var.get()))
        for c in group_codes:
            group_display_names[c].set(data.get("group_display_names", {}).get(c, c))
        for i, v in enumerate(data.get("group_hotkeys", [])):
            if i < len(group_hotkeys):
                group_hotkeys[i].set(v)
        entries = data.get("checkbox_entries", [])
        for i, entry in enumerate(entries):
            if i < len(checkbox_vars_entries):
                ent = checkbox_vars_entries[i][0]
                ent.config(state="normal")
                ent.delete(0, END)
                ent.insert(0, entry)
                ent.config(state="readonly")
        # 讀取所有 group_var1~4
        for idx, key in enumerate(["group_var1", "group_var2", "group_var3", "group_var4"]):
            group_vars = data.get(key, [])
            for i, v in enumerate(group_vars):
                if i < len(checkbox_vars_entries):
                    checkbox_vars_entries[i][1+idx].set(v)
        
        # 載入排程任務
        schedule_tasks = data.get("schedule_tasks", [])
        
        # 載入視窗佈局
        window_layouts = data.get("window_layouts", {})
        
        update_show_labels()
        update_group_name()
    except Exception as e:
        # 只在 log 函數存在時才記錄錯誤
        try:
            log(f"設定檔讀取失敗: {e}")
        except:
            print(f"設定檔讀取失敗: {e}")
# 在所有重要變動時呼叫 save_settings
def on_any_change(*args):
    save_settings()

folder_var.trace_add("write", on_any_change)
interval_var.trace_add("write", on_any_change)
for c in group_codes:
    group_display_names[c].trace_add("write", on_any_change)
for v in group_hotkeys:
    v.trace_add("write", on_any_change)
for entry, *_ in checkbox_vars_entries:
    entry.bind("<FocusOut>", lambda e: save_settings())
    entry.bind("<KeyRelease>", lambda e: save_settings())
for _, var1, var2, *_ in checkbox_vars_entries:
    var1.trace_add("write", on_any_change)
    var2.trace_add("write", on_any_change)

# 關閉時自動儲存
atexit.register(save_settings)

# 啟動時自動載入
load_settings()

# --- 啟動時延遲0.5秒再讀取設定檔 ---
def delayed_load_settings():
    time.sleep(0.5)
    app.after(0, lambda: [load_settings(), update_file_list(), update_window_list()])
threading.Thread(target=delayed_load_settings, daemon=True).start()

# 綁定分組名稱變動時自動更新
for c in group_codes:
    group_display_names[c].trace_add("write", update_group_name)

# 新增：視窗滾動條隱藏與顯示
def _on_window_vsb(*args):
    window_list_vsb.set(*args)
    # 自動隱藏/顯示
    if float(args[0]) <= 0.0 and float(args[1]) >= 1.0:
        window_list_vsb.grid_remove()
    else:
        window_list_vsb.grid(row=0, column=1, sticky="ns")
window_list_canvas.configure(yscrollcommand=_on_window_vsb)


def show_about_dialog():
    about_win = tb.Toplevel(app)
    about_win.title("關於 ChroLens_Portal")
    about_win.geometry("450x300")
    about_win.resizable(False, False)
    about_win.grab_set()
    # 置中顯示
    app.update_idletasks()
    x = app.winfo_x() + (app.winfo_width() // 2) - 175
    y = app.winfo_y() + 80
    about_win.geometry(f"+{x}+{y}")

    # 設定icon與主程式相同
    try:
        import sys, os
        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(sys._MEIPASS, "冥想貓貓.ico")
        else:
            icon_path = "冥想貓貓.ico"
        about_win.iconbitmap(icon_path)
    except Exception as e:
        print(f"無法設定 about 視窗 icon: {e}")

    frm = tb.Frame(about_win, padding=20)
    frm.pack(fill="both", expand=True)

    tb.Label(frm, text="ChroLens_Portal\n分組開啟/關閉程式\n分組視窗置頂顯示", font=("Microsoft JhengHei", 11,)).pack(anchor="w", pady=(0, 6))
    link = tk.Label(frm, text="ChroLens_模擬器討論區", font=("Microsoft JhengHei", 10, "underline"), fg="#5865F2", cursor="hand2")
    link.pack(anchor="w")
    link.bind("<Button-1>", lambda e: os.startfile("https://discord.gg/72Kbs4WPPn"))
    github = tk.Label(frm, text="查看更多工具(巴哈)", font=("Microsoft JhengHei", 10, "underline"), fg="#24292f", cursor="hand2")
    github.pack(anchor="w", pady=(8, 0))
    github.bind("<Button-1>", lambda e: os.startfile("https://home.gamer.com.tw/profile/index_creation.php?owner=umiwued&folder=524003"))
    
    # 作者資訊（包含贊助連結）
    author_frame = tb.Frame(frm)
    author_frame.pack(anchor="w", pady=(8, 6))
    tb.Label(author_frame, text="Created by Lucienwooo ", font=("Microsoft JhengHei", 11)).pack(side="left")
    kofi_link = tk.Label(author_frame, text="(好啦給你一點錢錢)", font=("Microsoft JhengHei", 10, "underline"), fg="#FF5E5B", cursor="hand2")
    kofi_link.pack(side="left")
    kofi_link.bind("<Button-1>", lambda e: os.startfile("https://ko-fi.com/B0B51FBVA8"))
    
    # 按鈕區域
    button_frame = tb.Frame(frm)
    button_frame.pack(anchor="e", pady=(16, 0))
    
    def check_for_updates():
        """檢查更新"""
        about_win.withdraw()  # 暫時隱藏關於視窗
        
        def check_thread():
            try:
                updater = UpdateManager(CURRENT_VERSION, logger=log)
                update_info = updater.check_for_updates()
                
                # 在主執行緒更新 UI
                app.after(0, lambda: show_update_result(update_info, updater))
            except Exception as e:
                app.after(0, lambda: show_update_error(str(e)))
        
        def show_update_result(update_info, updater):
            if update_info:
                # 有新版本
                about_win.destroy()  # 關閉關於視窗
                UpdateDialog(app, updater, update_info)
            else:
                # 已是最新版本
                about_win.deiconify()  # 恢復關於視窗
                NoUpdateDialog(about_win, CURRENT_VERSION)
        
        def show_update_error(error):
            about_win.deiconify()
            messagebox.showerror("檢查更新失敗", f"無法檢查更新：\n{error}", parent=about_win)
        
        # 在背景執行緒中檢查更新
        threading.Thread(target=check_thread, daemon=True).start()
    
    tb.Button(button_frame, text="檢查更新", command=check_for_updates, width=10, bootstyle=INFO).pack(side="left", padx=(0, 5))
    tb.Button(button_frame, text="關閉", command=about_win.destroy, width=8, bootstyle=SECONDARY).pack(side="left")

# --- row 0 新增「刷新視窗」按鈕（SVG 圖示版）---
def manual_refresh_window_list():
    if mini_mode_active:
        log("Mini 模式下無需刷新視窗列表")
        return
    update_window_list()
    log("已刷新視窗列表")

# 建立重新整理圖示
refresh_icon = """
<svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
  <path d="M11.534 7h3.932a.25.25 0 0 1 .192.41l-1.966 2.36a.25.25 0 0 1-.384 0l-1.966-2.36a.25.25 0 0 1 .192-.41zm-11 2h3.932a.25.25 0 0 0 .192-.41L2.692 6.23a.25.25 0 0 0-.384 0L.342 8.59A.25.25 0 0 0 .534 9z"/>
  <path fill-rule="evenodd" d="M8 3c-1.552 0-2.94.707-3.857 1.818a.5.5 0 1 1-.771-.636A6.002 6.002 0 0 1 13.917 7H12.9A5.002 5.002 0 0 0 8 3zM3.1 9a5.002 5.002 0 0 0 8.757 2.182.5.5 0 1 1 .771.636A6.002 6.002 0 0 1 2.083 9H3.1z"/>
</svg>
"""

# 使用 Unicode 刷新符號作為替代方案（因為 ttkbootstrap 不直接支援 SVG）
# 注意：ttkbootstrap.Button 不支援 font 參數
# 使用 Label 包裝來實現較大的 emoji 圖示
refresh_frame = tb.Frame(top_row_frame)
refresh_frame.grid(row=0, column=10, padx=(2,2), sticky="e")
refresh_font = tkfont.Font(family="Segoe UI Emoji", size=16, weight="bold")
refresh_label = tb.Label(refresh_frame, text="🔄", font=refresh_font, cursor="hand2")
refresh_label.pack()
refresh_label.bind("<Button-1>", lambda e: manual_refresh_window_list())

# --- 排程功能 ---
def show_schedule_dialog():
    """顯示排程設定視窗"""
    schedule_win = tb.Toplevel(app)
    schedule_win.title("排程設定")
    schedule_win.geometry("600x400")
    schedule_win.resizable(True, True)
    
    # 置中顯示
    app.update_idletasks()
    x = app.winfo_x() + (app.winfo_width() // 2) - 300
    y = app.winfo_y() + 50
    schedule_win.geometry(f"600x400+{x}+{y}")
    
    # 設定 icon
    try:
        schedule_win.iconbitmap(ico_path)
    except Exception:
        pass
    
    # 主框架
    main_frame = tb.Frame(schedule_win, padding=10)
    main_frame.pack(fill="both", expand=True)
    
    # 響應式配置
    main_frame.grid_rowconfigure(1, weight=1)
    main_frame.grid_columnconfigure(0, weight=1)
    
    # 上方控制區
    control_frame = tb.Frame(main_frame)
    control_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
    
    # 分組選擇
    tb.Label(control_frame, text="分組:").grid(row=0, column=1, padx=(10, 5), sticky="w")
    group_var = tk.StringVar(value=group_codes[0])
    group_combo = tb.Combobox(
        control_frame, 
        textvariable=group_var,
        values=[f"{c} ({group_display_names[c].get()})" for c in group_codes],
        width=15,
        state="readonly"
    )
    group_combo.grid(row=0, column=2, padx=(0, 10), sticky="w")
    
    # 時間設定
    tb.Label(control_frame, text="時間:").grid(row=0, column=3, padx=(10, 5), sticky="w")
    hour_var = tk.StringVar(value="00")
    minute_var = tk.StringVar(value="00")
    
    time_frame = tb.Frame(control_frame)
    time_frame.grid(row=0, column=4, padx=(0, 10), sticky="w")
    
    # 改為下拉式選單 (24小時制)
    hour_combo = tb.Combobox(time_frame, textvariable=hour_var, width=3, state="readonly",
                             values=[f"{h:02d}" for h in range(24)])
    hour_combo.pack(side="left")
    tb.Label(time_frame, text=":").pack(side="left", padx=2)
    minute_combo = tb.Combobox(time_frame, textvariable=minute_var, width=3, state="readonly",
                               values=[f"{m:02d}" for m in range(60)])
    minute_combo.pack(side="left")
    
    # 新增按鈕
    def add_schedule():
        group_text = group_var.get()
        group_code = group_text.split()[0]  # 取得分組代碼（A, B, C等）
        time_str = f"{hour_var.get()}:{minute_var.get()}"
        
        # 檢查是否已存在相同的排程
        for task in schedule_tasks:
            if task["group"] == group_code and task["time"] == time_str:
                log(f"排程已存在：{group_display_names[group_code].get()} 於 {time_str}")
                return
        
        schedule_tasks.append({
            "group": group_code,
            "time": time_str,
            "enabled": True
        })
        update_schedule_list()
        save_settings()  # 儲存排程到設定檔
        log(f"已新增排程：{group_display_names[group_code].get()} 於 {time_str}")
    
    add_btn = tb.Button(control_frame, text="新增", command=add_schedule, bootstyle="success", width=8)
    add_btn.grid(row=0, column=5, padx=(10, 0), sticky="w")
    
    # 排程列表區域
    list_frame = tb.Frame(main_frame, borderwidth=1, relief="solid")
    list_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
    list_frame.grid_rowconfigure(0, weight=1)
    list_frame.grid_columnconfigure(0, weight=1)
    
    # Canvas + Scrollbar
    canvas = tk.Canvas(list_frame, highlightthickness=0)
    canvas.grid(row=0, column=0, sticky="nsew")
    
    vsb = tb.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
    vsb.grid(row=0, column=1, sticky="ns")
    canvas.configure(yscrollcommand=vsb.set)
    
    inner_frame = tb.Frame(canvas)
    canvas_window = canvas.create_window((0, 0), window=inner_frame, anchor="nw")
    
    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(canvas_window, width=canvas.winfo_width())
    
    inner_frame.bind("<Configure>", on_frame_configure)
    
    # 更新排程列表顯示
    def update_schedule_list():
        for widget in inner_frame.winfo_children():
            widget.destroy()
        
        if not schedule_tasks:
            tb.Label(inner_frame, text="尚無排程", font=("Microsoft JhengHei", 10), foreground="#888").pack(pady=20)
            return
        
        for idx, task in enumerate(schedule_tasks):
            task_frame = tb.Frame(inner_frame, borderwidth=1, relief="groove", padding=5)
            task_frame.pack(fill="x", padx=5, pady=2)
            
            # 啟用/停用開關（只保留綠色的）
            enabled_var = tk.BooleanVar(value=task["enabled"])
            
            def toggle_task(index=idx, var=enabled_var):
                schedule_tasks[index]["enabled"] = var.get()
                save_settings()
                log(f"排程 {schedule_tasks[index]['time']} {group_display_names[schedule_tasks[index]['group']].get()} {'已啟用' if var.get() else '已停用'}")
            
            check = tb.Checkbutton(
                task_frame, 
                text="",
                variable=enabled_var,
                command=toggle_task,
                bootstyle="success-round-toggle"
            )
            check.pack(side="left", padx=(0, 10))
            
            # 時間顯示
            time_label = tb.Label(
                task_frame,
                text=task["time"],
                font=("Consolas", 12, "bold"),
                width=8
            )
            time_label.pack(side="left", padx=(0, 10))
            
            # 分組顯示（只顯示使用者設定的名稱）
            group_display = group_display_names[task['group']].get()
            group_label = tb.Label(
                task_frame,
                text=group_display,
                font=("Microsoft JhengHei", 10)
            )
            group_label.pack(side="left", padx=(0, 10))
            
            # 刪除按鈕
            def delete_task(index=idx):
                removed = schedule_tasks.pop(index)
                update_schedule_list()
                save_settings()
                log(f"已刪除排程：{group_display_names[removed['group']].get()} 於 {removed['time']}")
            
            delete_btn = tb.Button(
                task_frame,
                text="刪除",
                command=delete_task,
                bootstyle="danger-outline",
                width=6
            )
            delete_btn.pack(side="right", padx=(5, 0))
        
        inner_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
    
    # 底部控制按鈕
    bottom_frame = tb.Frame(main_frame)
    bottom_frame.grid(row=2, column=0, sticky="ew")
    
    # 開啟 Windows 工作排程器
    def open_task_scheduler():
        try:
            subprocess.Popen("taskschd.msc", shell=True)
            log("已開啟 Windows 工作排程器")
        except Exception as e:
            log(f"無法開啟工作排程器：{e}")
    
    task_scheduler_btn = tb.Button(
        bottom_frame,
        text="檢視工作排程",
        command=open_task_scheduler,
        bootstyle="info",
        width=15
    )
    task_scheduler_btn.pack(side="left", padx=(0, 10))
    
    # 關閉按鈕
    close_btn = tb.Button(
        bottom_frame,
        text="關閉",
        command=schedule_win.destroy,
        bootstyle="secondary",
        width=8
    )
    close_btn.pack(side="right")
    
    # 初始化列表
    update_schedule_list()

# --- 新增：mini 模式切換按鈕 ---
is_mini_mode = tk.BooleanVar(value=False)

def restore_from_mini():
    """從 mini 模式還原到一般模式"""
    is_mini_mode.set(False)
    mini_btn.config(text="mini")
    restore_normal_mode()

def toggle_mini_mode():
    if is_mini_mode.get():
        # 切換回一般模式
        is_mini_mode.set(False)
        restore_normal_mode()
    else:
        # 切換為 mini 模式
        is_mini_mode.set(True)
        enter_mini_mode()

def enter_mini_mode():
    """進入 mini 模式：隱藏除了 row1 和 row8~9 以外的所有元件"""
    global mini_mode_active
    mini_mode_active = True
    
    # 隱藏 row 0（頂部工具列）
    top_row_frame.grid_remove()
    # 隱藏 row 2（分組檔案列）
    for gf in group_frames:
        gf.grid_remove()
    # 隱藏 row 10（檔案與視窗列表）
    bottom_frame.grid_remove()
    # 隱藏動態日誌
    log_text.grid_remove()
    # 隱藏「置頂切換」文字標籤
    desc_label.grid_remove()
    # 隱藏所有快捷鍵輸入框，只保留分組按鈕
    for show_label, hotkey_entry in show_label_frames:
        hotkey_entry.pack_forget()
    # 顯示 mini 還原按鈕
    mini_restore_frame.grid(row=0, column=0, padx=2, pady=2, sticky="ew")
    
    # 調整 mini 模式下的 row 配置，讓按鈕區域緊湊排列
    frm.grid_rowconfigure(1, weight=0)  # 置頂切換區固定
    frm.grid_rowconfigure(8, weight=0)  # 按鈕區固定
    frm.grid_rowconfigure(9, weight=0)  # 按鈕區固定
    # 移除其他 row 的配置
    for row_num in [0, 2, 3, 4, 5, 6, 7, 10]:
        frm.grid_rowconfigure(row_num, weight=0, minsize=0)
    
    # 調整 row 1 的 padding，移除多餘空間
    second_row_frame.grid_configure(pady=(0, 2))
    
    # 將按鈕區域移到 row 2 位置，緊接在置頂切換區下方
    btns_outer_frame.grid_configure(row=2, column=1, rowspan=2, columnspan=6, sticky="new", padx=(0, 4), pady=(0, 0))
    
    # 調整視窗大小為更緊湊的尺寸（高度減少）
    app.geometry("600x120")
    log("已進入 mini 模式（降低資源使用）")

def restore_normal_mode():
    """還原一般模式：顯示所有元件"""
    global mini_mode_active
    mini_mode_active = False
    
    # 還原響應式 row 配置
    frm.grid_rowconfigure(0, weight=0)  # 頂部工具列 - 固定
    frm.grid_rowconfigure(1, weight=0)  # 置頂切換區 - 固定
    frm.grid_rowconfigure(2, weight=1)  # 分組檔案列 - 可擴展
    frm.grid_rowconfigure(8, weight=0)  # 按鈕區 - 固定
    frm.grid_rowconfigure(9, weight=0)  # 按鈕區 - 固定
    frm.grid_rowconfigure(10, weight=1)  # 檔案/視窗列表 - 可擴展
    
    # 還原 row 1 的 padding
    second_row_frame.grid_configure(pady=(0, 0))
    
    # 還原按鈕區域的位置到 row 8
    btns_outer_frame.grid_configure(row=8, column=1, rowspan=2, columnspan=6, sticky="new", padx=(0, 4), pady=(0, 0))
    
    # 顯示 row 0
    top_row_frame.grid()
    # 顯示 row 2
    for idx, gf in enumerate(group_frames):
        gf.grid()
    # 顯示 row 10
    bottom_frame.grid()
    # 顯示動態日誌
    log_text.grid()
    # 顯示「置頂切換」文字標籤
    desc_label.grid(row=0, column=0, padx=2, pady=2, sticky="ew")
    # 顯示所有快捷鍵輸入框
    for show_label, hotkey_entry in show_label_frames:
        hotkey_entry.pack(side="left", padx=(2,5))
    # 隱藏 mini 還原按鈕
    mini_restore_frame.grid_remove()
    # 還原視窗大小
    app.geometry("1200x720")
    
    # 還原後立即更新檔案和視窗列表
    update_file_list()
    update_window_list()
    log("已還原一般模式")

# --- 排程按鈕 ---
schedule_btn = tb.Button(top_row_frame, text="排程", command=show_schedule_dialog, bootstyle="warning", width=6)
schedule_btn.grid(row=0, column=12, padx=(2,2), sticky="e")

mini_btn = tb.Button(top_row_frame, text="mini", command=toggle_mini_mode, bootstyle=INFO, width=5)
mini_btn.grid(row=0, column=11, padx=(2,2), sticky="e")

# 更新 mini_restore_label 的點擊事件為實際的還原函數
mini_restore_label.bind("<Button-1>", lambda e: restore_from_mini())

# --- 刷新按鈕 ---
def refresh_lists():
    """刷新檔案列表和視窗列表"""
    update_file_list()
    update_window_list()
    log("已刷新檔案和視窗列表")

refresh_btn = tb.Button(top_row_frame, text="🔄", command=refresh_lists, bootstyle="info", width=3)
refresh_btn.grid(row=0, column=10, padx=(2,2), sticky="e")

# --- 關於按鈕（原本已存在） ---
about_btn = tb.Button(top_row_frame, text="關於", command=show_about_dialog, bootstyle=SECONDARY, width=6)
about_btn.grid(row=0, column=9, padx=(8,2), sticky="e")


def get_group_files(group_code):
    """取得指定分組的檔案名稱清單（從 row2 的 entry 取得）"""
    files = []
    for entry, var1, var2, var3, var4, *_ in checkbox_vars_entries:
        # 判斷這一行是否屬於該分組
        if (
            var1.get() == group_display_names[group_code].get()
            or var2.get() == group_display_names[group_code].get()
            or var3.get() == group_display_names[group_code].get()
            or var4.get() == group_display_names[group_code].get()
        ):
            val = entry.get().strip()
            if val:
                files.append(val)
    return files

# --- 分組循環聚焦功能 ---
group_focus_indexes = {code: 0 for code in group_codes}
group_hwnd_lists = {code: [] for code in group_codes}

def update_group_hwnd_list(group_code):
    files = get_group_files(group_code)
    target_titles = [os.path.splitext(os.path.basename(f))[0].lower() for f in files if f]
    hwnds = []
    def enum_handler(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        window_text = win32gui.GetWindowText(hwnd).lower().strip()
        if any(title and title in window_text for title in target_titles):
            hwnds.append(hwnd)
    win32gui.EnumWindows(enum_handler, None)
    group_hwnd_lists[group_code] = hwnds

def focus_next_in_group(group_code):
    """批次同時置頂分組中的所有視窗並恢復佈局（整合 FancyZones 功能）"""
    log(f"[快捷鍵] 觸發分組 {group_display_names[group_code].get()}")
    
    update_group_hwnd_list(group_code)
    hwnds = group_hwnd_lists[group_code]
    
    if not hwnds:
        log(f"[快捷鍵] 分組 {group_display_names[group_code].get()} 沒有視窗")
        return
    
    log(f"[快捷鍵] 開始置頂並恢復分組 {group_display_names[group_code].get()} 的 {len(hwnds)} 個視窗佈局")
    
    def topmost_and_restore_windows():
        """在背景執行緒中批次置頂視窗並恢復佈局"""
        try:
            # 獲取所有視窗的標題（用於匹配佈局）
            hwnd_titles = {}
            for hwnd in hwnds:
                try:
                    title = win32gui.GetWindowText(hwnd)
                    hwnd_titles[hwnd] = title
                except Exception:
                    hwnd_titles[hwnd] = ""
            
            # 第一步：還原所有最小化的視窗
            for hwnd in hwnds:
                try:
                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                except Exception:
                    pass
            
            # 第二步：嘗試恢復每個視窗的佈局
            restored_count = 0
            for hwnd in hwnds:
                title = hwnd_titles.get(hwnd, "")
                if title and restore_window_layout(group_code, hwnd, title):
                    restored_count += 1
            
            if restored_count > 0:
                app.after(0, lambda: log(f"已恢復 {restored_count} 個視窗的佈局"))
            
            # 短暫延遲確保佈局恢復完成
            time.sleep(0.15)
            
            # 第三步：將所有視窗置頂（使用 HWND_TOPMOST 然後立即取消）
            for hwnd in hwnds:
                try:
                    # 先置頂
                    win32gui.SetWindowPos(
                        hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER
                    )
                except Exception:
                    pass
            
            # 短暫延遲
            time.sleep(0.05)
            
            # 第四步：取消永久置頂（但視窗已經在最上層了）
            for hwnd in hwnds:
                try:
                    win32gui.SetWindowPos(
                        hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER
                    )
                except Exception:
                    pass
            
            # 第五步：激活第一個視窗
            if hwnds:
                try:
                    # 使用 SetForegroundWindow 而不是 ShowWindow，避免改變視窗大小
                    win32gui.SetForegroundWindow(hwnds[0])
                except Exception:
                    pass
            
            app.after(0, lambda: log(f"已完成置頂分組 {group_display_names[group_code].get()} 的所有 {len(hwnds)} 個視窗"))
        except Exception as e:
            app.after(0, lambda msg=f"置頂處理異常: {e}": log(msg))
    
    # 在背景執行緒中執行置頂操作
    threading.Thread(target=topmost_and_restore_windows, daemon=True).start()


# --- 全域快捷鍵系統（簡化版，穩定可靠） ---
hotkey_handlers = []

def register_global_hotkeys():
    """註冊全域快捷鍵"""
    global hotkey_handlers
    
    # 清除舊的快捷鍵
    try:
        for handler in hotkey_handlers:
            try:
                keyboard.remove_hotkey(handler)
            except:
                pass
        hotkey_handlers.clear()
    except:
        pass
    
    # 註冊新的快捷鍵
    success_count = 0
    for idx, code in enumerate(group_codes):
        hotkey = group_hotkeys[idx].get()
        try:
            # 直接使用小寫格式
            hotkey_str = hotkey.lower().replace(' ', '')
            
            # 創建回調函數（使用閉包捕獲正確的 code 值）
            def create_hotkey_callback(group_code):
                def callback():
                    try:
                        log(f"[快捷鍵] {hotkey} 已觸發 -> 分組 {group_display_names[group_code].get()}")
                        focus_next_in_group(group_code)
                    except Exception as e:
                        log(f"[快捷鍵] 執行錯誤: {e}")
                return callback
            
            # 註冊快捷鍵（按下時立即觸發）
            handler = keyboard.add_hotkey(
                hotkey_str,
                create_hotkey_callback(code),
                suppress=False
            )
            
            hotkey_handlers.append(handler)
            success_count += 1
            log(f"✓ {hotkey} → 分組 {group_display_names[code].get()}")
            
        except Exception as e:
            log(f"✗ 註冊失敗 {hotkey}: {e}")
    
    if success_count > 0:
        log(f"共註冊 {success_count} 個快捷鍵")
    else:
        log("警告：快捷鍵註冊失敗，請以管理員權限執行")

def cleanup_hotkeys():
    """清理所有快捷鍵"""
    try:
        for handler in hotkey_handlers:
            try:
                keyboard.remove_hotkey(handler)
            except:
                pass
        hotkey_handlers.clear()
    except:
        pass

# 程式結束時清理
atexit.register(cleanup_hotkeys)

# 熱鍵變更時重新註冊
for var in group_hotkeys:
    var.trace_add("write", lambda *args: register_global_hotkeys())

# 延遲註冊快捷鍵
def init_hotkeys():
    time.sleep(2)  # 等待 UI 完全初始化
    try:
        register_global_hotkeys()
    except Exception as e:
        app.after(0, lambda: log(f"快捷鍵初始化失敗: {e}"))

threading.Thread(target=init_hotkeys, daemon=True).start()

def open_lnk_target(lnk_path):
    """解析 .lnk 捷徑檔案，回傳 (目標路徑, 參數字串)"""
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

def open_entry_file(entry):
    file_path = entry.get().strip()
    folder = folder_var.get()
    if not file_path:
        log("此欄位無檔案名稱")
        return
    full_path = os.path.join(folder, file_path)
    if not os.path.exists(full_path):
        log(f"找不到檔案: {full_path}")
        return
    try:
        if full_path.lower().endswith('.lnk'):
            # 解析捷徑
            target, args = open_lnk_target(full_path)
            if target and os.path.exists(target):
                log(f"開啟捷徑目標: {target} {args}")
                subprocess.Popen(
                    f'"{target}" {args}',
                    shell=True,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                log(f"無法解析捷徑或目標不存在: {full_path}")
        else:
            os.startfile(full_path)
            log(f"已開啟檔案: {full_path}")
    except Exception as e:
        log(f"開啟檔案失敗: {e}")

app.mainloop()