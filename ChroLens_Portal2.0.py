### ChroLens_Portal 2.0 
### 2025/05/26 By Lucienwooo
### pyinstaller --onedir --noconsole --add-data "冥想貓貓.ico;." --icon=冥想貓貓.ico --hidden-import=win32timezone ChroLens_Portal2.0.py 
###### 分組視窗透過快捷鍵最上層顯示，半成品。
# 新增清單窗格，顯示當前所有開啟的視窗名稱
# 並且讓選單可以記憶所以從清單選取的檔案/視窗名稱
# 程式將不再只是能開啟或關閉程式
import os
import time
import win32gui
import win32con
import threading
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, END
import tkinter as tk
import pythoncom
from win32com.shell import shell, shellcon
import subprocess
import sys
import json
import win32con
import keyboard  # 請確保已安裝
from tkinter import font as tkfont
import functools

LAST_PATH_FILE = "last_path.txt"
SETTINGS_FILE = "chrolens_portal.json"

def resource_path(relative_path):
    """取得資源檔案的絕對路徑，支援 PyInstaller 打包後的路徑"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def open_lnk_target(lnk_path):
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

def open_files_in_folder(folder_path, interval=4, log_func=None):
    selected_files = []
    for var, entry in checkbox_vars_entries:
        if var.get():
            selected_files.append(entry.get())
    files = selected_files if selected_files else [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    files.sort()
    for file in files:
        file_path = os.path.join(folder_path, file)
        try:
            if file.lower().endswith('.lnk'):
                target, args = open_lnk_target(file_path)
                if target and os.path.exists(target):
                    if log_func:
                        log_func(f"Opening shortcut target: {target} {args}")
                    subprocess.Popen(
                        f'"{target}" {args}',
                        shell=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                else:
                    if log_func:
                        log_func(f"無法解析捷徑或目標不存在: {file_path}")
            else:
                if log_func:
                    log_func(f"Opening: {file_path}")
                os.startfile(file_path)
        except Exception as e:
            if log_func:
                log_func(f"無法開啟: {file_path}，錯誤：{e}")
        time.sleep(interval)

def open_files_by_names(folder_path, file_names, interval=4, log_func=None):
    files = [f for f in file_names if f]
    for file in files:
        file_path = os.path.join(folder_path, file)
        try:
            if file.lower().endswith('.lnk'):
                target, args = open_lnk_target(file_path)
                if target and os.path.exists(target):
                    if log_func:
                        log_func(f"Opening shortcut target: {target} {args}")
                    subprocess.Popen(
                        f'"{target}" {args}',
                        shell=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                else:
                    if log_func:
                        log_func(f"無法解析捷徑或目標不存在: {file_path}")
            else:
                if log_func:
                    log_func(f"Opening: {file_path}")
                os.startfile(file_path)
        except Exception as e:
            if log_func:
                log_func(f"無法開啟: {file_path}，錯誤：{e}")
        time.sleep(interval)

def start_opening():
    folder = folder_var.get()
    try:
        interval = float(interval_var.get())
    except ValueError:
        log("請輸入正確的間隔秒數")
        return
    if not os.path.isdir(folder):
        log("請選擇正確的資料夾")
        return
    log(f"開始開啟 {folder} 內的檔案，每 {interval} 秒一個")
    threading.Thread(target=open_files_in_folder, args=(folder, interval, log), daemon=True).start()

def choose_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_var.set(folder)
        save_last_path(folder)
        show_files_in_folder(folder)
        update_checkboxes(folder)
        update_file_list()  # 新增：刷新下方檔案清單

def show_files_in_folder(folder):
    pass

def show_log_window():
    if hasattr(app, 'log_win') and app.log_win.winfo_exists():
        app.log_win.deiconify()
        return
    log_win = tk.Toplevel(app)
    log_win.title("動態紀錄")
    log_win.geometry(f"400x600+{app.winfo_rootx()-400}+{app.winfo_rooty()}")  # 主程式左側外面
    log_win.resizable(True, True)
    log_win.attributes('-topmost', True)
    app.log_win = log_win

    # 可拖曳
    def start_move(event):
        log_win._drag_start_x = event.x
        log_win._drag_start_y = event.y
    def do_move(event):
        x = log_win.winfo_x() + event.x - log_win._drag_start_x
        y = log_win.winfo_y() + event.y - log_win._drag_start_y
        log_win.geometry(f"+{x}+{y}")
    log_win.bind("<Button-1>", start_move)
    log_win.bind("<B1-Motion>", do_move)

    # 日誌框
    log_text = tb.Text(log_win, height=30, width=60, state="disabled")
    log_text.pack(fill="both", expand=True)
    app.log_text = log_text

def toggle_log():
    if hasattr(app, 'log_win') and app.log_win.winfo_exists():
        if app.log_win.state() == 'normal':
            app.log_win.withdraw()
        else:
            app.log_win.deiconify()
    else:
        show_log_window()

# log() 函數需改为寫入 app.log_text
def log(msg):
    if hasattr(app, 'log_text'):
        app.log_text.configure(state="normal")
        app.log_text.insert("end", msg + "\n")
        app.log_text.see("end")
        app.log_text.configure(state="disabled")

def load_last_path():
    if os.path.exists(LAST_PATH_FILE):
        with open(LAST_PATH_FILE, "r", encoding="utf-8") as f:
            path = f.read().strip()
            if os.path.isdir(path):
                return path
    return os.path.dirname(os.path.abspath(__file__))

def save_last_path(path):
    with open(LAST_PATH_FILE, "w", encoding="utf-8") as f:
        f.write(path)

app = tb.Window(themename="darkly")
app.title("ChroLens_Portal 1.0.0")
try:
    ico_path = resource_path("冥想貓貓.ico")
    app.iconbitmap(ico_path)
except Exception as e:
    print(f"無法設定 icon: {e}")

# 統一字體
default_font = tkfont.Font(family="Microsoft JhengHei", size=12)
app.option_add("*Font", default_font)

# === 分組代碼與顯示名稱 ===
group_codes = ["A", "B", "C", "D"]
group_display_names = {c: tk.StringVar(value=c) for c in group_codes}
group_buttons = {}
close_buttons = {}
# === 介面區塊 ===

folder_var = tb.StringVar(value=load_last_path())
interval_var = tb.StringVar(value="4")

frm = tb.Frame(app, padding=2)
frm.pack(fill="both", expand=True)

top_row_frame = tb.Frame(frm, padding=2)
top_row_frame.grid(row=0, column=0, columnspan=6, sticky="ew", pady=(2, 2))
top_row_frame.grid_columnconfigure(0, weight=0)
top_row_frame.grid_columnconfigure(1, weight=0)
top_row_frame.grid_columnconfigure(2, weight=0)

folder_frame = tb.Frame(top_row_frame, padding=(2,2))
folder_frame.grid(row=0, column=0, sticky="w", padx=(0, 4))
tb.Entry(folder_frame, textvariable=folder_var, width=38).grid(row=0, column=0, padx=(2,2), sticky="ew")
tb.Button(folder_frame, text="選擇開啟路徑", command=choose_folder, bootstyle=SECONDARY).grid(row=0, column=1, padx=(2,0), sticky="ew")

interval_frame = tb.Frame(top_row_frame, padding=(2,2))
interval_frame.grid(row=0, column=1, sticky="w", padx=(0, 4))
tb.Label(interval_frame, text="間隔秒數:").grid(row=0, column=0, sticky="w")
tb.Entry(interval_frame, textvariable=interval_var, width=5).grid(row=0, column=1, padx=(2,0), sticky="w")

# === 新版：快捷鍵設定（僅允許 ALT+任意鍵 或 CTRL+任意鍵）===
default_hotkeys = ["Alt+1", "Alt+2", "Alt+3", "Alt+4"]
group_hotkeys = [tk.StringVar(value=default_hotkeys[i]) for i in range(4)]

def format_hotkey(event):
    # 只允許 ALT+任意 或 CTRL+任意
    keys = []
    if event.state & 0x0004:  # Ctrl
        keys.append("Ctrl")
    if event.state & 0x0008:  # Alt
        keys.append("Alt")
    if not keys:
        return ""  # 沒有 Ctrl/Alt 不允許
    key = event.keysym
    # 過濾掉純修飾鍵
    if key in ("Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R"):
        return ""
    if key.startswith("KP_"):
        key = key.replace("KP_", "Num")
    keys.append(key.capitalize())
    return "+".join(keys)

def on_hotkey_entry_key(event, idx):
    if event.keysym in ("Escape", "BackSpace"):
        group_hotkeys[idx].set(default_hotkeys[idx])
        return "break"
    hotkey = format_hotkey(event)
    if hotkey:
        group_hotkeys[idx].set(hotkey)
    return "break"

# --- 第二排：啟動A/B/C/D顯示快捷鍵（加框） ---
show_label_frames = []
second_row_frame = tb.Frame(frm)
second_row_frame.grid(row=1, column=0, columnspan=6, sticky="ew")
second_row_frame.grid_columnconfigure((0,1,2,3,4), weight=1)  # 讓五個元件平均分配

show_label_font = tkfont.Font(family="Microsoft JhengHei", size=12)  # 統一字體大小

# 新增最左邊的說明文字
desc_label = tb.Label(second_row_frame, text="置頂顯示切換", width=12, anchor="center", font=show_label_font)
desc_label.grid(row=0, column=0, padx=2, pady=2, sticky="ew")

for idx, code in enumerate(group_codes):
    frame = tb.Frame(second_row_frame, borderwidth=2, relief="groove")
    frame.grid(row=0, column=idx+1, padx=2, pady=2, sticky="ew")  # idx+1，往右推一格
    show_label = tb.Label(frame, text=f"{group_display_names[code].get()} 組", width=6, font=show_label_font)
    show_label.pack(side="left")
    hotkey_entry = tb.Entry(frame, textvariable=group_hotkeys[idx], width=8, state="readonly", justify="center", font=show_label_font)
    hotkey_entry.pack(side="left", padx=(2,5))
    def make_on_key(idx):
        return lambda event, i=idx: on_hotkey_entry_key(event, i)
    hotkey_entry.bind("<Key>", make_on_key(idx))
    hotkey_entry.bind("<Button-1>", lambda e, entry=hotkey_entry: entry.focus_set())
    show_label_frames.append((show_label, hotkey_entry))

def update_show_labels(*args):
    for idx, code in enumerate(group_codes):
        show_label_frames[idx][0].config(text=f"{group_display_names[code].get()} 組")  

for c in group_codes:
    group_display_names[c].trace_add("write", update_show_labels)

# === 新版：全域快捷鍵觸發（僅 ALT+任意 或 CTRL+任意）===
def global_hotkey_handler(event):
    pressed = format_hotkey(event)
    for idx, code in enumerate(group_codes):
        if pressed and pressed == group_hotkeys[idx].get():
            set_group_windows_topmost(code)
            log(f"分組 {group_code} 置頂顯示（請補上實際置頂邏輯）")
            break

app.bind_all("<Key>", global_hotkey_handler, add="+")

group_select_code = tk.StringVar(value=group_codes[0])

def get_group_code_by_display_name(display_name):
    for code in group_codes:
        if group_display_names[code].get() == display_name:
            return code
    return None

def set_group_entry_placeholder():
    if not group_rename_var.get():
        group_entry.delete(0, tk.END)
        group_entry.insert(0, "修改分組名稱")
        group_entry.config(foreground="gray")

def clear_group_entry_placeholder(event=None):
    if group_entry.get() == "修改分組名稱":
        group_entry.delete(0, tk.END)
        group_entry.config(foreground="black")

def on_group_entry_focus_out(event=None):
    if not group_entry.get():
        set_group_entry_placeholder()

def update_group_name(*args):
    code = group_select_code.get()
    new_name = group_rename_var.get().strip()
    default_names = {c: c for c in group_codes}
    if not new_name or new_name == "修改分組名稱":
        group_display_names[code].set(default_names[code])
    else:
        group_display_names[code].set(new_name)
    for c in group_codes:
        group_buttons[c]['text'] = f"啟動 {group_display_names[c].get()}"
        close_buttons[c]['text'] = f"關閉 {group_display_names[c].get()}"
    # 重新設定所有下拉選單的 values 與顯示值
    for i in range(15):
        combo1 = checkbox_vars_entries[i][3]
        combo2 = checkbox_vars_entries[i][4]
        # 取得目前選取的顯示名稱
        val1 = combo1.get()
        val2 = combo2.get()
        combo1['values'] = [""] + [group_display_names[c].get() for c in group_codes]
        combo2['values'] = [""] + [group_display_names[c].get() for c in group_codes]
        # 若目前選取的顯示名稱有變動，則同步顯示
        if val1 and val1 not in combo1['values']:
            # 依 group_var1 的 group_code 重新設值
            code1 = get_group_code_by_display_name(val1)
            if code1:
                combo1.set(group_display_names[code1].get())
        if val2 and val2 not in combo2['values']:
            code2 = get_group_code_by_display_name(val2)
            if code2:
                combo2.set(group_display_names[code2].get())
    group_combo['values'] = [group_display_names[c].get() for c in group_codes]
    group_select_code.set(code)
    group_rename_var.set("")
    set_group_entry_placeholder()

group_combo = tb.Combobox(
    top_row_frame, textvariable=group_select_code,
    values=[group_display_names[c].get() for c in group_codes], width=4, state="readonly"
)
group_combo.grid(row=0, column=3, padx=(8,2))

group_rename_var = tk.StringVar()
group_entry = tb.Entry(top_row_frame, textvariable=group_rename_var, width=14)
group_entry.grid(row=0, column=4, padx=(2,0))
group_entry.bind("<FocusIn>", clear_group_entry_placeholder)
group_entry.bind("<FocusOut>", on_group_entry_focus_out)
group_entry.bind("<Return>", lambda e: update_group_name())
set_group_entry_placeholder()

def on_group_combo_selected(event=None):
    display_name = group_combo.get()
    code = get_group_code_by_display_name(display_name)
    if code:
        group_select_code.set(code)
        group_combo.set(group_display_names[code].get())

group_combo.bind("<<ComboboxSelected>>", on_group_combo_selected)

group_frames = []
for col in range(3):
    group_frame = tb.Frame(frm, borderwidth=1, relief="solid", padding=2)
    group_frame.grid(row=2, column=col, padx=2, pady=2, sticky="n")
    group_frames.append(group_frame)

# --- 分組檔案列（數字放大、粗體、分組更緊湊） ---
num_font = tkfont.Font(family="Microsoft JhengHei", size=12, weight="bold")  # 放大一級且粗體

checkbox_vars_entries = []
for i in range(15):
    entry = tb.Entry(group_frames[i // 5], width=14, state="readonly")
    group_var1 = tk.StringVar(value="")
    group_var2 = tk.StringVar(value="")
    group_combo1 = tb.Combobox(
        group_frames[i // 5], textvariable=group_var1,
        values=[""] + [group_display_names[c].get() for c in group_codes], width=4, state="readonly"
    )
    group_combo2 = tb.Combobox(
        group_frames[i // 5], textvariable=group_var2,
        values=[""] + [group_display_names[c].get() for c in group_codes], width=4, state="readonly"
    )
    row = i % 5
    # 數字標籤：放大、粗體，無多餘空格
    num_label = tb.Label(group_frames[i // 5], text=str(i+1), width=2)
    num_label['font'] = num_font
    num_label.grid(row=row, column=0, sticky="w", padx=0)
    # 其餘欄位緊湊排列，無多餘 padx
    entry.grid(row=row, column=1, padx=0, pady=1)
    group_combo1.grid(row=row, column=2, padx=0, pady=1)
    group_combo2.grid(row=row, column=3, padx=0, pady=1)
    checkbox_vars_entries.append((entry, group_var1, group_var2, group_combo1, group_combo2))

def get_group_files(group_code):
    files = []
    for entry, group_var1, group_var2, *_ in checkbox_vars_entries:
        code1 = get_group_code_by_display_name(group_var1.get())
        code2 = get_group_code_by_display_name(group_var2.get())
        filename = entry.get()
        if filename and (group_code == code1 or group_code == code2):
            files.append(filename)
    return files

# --- 啟動/關閉/動態紀錄按鈕區域 ---
btns_outer_frame = tb.Frame(frm)
btns_outer_frame.grid(row=8, column=0, columnspan=6, sticky="ew", padx=(0, 4), pady=(8, 4))
for i in range(5):
    btns_outer_frame.grid_columnconfigure(i, weight=1)

from tkinter import font as tkfont
big_btn_font = tkfont.Font(family="Microsoft JhengHei", size=18, weight="bold")
mid_btn_font = tkfont.Font(family="Microsoft JhengHei", size=13, weight="bold")  # 放大一級

style = tb.Style()
style.configure("Big.TButton", font=big_btn_font)
style.configure("Mid.TButton", font=mid_btn_font)

dynamic_log_btn = tb.Button(
    btns_outer_frame, text="動態\n紀錄", width=6, command=toggle_log, bootstyle="info", style="Big.TButton"
)
dynamic_log_btn.grid(row=8, column=0, rowspan=2, sticky="nsew", padx=(0, 8), pady=(0, 0))

# 讓 col=1~4 都能自動撐開
for i in range(1, 5):
    btns_outer_frame.grid_columnconfigure(i, weight=1)

# 啟動/關閉按鈕
group_btn_grid = [
    # row, col, text, group_code, bootstyle, command
    (8, 1, "啟動", "A", "success-outline", lambda: start_group_opening("A")),
    (8, 2, "啟動", "B", "success-outline", lambda: start_group_opening("B")),
    (8, 3, "關閉", "A", "danger-outline", lambda: close_group_windows("A")),
    (8, 4, "關閉", "B", "danger-outline", lambda: close_group_windows("B")),
    (9, 1, "啟動", "C", "success-outline", lambda: start_group_opening("C")),
    (9, 2, "啟動", "D", "success-outline", lambda: start_group_opening("D")),
    (9, 3, "關閉", "C", "danger-outline", lambda: close_group_windows("C")),
    (9, 4, "關閉", "D", "danger-outline", lambda: close_group_windows("D")),
]

for row, col, text, code, bootstyle, cmd in group_btn_grid:
    btn = tb.Button(
        btns_outer_frame,
        text=f"{text} {group_display_names[code].get()}",
        bootstyle=bootstyle,
        command=cmd,
        width=8,
        style="Mid.TButton"
    )
    btn.grid(row=row, column=col, padx=(4, 4), pady=(2, 2), sticky="nsew")
    if text == "啟動":
        group_buttons[code] = btn
    else:
        close_buttons[code] = btn

# 動態紀錄按鈕下方顯示「動態」「紀錄」兩行
btns_outer_frame.grid_rowconfigure(8, weight=1)
btns_outer_frame.grid_rowconfigure(9, weight=1)

# --- 下方清單區塊（左一欄為檔案，右三欄為視窗，四欄平均分配） ---
bottom_frame = tb.Frame(frm)
bottom_frame.grid(row=10, column=0, columnspan=6, sticky="ew", pady=(8, 2))
for i in range(4):
    bottom_frame.grid_columnconfigure(i, weight=1)
bottom_frame.grid_rowconfigure(0, weight=1)

# 左一欄：檔案清單（有下拉卷軸）
file_list_outer = tb.Frame(bottom_frame)
file_list_outer.grid(row=0, column=0, sticky="nsew")
file_list_outer.grid_rowconfigure(0, weight=1)
file_list_outer.grid_columnconfigure(0, weight=1)

file_list_canvas = tk.Canvas(file_list_outer, highlightthickness=0)
file_list_canvas.grid(row=0, column=0, sticky="nsew")
file_list_vscroll = tk.Scrollbar(file_list_outer, orient="vertical", command=file_list_canvas.yview)
file_list_vscroll.grid(row=0, column=1, sticky="ns")
file_list_canvas.configure(yscrollcommand=file_list_vscroll.set)

file_list_inner_frame = tb.Frame(file_list_canvas)
file_list_canvas.create_window((0, 0), window=file_list_inner_frame, anchor="nw")

def _on_file_frame_configure(event):
    file_list_canvas.configure(scrollregion=file_list_canvas.bbox("all"))
file_list_inner_frame.bind("<Configure>", _on_file_frame_configure)

def update_file_list():
    for widget in file_list_inner_frame.winfo_children():
        widget.destroy()
    folder = folder_var.get()
    if not os.path.isdir(folder):
        return
    files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    files.sort()
    # 垂直顯示，超過5個用下拉卷軸
    for row, fname in enumerate(files):
        lbl = tb.Label(
            file_list_inner_frame,
            text=fname,
            width=22,
            anchor="w",
            background="#333333",
            foreground="#ffffff",
            font=default_font
        )
        lbl.grid(row=row, column=0, padx=2, pady=1, sticky="ew")
        # 拖移功能
        lbl.bind("<ButtonPress-1>", lambda e, t=fname: on_label_drag_start(e, t))
    file_list_inner_frame.update_idletasks()
    # 自動隱藏垂直卷軸
    if len(files) > 5:
        file_list_vscroll.grid()
    else:
        file_list_vscroll.grid_remove()
    file_list_canvas.config(scrollregion=file_list_canvas.bbox("all"), height=120)

# 右三欄：捕捉視窗列表
window_list_frames = []
for i in range(3):
    frame = tb.Frame(bottom_frame)
    frame.grid(row=0, column=i+1, sticky="nsew")
    window_list_frames.append(frame)

def get_taskbar_window_titles():
    titles = []
    hidden_keywords = [
        "設定", "windows 輸入體驗", "windows input experience", "searchui", "cortana", "lockapp", "program manager"
    ]
    def enum_handler(hwnd, _):
        if (
            win32gui.IsWindowVisible(hwnd)
            and win32gui.GetWindowText(hwnd)
            and not win32gui.IsIconic(hwnd)
            and win32gui.GetParent(hwnd) == 0
            and win32gui.GetWindow(hwnd, win32con.GW_OWNER) == 0
        ):
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            if ex_style & win32con.WS_EX_TOOLWINDOW:
                return
            t = win32gui.GetWindowText(hwnd)
            t_lower = t.strip().lower()
            if not t_lower:
                return
            for kw in hidden_keywords:
                if kw in t_lower:
                    return
            titles.append(t)
    win32gui.EnumWindows(enum_handler, None)
    return titles

def update_window_list():
    titles = get_taskbar_window_titles()
    max_rows = 5
    max_cols = 3
    # 清空
    for frame in window_list_frames:
        for widget in frame.winfo_children():
            widget.destroy()
    # 填入
    for col in range(max_cols):
        for row in range(max_rows):
            idx = col * max_rows + row
            if idx < len(titles):
                title = titles[idx]
                lbl = tb.Label(window_list_frames[col], text=title, width=22, anchor="w", font=default_font)
                lbl.grid(row=row, column=0, padx=2, pady=1, sticky="ew")
                # 拖移功能
                lbl.bind("<ButtonPress-1>", lambda e, t=title: on_label_drag_start(e, t))

# --- 拖移功能（確保這段有在你的主程式） ---
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
    label = tb.Label(drag_label_popup["win"], text=title, background="#222", foreground="#fff", font=default_font)
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

# --- 初始化 ---
update_file_list()
update_window_list()

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
    import ctypes
    files = get_group_files(group_code)
    if not files:
        log(f"分組 {group_display_names[group_code].get()} 沒有檔案")
        return
    # 取得分組視窗標題（小寫）
    target_titles = [os.path.splitext(os.path.basename(f))[0].lower() for f in files if f]
    hwnd_list = []
    other_hwnd_list = []
    all_group_titles = set()
    # 收集所有分組的視窗標題
    for c in group_codes:
        if c == group_code:
            continue
        for f in get_group_files(c):
            all_group_titles.add(os.path.splitext(os.path.basename(f))[0].lower())

    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            window_text_lower = window_text.lower()
            if any(title and (window_text_lower == title or title in window_text_lower) for title in target_titles):
                hwnd_list.append(hwnd)
            elif any(title and (window_text_lower == title or title in window_text_lower) for title in all_group_titles):
                other_hwnd_list.append(hwnd)
            elif window_text.strip():
                # 沒有分組的視窗也往下層退
                other_hwnd_list.append(hwnd)
    win32gui.EnumWindows(enum_handler, None)

    # 先將所有非分組視窗往下層退
    for hwnd in other_hwnd_list:
        try:
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_BOTTOM, 0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
            )
        except Exception:
            pass

    # 再將分組視窗全部浮到最上層，並還原被縮小的
    found_any = False
    for hwnd in hwnd_list:
        try:
            # 還原被縮小的視窗
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            try:
                ctypes.windll.user32.AllowSetForegroundWindow(-1)
            except Exception:
                pass
            win32gui.SetForegroundWindow(hwnd)
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
            )
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
            )
            found_any = True
        except Exception as e:
            log(f"bring to front 失敗: {win32gui.GetWindowText(hwnd)} ({e})")
    if found_any:
        log(f"已將分組 {group_display_names[group_code].get()} 的視窗全部浮現到最前方")
    else:
        log(f"找不到符合的視窗，已將其他視窗往下層退。")

def register_global_hotkeys():
    for idx, code in enumerate(group_codes):
        hotkey = group_hotkeys[idx].get()
        try:
            keyboard.remove_hotkey(f"group_{code}")
        except Exception:
            pass
        keyboard.add_hotkey(
            hotkey,
            functools.partial(set_group_windows_topmost, code),
            suppress=False,
            trigger_on_release=False
        )

for idx, var in enumerate(group_hotkeys):
    var.trace_add("write", lambda *a, i=idx: register_global_hotkeys())

register_global_hotkeys()

app.mainloop()