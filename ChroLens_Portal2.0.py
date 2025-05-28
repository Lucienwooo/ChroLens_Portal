### ChroLens_Portal 2.0 
### 2025/05/26 By Lucienwooo
### pyinstaller --onefile --noconsole --add-data "冥想貓貓.ico;." --icon=冥想貓貓.ico --hidden-import=win32timezone ChroLens_Portal.py 
###### 分組視窗透過快捷鍵最上層顯示，半成品。
# 新增清單窗格，顯示當前所有開啟的視窗名稱
# 並且讓選單可以記憶所以從清單選取的檔案/視窗名稱
# 程式將不再只是能開啟或關閉程式
import os
import time
import win32gui
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
second_row_frame.grid_columnconfigure((0,1,2,3), weight=1)  # 讓四個元件平均分配

show_label_font = tkfont.Font(family="Microsoft JhengHei", size=13)  # 放大一級
for idx, code in enumerate(group_codes):
    frame = tb.Frame(second_row_frame, borderwidth=2, relief="groove")  # 加上邊框
    frame.grid(row=0, column=idx, padx=2, pady=2, sticky="ew")  # padx/pady 控制間距
    show_label = tb.Label(frame, text=f"{group_display_names[code].get()} 顯示", width=10)  # width=10 控制標籤寬度
    show_label['font'] = show_label_font  # 設定放大字體
    show_label.pack(side="left")
    hotkey_entry = tb.Entry(frame, textvariable=group_hotkeys[idx], width=8, state="readonly", justify="center")
    hotkey_entry.pack(side="left", padx=(2,5))  # padx 控制熱鍵框與標籤間距
    def make_on_key(idx):
        return lambda event, i=idx: on_hotkey_entry_key(event, i)
    hotkey_entry.bind("<Key>", make_on_key(idx))
    hotkey_entry.bind("<Button-1>", lambda e, entry=hotkey_entry: entry.focus_set())
    show_label_frames.append((show_label, hotkey_entry))

def update_show_labels(*args):
    for idx, code in enumerate(group_codes):
        show_label_frames[idx][0].config(text=f"{group_display_names[code].get()} 顯示")  # 只顯示「A 顯示」

for c in group_codes:
    group_display_names[c].trace_add("write", update_show_labels)

# === 新版：全域快捷鍵觸發（僅 ALT+任意 或 CTRL+任意）===
def global_hotkey_handler(event):
    pressed = format_hotkey(event)
    for idx, code in enumerate(group_codes):
        if pressed and pressed == group_hotkeys[idx].get():
            set_group_windows_topmost(code)
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
    for i in range(15):
        combo1 = checkbox_vars_entries[i][4]
        combo2 = checkbox_vars_entries[i][5]
        combo1['values'] = [""] + [group_display_names[c].get() for c in group_codes]
        combo2['values'] = [""] + [group_display_names[c].get() for c in group_codes]
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

# --- 清單列表區塊（取代原本日誌顯示框） ---
bottom_frame = tb.Frame(frm)
bottom_frame.grid(row=10, column=0, columnspan=6, sticky="ew", pady=(8, 2))
bottom_frame.grid_columnconfigure(0, weight=1)
bottom_frame.grid_rowconfigure(0, weight=1)
bottom_frame.grid_rowconfigure(1, weight=0)

# --- 日誌顯示框（預設隱藏，顯示時高度與主程式同高、寬度為主程式1/2） ---
log_text = tb.Text(frm, height=30, width=60, state="disabled")

def update_checkboxes(folder):
    files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    files.sort()
    for i in range(15):
        entry, group_var1, group_var2, *_ = checkbox_vars_entries[i]
        entry.config(state="normal")
        if i < len(files):
            entry.delete(0, END)
            entry.insert(0, files[i])
        else:
            entry.delete(0, END)
            group_var1.set("")
            group_var2.set("")
        entry.config(state="readonly")

def on_folder_var_change(*args):
    # 僅在沒有設定檔時才初始化
    if not os.path.exists(SETTINGS_FILE):
        show_files_in_folder(folder_var.get())
        update_checkboxes(folder_var.get())

folder_var.trace_add("write", on_folder_var_change)

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
    display_name = group_display_names[group_code].get()
    if not files:
        log(f"分組 {display_name} 沒有檔案")
        return
    log(f"開始開啟分組 {display_name} 的檔案，每 {interval} 秒一個")
    threading.Thread(target=open_files_by_names, args=(folder, files, interval, log), daemon=True).start()

def close_group_windows(group_code):
    folder = folder_var.get()
    files = get_group_files(group_code)
    display_name = group_display_names[group_code].get()
    if not files:
        log(f"分組 {display_name} 沒有檔案可關閉")
        return
    window_titles = [os.path.splitext(os.path.basename(f))[0] for f in files]
    close_windows_by_titles(window_titles)
    log(f"已嘗試關閉分組 {display_name} 的視窗（標題：{', '.join(window_titles)}）")

def close_windows_by_titles(window_titles):
    def enum_handler(hwnd, titles):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            for title in titles:
                if title and title.strip() in window_text:
                    try:
                        win32gui.PostMessage(hwnd, 0x0010, 0, 0)
                    except Exception:
                        pass

    win32gui.EnumWindows(enum_handler, window_titles)

def save_settings():
    data = {
        "group_display_names": {c: group_display_names[c].get() for c in group_codes},
        "file_groups": [
            {
                "filename": checkbox_vars_entries[i][0].get(),  # entry
                "group1": checkbox_vars_entries[i][1].get(),    # group_var1
                "group2": checkbox_vars_entries[i][2].get()     # group_var2
            }
            for i in range(len(checkbox_vars_entries))
        ],
        "folder": folder_var.get(),
        "interval": interval_var.get()
    }
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        update_checkboxes(folder_var.get())  # 只在沒設定檔時初始化
        return
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        for c in group_codes:
            if c in data.get("group_display_names", {}):
                group_display_names[c].set(data["group_display_names"][c])
        for c in group_codes:
            group_buttons[c]['text'] = f"啟動 {group_display_names[c].get()}"  # 修正為啟動
            close_buttons[c]['text'] = f"關閉 {group_display_names[c].get()}"
        file_groups = data.get("file_groups", [])
        for i, fg in enumerate(file_groups):
            if i < len(checkbox_vars_entries):
                entry, group_var1, group_var2, group_combo1, group_combo2 = checkbox_vars_entries[i]
                entry.config(state="normal")
                entry.delete(0, END)
                entry.insert(0, fg.get("filename", ""))
                group_combo1['values'] = [""] + [group_display_names[c].get() for c in group_codes]
                group_combo2['values'] = [""] + [group_display_names[c].get() for c in group_codes]
                group_var1.set(fg.get("group1", ""))
                group_var2.set(fg.get("group2", ""))
                entry.config(state="readonly")
        if "folder" in data:
            folder_var.set(data["folder"])
        if "interval" in data:
            interval_var.set(data["interval"])
        for i in range(len(checkbox_vars_entries)):
            entry, group_var1, group_var2, group_combo1, group_combo2 = checkbox_vars_entries[i]
            group_combo1['values'] = [""] + [group_display_names[c].get() for c in group_codes]
            group_combo2['values'] = [""] + [group_display_names[c].get() for c in group_codes]
            group_combo1.set(group_var1.get())
            group_combo2.set(group_var2.get())
        group_combo['values'] = [group_display_names[c].get() for c in group_codes]
        group_combo.set(group_display_names[group_codes[0]].get())
    except Exception as e:
        log(f"載入設定失敗: {e}")

def auto_save(*args, **kwargs):
    save_settings()

for c in group_codes:
    group_display_names[c].trace_add("write", auto_save)
folder_var.trace_add("write", auto_save)
interval_var.trace_add("write", auto_save)
for i in range(len(checkbox_vars_entries)):
    entry, group_var1, group_var2, group_combo1, group_combo2 = checkbox_vars_entries[i]
    group_var1.trace_add("write", auto_save)
    group_var2.trace_add("write", auto_save)

load_settings()
save_btn = tb.Button(top_row_frame, text="存檔", width=4, command=save_settings, bootstyle=SUCCESS)  # width=4 控制按鈕寬度
save_btn.grid(row=0, column=5, padx=(4,2), sticky="w")  # column=5 放在最右側，padx 控制與左側間距

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
    

    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            window_text_lower = window_text.lower()
            if any(title and (window_text_lower == title or title in window_text_lower) for title in target_titles):
                hwnd_list.append(hwnd)
            elif window_text.strip():
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

    # 再將分組視窗全部浮到最上層
    found_any = False
    for hwnd in hwnd_list:
        try:
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
        log(f"找不到符合的視窗。")

def bind_hotkey(idx, code):
    def on_key_press(event):
        if format_hotkey(event) == group_hotkeys[idx].get():
            set_group_windows_topmost(code)
    app.bind_all("<Key>", on_key_press, add="+")

for idx, code in enumerate(group_codes):
    bind_hotkey(idx, code)

def format_keyboard_hotkey(hotkey_str):
    # 轉換顯示用的 "Alt+1" 為 keyboard 用的 "alt+1"
    return hotkey_str.lower().replace("ctrl", "control")

def register_global_hotkeys():
    # 先清除舊的
    try:
        keyboard.remove_all_hotkeys()
    except Exception:
        pass
    for idx, code in enumerate(group_codes):
        hotkey = group_hotkeys[idx].get()
        if hotkey:
            keyboard.add_hotkey(
                format_keyboard_hotkey(hotkey),
                lambda c=code: set_group_windows_topmost(c)
            )

# 每次快捷鍵設定變動時都重新註冊
for idx in range(4):
    group_hotkeys[idx].trace_add("write", lambda *a, **k: register_global_hotkeys())

# 程式啟動時註冊一次
register_global_hotkeys()

import win32con

def get_taskbar_window_titles():
    titles = []
    def enum_handler(hwnd, _):
        if (
            win32gui.IsWindowVisible(hwnd)
            and win32gui.GetWindowText(hwnd)
            and not win32gui.IsIconic(hwnd)
            and win32gui.GetParent(hwnd) == 0
            and win32gui.GetWindow(hwnd, win32con.GW_OWNER) == 0
        ):
            t = win32gui.GetWindowText(hwnd)
            if t.strip():
                titles.append(t)
    win32gui.EnumWindows(enum_handler, None)
    return titles

# 拖曳功能全域變數
dragged_window_title = {"title": None}
drag_label_popup = {"win": None}

def on_label_drag_start(event, title):
    dragged_window_title["title"] = title
    # 顯示懸浮視窗
    if drag_label_popup["win"] is not None:
        drag_label_popup["win"].destroy()
    drag_win = tk.Toplevel(app)
    drag_win.overrideredirect(True)
    drag_win.attributes("-topmost", True)
    drag_win.geometry(f"+{event.x_root+10}+{event.y_root+10}")
    lbl = tk.Label(drag_win, text=title, font=list_font, bg="#ffe066", relief="solid", bd=1)
    lbl.pack()
    drag_label_popup["win"] = drag_win
    # 跟隨滑鼠移動
    def follow_mouse(ev):
        drag_win.geometry(f"+{ev.x_root+10}+{ev.y_root+10}")
    app.bind("<Motion>", follow_mouse)
    # 滑鼠放開時自動關閉
    def on_release(ev):
        if drag_label_popup["win"]:
            drag_label_popup["win"].destroy()
            drag_label_popup["win"] = None
        app.unbind("<Motion>")
        app.unbind("<ButtonRelease-1>")
    app.bind("<ButtonRelease-1>", on_release)

def on_entry_drag_enter(event, entry):
    # 拖曳進入時自動寫入
    if dragged_window_title["title"]:
        entry.config(state="normal", background="#ffe066")
        entry.delete(0, tk.END)
        entry.insert(0, dragged_window_title["title"])
        entry.config(state="readonly", background="white")
        dragged_window_title["title"] = None
        # 關閉懸浮視窗
        if drag_label_popup["win"]:
            drag_label_popup["win"].destroy()
            drag_label_popup["win"] = None
        app.unbind("<Motion>")
        app.unbind("<ButtonRelease-1>")
    else:
        entry.config(background="#ffe066")

def on_entry_drag_leave(event, entry):
    entry.config(background="white")

# --- 捕捉視窗列表區塊 ---
list_font = tkfont.Font(family="Microsoft JhengHei", size=12)
def update_window_list():
    for widget in window_list_inner_frame.winfo_children():
        widget.destroy()
    titles = get_taskbar_window_titles()
    col_count = (len(titles) + 4) // 5
    if col_count < 5:
        col_count = 5
    for col in range(col_count):
        for row in range(5):
            idx = col * 5 + row
            if idx < len(titles):
                title = titles[idx]
                lbl = tb.Label(window_list_inner_frame, text=title, width=22, anchor="w")
                lbl['font'] = list_font
                lbl.grid(row=row, column=col, padx=2, pady=1, sticky="w")
                lbl.bind("<ButtonPress-1>", lambda e, t=title: on_label_drag_start(e, t))
    window_list_inner_frame.update_idletasks()
    width = col_count * 150
    window_list_canvas.config(scrollregion=window_list_canvas.bbox("all"), width=750)
    if col_count > 5:
        window_list_scroll.grid()
    else:
        window_list_scroll.grid_remove()

# 讓每個 entry 支援 drop
for entry, *_ in checkbox_vars_entries:
    entry.bind("<Enter>", lambda e, ent=entry: on_entry_drag_enter(e, ent))
    entry.bind("<Leave>", lambda e, ent=entry: on_entry_drag_leave(e, ent))

# --- 捕捉視窗列表區塊 UI 建立 ---
window_list_canvas = tk.Canvas(bottom_frame, height=120)
window_list_canvas.grid(row=0, column=0, sticky="nsew")
window_list_scroll = tk.Scrollbar(bottom_frame, orient="horizontal", command=window_list_canvas.xview)
window_list_scroll.grid(row=1, column=0, sticky="ew")
window_list_canvas.configure(xscrollcommand=window_list_scroll.set)

window_list_inner_frame = tb.Frame(window_list_canvas)
window_list_canvas.create_window((0, 0), window=window_list_inner_frame, anchor="nw")

def _on_frame_configure(event):
    window_list_canvas.configure(scrollregion=window_list_canvas.bbox("all"))
window_list_inner_frame.bind("<Configure>", _on_frame_configure)


# 初始化時呼叫一次
update_window_list()

def on_app_close():
    save_settings()
    app.destroy()

app.protocol("WM_DELETE_WINDOW", on_app_close)

app.mainloop()