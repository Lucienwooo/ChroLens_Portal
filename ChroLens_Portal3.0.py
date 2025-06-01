### ChroLens_Portal 2.0 
### 2025/05/26 By Lucienwooo
### pyinstaller --onedir --noconsole --add-data "冥想貓貓.ico;." --icon=冥想貓貓.ico --hidden-import=win32timezone ChroLens_Portal2.0.py
###### 分組視窗透過快捷鍵最上層顯示，半成品。
# 目前頂層顯示，某些視窗還是無法正常顯示
# 檔案只能開啟當前路徑，即便先前在別的路徑取得檔案名稱到分組
# 仍然會只能開啟當前資料夾
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
import keyboard
from tkinter import font as tkfont
import functools
import atexit

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

def open_files(folder_path, file_names=None, interval=4, log_func=None):
    if file_names is None:
        files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    else:
        files = [f for f in file_names if f]
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
        update_file_list()
        update_window_list()  # 新增：刷新右側視窗清單

def show_files_in_folder(folder):
    pass

def show_log_window():
    if hasattr(app, 'log_win') and app.log_win.winfo_exists():
        app.log_win.deiconify()
        return
    log_win = tk.Toplevel(app)
    log_win.title("動態紀錄")
    log_win.geometry(f"400x600+{app.winfo_rootx()-400}+{app.winfo_rooty()}")
    log_win.resizable(True, True)
    log_win.attributes('-topmost', True)
    try:
        ico_path = resource_path("冥想貓貓.ico")
        log_win.iconbitmap(ico_path)
    except Exception as e:
        print(f"無法設定 log 視窗 icon: {e}")
    app.log_win = log_win

    def start_move(event):
        log_win._drag_start_x = event.x
        log_win._drag_start_y = event.y
    def do_move(event):
        x = log_win.winfo_x() + event.x - log_win._drag_start_x
        y = log_win.winfo_y() + event.y - log_win._drag_start_y
        log_win.geometry(f"+{x}+{y}")
    log_win.bind("<Button-1>", start_move)
    log_win.bind("<B1-Motion>", do_move)

    log_text = tb.Text(log_win, height=30, width=60, state="normal")
    log_text.pack(fill="both", expand=True)
    app.log_text = log_text

    # 顯示所有歷史紀錄
    log_text.config(state="normal")
    log_text.delete("1.0", "end")
    for line in log_history:
        log_text.insert("end", line + "\n")
    log_text.see("end")
    log_text.config(state="disabled")

def toggle_log():
    if hasattr(app, 'log_win') and app.log_win.winfo_exists():
        if app.log_win.state() == 'normal':
            app.log_win.withdraw()
        else:
            app.log_win.deiconify()
    else:
        show_log_window()

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

def open_files(folder_path, file_names=None, interval=4, log_func=None):
    if file_names is None:
        files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    else:
        files = [f for f in file_names if f]
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
        update_file_list()
        update_window_list()  # 新增：刷新右側視窗清單

def show_files_in_folder(folder):
    pass

log_history = []

def log(msg):
    timestamp = time.strftime("%H:%M:%S")
    full_msg = f"[{timestamp}] {msg}"
    log_history.append(full_msg)
    # 若日誌視窗已存在，則即時顯示
    if hasattr(app, 'log_text') and app.log_text.winfo_exists():
        app.log_text.config(state="normal")
        app.log_text.insert("end", full_msg + "\n")
        app.log_text.see("end")
        app.log_text.config(state="disabled")

def show_log_window():
    if hasattr(app, 'log_win') and app.log_win.winfo_exists():
        app.log_win.deiconify()
        return
    log_win = tk.Toplevel(app)
    log_win.title("動態紀錄")
    log_win.geometry(f"400x600+{app.winfo_rootx()-400}+{app.winfo_rooty()}")
    log_win.resizable(True, True)
    log_win.attributes('-topmost', True)
    # 設定 icon
    try:
        ico_path = resource_path("冥想貓貓.ico")
        log_win.iconbitmap(ico_path)
    except Exception as e:
        print(f"無法設定 log 視窗 icon: {e}")
    app.log_win = log_win

    def start_move(event):
        log_win._drag_start_x = event.x
        log_win._drag_start_y = event.y
    def do_move(event):
        x = log_win.winfo_x() + event.x - log_win._drag_start_x
        y = log_win.winfo_y() + event.y - log_win._drag_start_y
        log_win.geometry(f"+{x}+{y}")
    log_win.bind("<Button-1>", start_move)
    log_win.bind("<B1-Motion>", do_move)

    log_text = tb.Text(log_win, height=30, width=60, state="normal")
    log_text.pack(fill="both", expand=True)
    app.log_text = log_text

    # 顯示所有歷史紀錄
    for line in log_history:
        log_text.insert("end", line + "\n")
    log_text.see("end")
    log_text.config(state="disabled")

def toggle_log():
    if hasattr(app, 'log_win') and app.log_win.winfo_exists():
        if app.log_win.state() == 'normal':
            app.log_win.withdraw()
        else:
            app.log_win.deiconify()
    else:
        show_log_window()

log_history = []

def log(msg):
    timestamp = time.strftime("%H:%M:%S")
    full_msg = f"[{timestamp}] {msg}"
    log_history.append(full_msg)
    # 無論日誌視窗是否顯示，都更新內容
    if 'log_text' in globals() and log_text.winfo_exists():
        log_text.config(state="normal")
        log_text.insert("end", full_msg + "\n")
        log_text.see("end")
        log_text.config(state="disabled")

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

# === 介面區塊 ===
app = tb.Window(themename="darkly")
app.title("ChroLens_Portal 2.0.0")
try:
    ico_path = resource_path("冥想貓貓.ico")
    app.iconbitmap(ico_path)
except Exception as e:
    print(f"無法設定 icon: {e}")

# 統一字體
default_font = tkfont.Font(family="Microsoft JhengHei", size=12)
app.option_add("*Font", default_font)

# === 分組代碼與顯示名稱 ===
group_codes = ["A", "B", "C", "D", "E", "F"]
group_display_names = {c: tk.StringVar(value=c) for c in group_codes}
group_buttons = {}
close_buttons = {}
frm = tb.Frame(app, padding=2)
frm.pack(fill="both", expand=True)

top_row_frame = tb.Frame(frm, padding=2)
top_row_frame.grid(row=0, column=0, columnspan=6, sticky="ew", pady=(2, 2))
top_row_frame.grid_columnconfigure(0, weight=0)
top_row_frame.grid_columnconfigure(1, weight=0)
top_row_frame.grid_columnconfigure(2, weight=0)

folder_var = tb.StringVar(value="")
interval_var = tb.StringVar(value="4")

folder_frame = tb.Frame(top_row_frame, padding=(2,2))
folder_frame.grid(row=0, column=0, sticky="w", padx=(0, 4))
tb.Entry(folder_frame, textvariable=folder_var, width=38).grid(row=0, column=0, padx=(2,2), sticky="ew")
tb.Button(folder_frame, text="選擇開啟路徑", command=choose_folder, bootstyle=SECONDARY).grid(row=0, column=1, padx=(2,0), sticky="ew")

interval_frame = tb.Frame(top_row_frame, padding=(2,2))
interval_frame.grid(row=0, column=1, sticky="w", padx=(0, 4))
tb.Label(interval_frame, text="間隔秒數:").grid(row=0, column=0, sticky="w")
tb.Entry(interval_frame, textvariable=interval_var, width=5).grid(row=0, column=1, padx=(2,0), sticky="w")

# 新增「存檔」按鈕
def manual_save():
    save_settings()
    log("已手動儲存設定檔")
save_btn = tb.Button(top_row_frame, text="存檔", command=manual_save, bootstyle="info")
save_btn.grid(row=0, column=5, padx=(8,2), sticky="e")

# === 新版：快捷鍵設定（僅允許 ALT+任意鍵 或 CTRL+任意鍵）===
default_hotkeys = ["Alt+1", "Alt+2", "Alt+3", "Alt+4", "Alt+5", "Alt+6"]
group_hotkeys = [tk.StringVar(value=default_hotkeys[i]) for i in range(6)]

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
second_row_frame.grid(row=1, column=0, columnspan=8, sticky="ew")
for i in range(7):  # 0~6 共7格
    second_row_frame.grid_columnconfigure(i, weight=1)

show_label_font = tkfont.Font(family="Microsoft JhengHei", size=12)  # 統一字體大小

# 新增最左邊的說明文字
desc_label = tb.Label(second_row_frame, text="置頂顯示切換", width=12, anchor="center", font=show_label_font)
desc_label.grid(row=0, column=0, padx=2, pady=2, sticky="ew")

for idx, code in enumerate(group_codes):
    frame = tb.Frame(second_row_frame, borderwidth=2, relief="groove")
    frame.grid(row=0, column=idx+1, padx=2, pady=2, sticky="ew")
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
def is_entry_focused():
    widget = app.focus_get()
    return isinstance(widget, (tk.Entry, tb.Entry, tk.Text, tb.Combobox))

def global_hotkey_handler(event):
    if is_entry_focused():
        return  # 有輸入框聚焦時不觸發
    pressed = format_hotkey(event)
    for idx, code in enumerate(group_codes):
        if pressed and pressed == group_hotkeys[idx].get():
            set_group_windows_topmost(code)
            log(f"分組 {group_display_names[code].get()} 置頂顯示")
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
    (8, 3, "啟動", "E", "success-outline", lambda: start_group_opening("E")),
    (8, 4, "關閉", "A", "danger-outline", lambda: close_group_windows("A")),
    (8, 5, "關閉", "B", "danger-outline", lambda: close_group_windows("B")),
    (8, 6, "關閉", "E", "danger-outline", lambda: close_group_windows("E")),
    (9, 1, "啟動", "C", "success-outline", lambda: start_group_opening("C")),
    (9, 2, "啟動", "D", "success-outline", lambda: start_group_opening("D")),
    (9, 3, "啟動", "F", "success-outline", lambda: start_group_opening("F")),
    (9, 4, "關閉", "C", "danger-outline", lambda: close_group_windows("C")),
    (9, 5, "關閉", "D", "danger-outline", lambda: close_group_windows("D")),
    (9, 6, "關閉", "F", "danger-outline", lambda: close_group_windows("F")),
]

for row, col, text, code, bootstyle, cmd in group_btn_grid:
    btn = tb.Button(
        btns_outer_frame,
        text=f"{text} {group_display_names[code].get()}",
        bootstyle=bootstyle,
        command=cmd,
        width=8
    )
    btn.grid(row=row, column=col, padx=(4, 4), pady=(2, 2), sticky="nsew")
    if text == "啟動":
        group_buttons[code] = btn
    else:
        close_buttons[code] = btn

# 動態紀錄按鈕下方顯示「動態」「紀錄」兩行
btns_outer_frame.grid_rowconfigure(8, weight=1)
btns_outer_frame.grid_rowconfigure(9, weight=1)

# --- 下方清單區塊（左：檔案名稱，右：視窗名稱） ---
bottom_frame = tb.Frame(frm)
bottom_frame.grid(row=10, column=0, columnspan=6, sticky="ew", pady=(8, 2))
bottom_frame.grid_columnconfigure(0, weight=0)  # 檔案列表固定寬度
bottom_frame.grid_columnconfigure(1, weight=1)  # 視窗列表自動撐滿
bottom_frame.grid_rowconfigure(0, weight=1)

# 左：檔案名稱列表（寬度與動態紀錄按鈕一致，建議 width=90，可依需求微調）
file_list_outer = tb.Frame(bottom_frame, width=290)
file_list_outer.grid(row=0, column=0, sticky="nsw")
file_list_outer.grid_propagate(False)  # 固定寬度不被內容撐開
file_list_outer.grid_rowconfigure(0, weight=1)
file_list_outer.grid_columnconfigure(0, weight=1)

file_list_canvas = tk.Canvas(file_list_outer, highlightthickness=0, height=120)
file_list_canvas.grid(row=0, column=0, sticky="nsew")
file_list_vscroll = tk.Scrollbar(file_list_outer, orient="vertical", command=file_list_canvas.yview)
file_list_canvas.configure(yscrollcommand=file_list_vscroll.set)
# 卷軸隱藏，不 grid 也不 pack

file_list_inner_frame = tb.Frame(file_list_canvas)
file_list_canvas.create_window((0, 0), window=file_list_inner_frame, anchor="nw")

def _on_file_frame_configure(event):
    file_list_canvas.configure(scrollregion=file_list_canvas.bbox("all"))
file_list_inner_frame.bind("<Configure>", _on_file_frame_configure)

# 右：視窗名稱列表（自動撐滿剩餘空間）
window_list_outer = tb.Frame(bottom_frame)
window_list_outer.grid(row=0, column=1, sticky="nsew")
window_list_outer.grid_rowconfigure(0, weight=1)
window_list_outer.grid_columnconfigure(0, weight=1)

window_list_canvas = tk.Canvas(window_list_outer, highlightthickness=0)
window_list_canvas.grid(row=0, column=0, sticky="nsew")
window_list_hscroll = tk.Scrollbar(window_list_outer, orient="horizontal", command=window_list_canvas.xview)
window_list_canvas.configure(xscrollcommand=window_list_hscroll.set)
# 卷軸隱藏，不 grid 也不 pack

window_list_inner_frame = tb.Frame(window_list_canvas)
window_list_canvas.create_window((0, 0), window=window_list_inner_frame, anchor="nw")

window_list_frames = []
max_cols = 3
for i in range(max_cols):
    frame = tb.Frame(window_list_inner_frame)
    frame.grid(row=0, column=i, sticky="nsew")
    window_list_frames.append(frame)

def _on_window_frame_configure(event):
    window_list_canvas.configure(scrollregion=window_list_canvas.bbox("all"))
window_list_inner_frame.bind("<Configure>", _on_window_frame_configure)

def get_taskbar_window_titles():
    exclude_titles = ["設定", "windows 輸入體驗", "windows input experience"]  # 可自行擴充
    titles = []
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            t = win32gui.GetWindowText(hwnd)
            t_lower = t.strip().lower()
            if t and all(ex not in t_lower for ex in exclude_titles):
                titles.append(t)
    win32gui.EnumWindows(enum_handler, None)
    return titles

def update_file_list():
    # 取得目前資料夾檔案
    for widget in file_list_inner_frame.winfo_children():
        widget.destroy()
    folder = folder_var.get()
    if not os.path.isdir(folder):
        return
    files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    files.sort()
    for row, fname in enumerate(files):
        lbl = tb.Label(file_list_inner_frame, text=fname, anchor="w", width=15, font=default_font)
        lbl.grid(row=row, column=0, sticky="ew", padx=2, pady=1)
        lbl.bind("<ButtonPress-1>", lambda e, t=fname: on_label_drag_start(e, t))
    file_list_inner_frame.update_idletasks()
    file_list_canvas.config(scrollregion=file_list_canvas.bbox("all"))

def update_window_list():
    for frame in window_list_frames:
        for widget in frame.winfo_children():
            widget.destroy()
    titles = get_taskbar_window_titles()
    max_rows = 5
    max_cols = 3
    for col in range(max_cols):
        for row in range(max_rows):
            idx = col * max_rows + row
            if idx < len(titles):
                title = titles[idx]
                lbl = tb.Label(window_list_frames[col], text=title, anchor="w", width=22, font=default_font)
                lbl.grid(row=row, column=0, sticky="ew", padx=2, pady=1)
                lbl.bind("<ButtonPress-1>", lambda e, t=title: on_label_drag_start(e, t))
    window_list_inner_frame.update_idletasks()
    window_list_canvas.config(scrollregion=window_list_canvas.bbox("all"))

# 啟動時呼叫
update_file_list()
update_window_list()

# 滑鼠橫向滾輪支援
def _on_window_mousewheel(event):
    window_list_canvas.xview_scroll(int(-1*(event.delta/120)), "units")
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

retreated_hwnds = set()

def set_group_windows_topmost(group_code):
    global retreated_hwnds
    files = get_group_files(group_code)
    if not files:
        log(f"分組 {group_display_names[group_code].get()} 沒有檔案")
        return

    target_titles = [os.path.splitext(os.path.basename(f))[0].lower() for f in files if f]
    my_hwnd = app.winfo_id()
    group_hwnds = set()
    parent_hwnds = set()

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
            group_hwnds.add(hwnd)
            parent = win32gui.GetParent(hwnd)
            if parent and parent != 0 and parent != hwnd:
                parent_hwnds.add(parent)
    win32gui.EnumWindows(enum_handler, None)

    other_hwnd_list = []
    def enum_others(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        if hwnd == my_hwnd or hwnd in group_hwnds or hwnd in parent_hwnds:
            return
        for g_hwnd in group_hwnds:
            if win32gui.IsChild(g_hwnd, hwnd):
                return
        other_hwnd_list.append(hwnd)
    win32gui.EnumWindows(enum_others, None)

    # 只對新出現的視窗退後一次
    new_to_retreated = [hwnd for hwnd in other_hwnd_list if hwnd not in retreated_hwnds]
    for hwnd in new_to_retreated:
        try:
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_BOTTOM, 0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
            )
            retreated_hwnds.add(hwnd)
        except Exception:
            pass

    log(f"已將分組 {group_display_names[group_code].get()} 以外的視窗全部退到最下層（本次退後 {len(new_to_retreated)} 個）")

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
    def open_files():
        for file in files:
            file_path = os.path.join(folder, file)
            if not os.path.exists(file_path):
                log(f"找不到檔案: {file_path}")
                continue
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
                else:
                    log(f"Opening: {file_path}")
                    os.startfile(file_path)
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

    log(f"【DEBUG】關閉時關鍵字: {keywords}")

    closed_any = False

    # 新增：列出所有視窗標題
    all_titles = []
    def collect_titles(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            t = win32gui.GetWindowText(hwnd)
            if t:
                all_titles.append(t)
    win32gui.EnumWindows(collect_titles, None)
    log(f"【DEBUG】目前所有視窗標題: {all_titles}")

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
            "group_var1": [var1.get() for _, var1, _, _, _ in checkbox_vars_entries],
            "group_var2": [var2.get() for _, _, var2, _, _ in checkbox_vars_entries],
        }
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        # 程式結束時 Tk 物件已銷毀，這裡忽略錯誤即可
        pass

def load_settings():
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
        group_var1s = data.get("group_var1", [])
        group_var2s = data.get("group_var2", [])
        for i, v in enumerate(group_var1s):
            if i < len(checkbox_vars_entries):
                checkbox_vars_entries[i][1].set(v)
        for i, v in enumerate(group_var2s):
            if i < len(checkbox_vars_entries):
                checkbox_vars_entries[i][2].set(v)
        # --- 新增：強制同步 UI 顯示 ---
        update_show_labels()
        update_group_name()
    except Exception as e:
        log(f"設定檔讀取失敗: {e}")
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

def _on_mousewheel(event):
    file_list_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
file_list_canvas.bind_all("<MouseWheel>", _on_mousewheel)

def _on_file_mousewheel(event):
    file_list_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
file_list_canvas.bind("<Enter>", lambda e: file_list_canvas.bind_all("<MouseWheel>", _on_file_mousewheel))
file_list_canvas.bind("<Leave>", lambda e: file_list_canvas.unbind_all("<MouseWheel>"))

def _on_window_mousewheel(event):
    window_list_canvas.xview_scroll(int(-1*(event.delta/120)), "units")
window_list_canvas.bind("<Enter>", lambda e: window_list_canvas.bind_all("<MouseWheel>", _on_window_mousewheel))
window_list_canvas.bind("<Leave>", lambda e: window_list_canvas.unbind_all("<MouseWheel>"))

# --- 啟動時延遲0.5秒再讀取設定檔 ---
def delayed_load_settings():
    time.sleep(0.5)
    app.after(0, lambda: [load_settings(), update_file_list(), update_window_list()])
threading.Thread(target=delayed_load_settings, daemon=True).start()

def auto_refresh_window_list():
    # 若主視窗被縮小（iconic），則不刷新
    if app.state() != "iconic":
        update_window_list()
    # 每5秒再呼叫自己
    app.after(5000, auto_refresh_window_list)

# 啟動時呼叫一次
auto_refresh_window_list()

app.mainloop()