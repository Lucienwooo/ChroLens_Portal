### ChroLens_Portal 1.0.0 
### 2025/05/26 By Lucienwooo
# pyinstaller --onefile --noconsole --add-data "冥想貓貓.ico;." --icon=冥想貓貓.ico --hidden-import=win32timezone ChroLens_Portal.py 
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

def log(msg):
    log_text.configure(state="normal")
    log_text.insert("end", msg + "\n")
    log_text.see("end")
    log_text.configure(state="disabled")

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

def select_all_changed():
    if select_all_var.get():
        for var, entry, *_ in checkbox_vars_entries:
            if entry.get():
                var.set(True)
    else:
        for var, entry, *_ in checkbox_vars_entries:
            if entry.get():
                var.set(False)

app = tb.Window(themename="darkly")
app.title("ChroLens_Portal 1.0.0")
try:
    ico_path = resource_path("冥想貓貓.ico")
    app.iconbitmap(ico_path)
except Exception as e:
    print(f"無法設定 icon: {e}")
# === 介面區塊 ===

folder_var = tb.StringVar(value=load_last_path())
interval_var = tb.StringVar(value="4")
select_all_var = tk.BooleanVar(value=True)

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

selectall_frame = tb.Frame(top_row_frame, padding=(2,2))
selectall_frame.grid(row=0, column=2, sticky="e")

tb.Checkbutton(selectall_frame, variable=select_all_var, command=select_all_changed).grid(row=0, column=0, sticky="e")
tb.Label(selectall_frame, text="全選").grid(row=0, column=1, sticky="w", padx=(2,0))

group_codes = ["A", "B", "C", "D"]
group_display_names = {code: tk.StringVar(value=code) for code in group_codes}
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
    selectall_frame, textvariable=tk.StringVar(),
    values=[group_display_names[c].get() for c in group_codes], width=4, state="readonly"
)
group_combo.grid(row=0, column=2, padx=(8,2))

def on_group_combo_selected(event=None):
    display_name = group_combo.get()
    code = get_group_code_by_display_name(display_name)
    group_select_code.set(code)
    group_combo.set(group_display_names[code].get())

group_combo.bind("<<ComboboxSelected>>", on_group_combo_selected)

group_rename_var = tk.StringVar()
group_entry = tb.Entry(selectall_frame, textvariable=group_rename_var, width=14)
group_entry.grid(row=0, column=3, padx=(2,0))
group_entry.bind("<FocusIn>", clear_group_entry_placeholder)
group_entry.bind("<FocusOut>", on_group_entry_focus_out)
group_entry.bind("<Return>", lambda e: update_group_name())
set_group_entry_placeholder()

group_frames = []
for col in range(3):
    group_frame = tb.Frame(frm, borderwidth=1, relief="solid", padding=2)
    group_frame.grid(row=2, column=col, padx=2, pady=2, sticky="n")
    group_frames.append(group_frame)

checkbox_vars_entries = []
for i in range(15):
    var = tk.BooleanVar(value=False)
    entry = tb.Entry(group_frames[i // 5], width=11, state="readonly")
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
    tb.Label(group_frames[i // 5], text=str(i+1), width=2).grid(row=row, column=0, sticky="w", padx=(0,1))
    chk = tb.Checkbutton(group_frames[i // 5], variable=var, width=2)
    chk.grid(row=row, column=1, sticky="w", padx=(0,1), pady=1)
    entry.grid(row=row, column=2, padx=1, pady=1)
    group_combo1.grid(row=row, column=3, padx=1, pady=1)
    group_combo2.grid(row=row, column=4, padx=1, pady=1)
    tb.Label(group_frames[i // 5], text="組", width=2).grid(row=row, column=5, sticky="w")
    checkbox_vars_entries.append((var, entry, group_var1, group_var2, group_combo1, group_combo2))

def get_group_files(group_code):
    files = []
    for var, entry, group_var1, group_var2, *_ in checkbox_vars_entries:
        code1 = get_group_code_by_display_name(group_var1.get())
        code2 = get_group_code_by_display_name(group_var2.get())
        filename = entry.get()
        if filename and (group_code == code1 or group_code == code2):
            files.append(filename)
    return files

from tkinter import font as tkfont
big_btn_font = tkfont.Font(family="Microsoft JhengHei", size=16, weight="bold")

group_labels = group_codes

group_buttons = {}
close_buttons = {}

btns_outer_frame = tb.Frame(frm)
btns_outer_frame.grid(row=8, column=0, columnspan=6, sticky="ew", padx=(0, 8), pady=(8, 4))
btns_outer_frame.columnconfigure(0, weight=1)
btns_outer_frame.columnconfigure(1, weight=1)

open_btns_frame = tb.Frame(btns_outer_frame, borderwidth=1, relief="solid", padding=4)
open_btns_frame.grid(row=0, column=0, sticky="ew")
open_btns_frame.grid_columnconfigure((0,1), weight=1)

close_btns_frame = tb.Frame(btns_outer_frame, borderwidth=1, relief="solid", padding=4)
close_btns_frame.grid(row=0, column=1, sticky="ew")
close_btns_frame.grid_columnconfigure((0,1), weight=1)

for idx, code in enumerate(group_labels):
    row = 0 if idx < 2 else 1
    col = idx % 2
    btn = tb.Button(
        open_btns_frame,
        text=f"啟動 {group_display_names[code].get()}",
        bootstyle="success-outline",
        command=lambda c=code: start_group_opening(c),
        width=8
    )
    btn.grid(row=row, column=col, padx=(8, 8), pady=(4, 4), sticky="ew")
    group_buttons[code] = btn

for idx, code in enumerate(group_labels):
    row = 0 if idx < 2 else 1
    col = idx % 2
    close_btn = tb.Button(
        close_btns_frame,
        text=f"關閉 {group_display_names[code].get()}",
        bootstyle="danger-outline",
        command=lambda c=code: close_group_windows(c),
        width=8
    )
    close_btn.grid(row=row, column=col, padx=(8, 8), pady=(4, 4), sticky="ew")
    close_buttons[code] = close_btn

log_text = tb.Text(frm, height=8, width=90, state="disabled")
log_text.grid(row=10, column=0, columnspan=6, pady=2, sticky="ew")
frm.grid_columnconfigure(0, weight=1)
frm.grid_columnconfigure(1, weight=1)
frm.grid_columnconfigure(2, weight=1)
frm.grid_columnconfigure(3, weight=1)
frm.grid_columnconfigure(4, weight=1)
frm.grid_columnconfigure(5, weight=1)
frm.grid_columnconfigure(6, weight=1)
frm.grid_columnconfigure(7, weight=1)

def update_checkboxes(folder):
    files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    files.sort()
    for i in range(15):
        var, entry, group_var1, group_var2, *_ = checkbox_vars_entries[i]
        entry.config(state="normal")
        if i < len(files):
            entry.delete(0, END)
            entry.insert(0, files[i])
            var.set(True)
        else:
            entry.delete(0, END)
            var.set(False)
            group_var1.set("")
            group_var2.set("")
        entry.config(state="readonly")

def on_folder_var_change(*args):
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
                "filename": checkbox_vars_entries[i][1].get(),
                "group1": checkbox_vars_entries[i][2].get(),
                "group2": checkbox_vars_entries[i][3].get(),
                "checked": checkbox_vars_entries[i][0].get()
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
        return
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        for c in group_codes:
            if c in data.get("group_display_names", {}):
                group_display_names[c].set(data["group_display_names"][c])
        for c in group_codes:
            group_buttons[c]['text'] = f"啟動 {group_display_names[c].get()}"
            close_buttons[c]['text'] = f"關閉 {group_display_names[c].get()}"
        file_groups = data.get("file_groups", [])
        for i, fg in enumerate(file_groups):
            if i < len(checkbox_vars_entries):
                var, entry, group_var1, group_var2, group_combo1, group_combo2 = checkbox_vars_entries[i]
                entry.config(state="normal")
                entry.delete(0, END)
                entry.insert(0, fg.get("filename", ""))
                var.set(fg.get("checked", False))
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
            _, _, group_var1, group_var2, group_combo1, group_combo2 = checkbox_vars_entries[i]
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
    var, entry, group_var1, group_var2, *_ = checkbox_vars_entries[i]
    group_var1.trace_add("write", auto_save)
    group_var2.trace_add("write", auto_save)
    var.trace_add("write", auto_save)

load_settings()

app.mainloop()