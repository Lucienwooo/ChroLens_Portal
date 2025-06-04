### ChroLens_Portal 2.0 
### 2025/05/26 By Lucienwooo

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

# row 0：頂部工具列（資料夾選擇、間隔秒數、分組名稱編輯、存檔按鈕）
# row 1：分組置頂顯示切換區（顯示分組名稱與快捷鍵）
# row 2：分組檔案列（15組分組欄位，分3欄顯示，每行3個分組下拉選單）
# row 8：動態日誌、啟動按鈕..、關閉按鈕..
# row 9：動態日誌、啟動按鈕..、關閉按鈕..
# row 10：動態日誌、檔案名稱列表、視窗名稱列表

LAST_PATH_FILE = "last_path.txt"
SETTINGS_FILE = "chrolens_portal.json"

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# === 介面區塊 ===
app = tb.Window(themename="darkly")
app.title("ChroLens_Portal 2.0.0")
try:
    ico_path = resource_path("冥想貓貓.ico")
    app.iconbitmap(ico_path)
except Exception as e:
    print(f"無法設定 icon: {e}")

# --- 主 Frame ---
frm = tb.Frame(app, padding=2)
frm.pack(fill="both", expand=True)

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
        update_file_list()  # 新增：選擇資料夾後即時刷新檔案列表

folder_frame = tb.Frame(top_row_frame, padding=(2,2))
folder_frame.grid(row=0, column=0, sticky="w", padx=(0, 4))
tb.Entry(folder_frame, textvariable=folder_var, width=38).grid(row=0, column=0, padx=(2,2), sticky="ew")
tb.Button(folder_frame, text="選擇開啟路徑", command=lambda: choose_folder(), bootstyle=SECONDARY).grid(row=0, column=1, padx=(2,0), sticky="ew")

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

# --- 分組與快捷鍵 ---
group_codes = ["A", "B", "C", "D", "E", "F"]
group_display_names = {c: tk.StringVar(value=c) for c in group_codes}
default_hotkeys = ["Alt+1", "Alt+2", "Alt+3", "Alt+4", "Alt+Q", "Alt+W"]
group_hotkeys = [tk.StringVar(value=default_hotkeys[i]) for i in range(6)]
group_buttons = {}
close_buttons = {}

def format_hotkey(event):
    keys = []
    if event.state & 0x0004:  # Ctrl
        keys.append("Ctrl")
    if event.state & 0x0008:  # Alt
        keys.append("Alt")
    if not keys:
        return ""
    key = event.keysym
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

# --- row 1：分組置頂顯示切換區 ---
show_label_frames = []
second_row_frame = tb.Frame(frm)
second_row_frame.grid(row=1, column=0, columnspan=8, sticky="ew")
for i in range(7):
    second_row_frame.grid_columnconfigure(i, weight=1)
show_label_font = tkfont.Font(family="Microsoft JhengHei", size=12)
desc_label = tb.Label(second_row_frame, text="置頂切換", width=12, anchor="center", font=show_label_font)
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

# --- row 2：分組檔案列 ---
group_frames = []
for col in range(3):
    group_frame = tb.Frame(frm, borderwidth=1, relief="solid", padding=2)
    group_frame.grid(row=2, column=col, padx=2, pady=2, sticky="nsew")
    frm.grid_columnconfigure(col, weight=1)
    group_frames.append(group_frame)
num_font = tkfont.Font(family="Microsoft JhengHei", size=12, weight="bold")
checkbox_vars_entries = []
for i in range(12):  # 12行
    entry = tb.Entry(group_frames[i // 4], width=14, state="readonly")
    group_var1 = tk.StringVar(value="")
    group_var2 = tk.StringVar(value="")
    group_var3 = tk.StringVar(value="")
    group_var4 = tk.StringVar(value="")
    group_combo1 = tb.Combobox(
        group_frames[i // 4], textvariable=group_var1,
        values=[""] + [group_display_names[c].get() for c in group_codes], width=4, state="readonly"
    )
    group_combo2 = tb.Combobox(
        group_frames[i // 4], textvariable=group_var2,
        values=[""] + [group_display_names[c].get() for c in group_codes], width=4, state="readonly"
    )
    group_combo3 = tb.Combobox(
        group_frames[i // 4], textvariable=group_var3,
        values=[""] + [group_display_names[c].get() for c in group_codes], width=4, state="readonly"
    )
    group_combo4 = tb.Combobox(
        group_frames[i // 4], textvariable=group_var4,
        values=[""] + [group_display_names[c].get() for c in group_codes], width=4, state="readonly"
    )
    row = i % 4
    num_label = tb.Label(group_frames[i // 4], text=str(i+1), width=2)
    num_label['font'] = num_font
    num_label.grid(row=row, column=0, sticky="w", padx=0)
    entry.grid(row=row, column=1, padx=0, pady=1)
    group_combo1.grid(row=row, column=2, padx=0, pady=1)
    group_combo2.grid(row=row, column=3, padx=0, pady=1)
    group_combo3.grid(row=row, column=4, padx=0, pady=1)
    group_combo4.grid(row=row, column=5, padx=0, pady=1)
    checkbox_vars_entries.append((entry, group_var1, group_var2, group_var3, group_var4, group_combo1, group_combo2, group_combo3, group_combo4))

def get_group_code_by_display_name(display_name):
    for code in group_codes:
        if group_display_names[code].get() == display_name:
            return code
    return None

def get_group_files(group_code):
    files = []
    for entry, group_var1, group_var2, group_var3, *_ in checkbox_vars_entries:
        code1 = get_group_code_by_display_name(group_var1.get())
        code2 = get_group_code_by_display_name(group_var2.get())
        code3 = get_group_code_by_display_name(group_var3.get())
        filename = entry.get()
        if filename and (group_code == code1 or group_code == code2 or group_code == code3):
            files.append(filename)
    return files

# --- row 8~10 動態日誌區塊 ---
log_text = tb.Text(frm, height=18, width=18, state="disabled", wrap="word", font=tkfont.Font(family="Microsoft JhengHei", size=10))
log_text.grid(row=8, column=0, rowspan=3, sticky="nsew", padx=(0, 8), pady=(0, 0))
frm.grid_rowconfigure(8, weight=1)
frm.grid_rowconfigure(9, weight=1)
frm.grid_rowconfigure(10, weight=1)
frm.grid_columnconfigure(0, weight=0)

# --- row 8~9 啟動/關閉分組按鈕區域 ---
btns_outer_frame = tb.Frame(frm)
btns_outer_frame.grid(row=8, column=1, rowspan=2, columnspan=6, sticky="ew", padx=(0, 4), pady=(8, 4))
for i in range(6):
    btns_outer_frame.grid_columnconfigure(i, weight=1)
group_btn_grid = [
    (8, 0, "啟動", "A", "success-outline", lambda: start_group_opening("A")),
    (8, 1, "啟動", "B", "success-outline", lambda: start_group_opening("B")),
    (8, 2, "啟動", "E", "success-outline", lambda: start_group_opening("E")),
    (8, 3, "關閉", "A", "danger-outline", lambda: close_group_windows("A")),
    (8, 4, "關閉", "B", "danger-outline", lambda: close_group_windows("B")),
    (8, 5, "關閉", "E", "danger-outline", lambda: close_group_windows("E")),
    (9, 0, "啟動", "C", "success-outline", lambda: start_group_opening("C")),
    (9, 1, "啟動", "D", "success-outline", lambda: start_group_opening("D")),
    (9, 2, "啟動", "F", "success-outline", lambda: start_group_opening("F")),
    (9, 3, "關閉", "C", "danger-outline", lambda: close_group_windows("C")),
    (9, 4, "關閉", "D", "danger-outline", lambda: close_group_windows("D")),
    (9, 5, "關閉", "F", "danger-outline", lambda: close_group_windows("F")),
]
for row, col, text, code, bootstyle, cmd in group_btn_grid:
    btn = tb.Button(
        btns_outer_frame,
        text=f"{text} {group_display_names[code].get()}",
        bootstyle=bootstyle,
        command=cmd,
        width=8
    )
    btn.grid(row=row-8, column=col, padx=(4, 4), pady=(2, 2), sticky="nsew")
    if text == "啟動":
        group_buttons[code] = btn
    else:
        close_buttons[code] = btn

# --- row 10 檔案名稱/視窗名稱列表 ---
bottom_frame = tb.Frame(frm)
bottom_frame.grid(row=10, column=1, columnspan=2, sticky="ew", pady=(8, 2))
bottom_frame.grid_columnconfigure(0, weight=1)
bottom_frame.grid_columnconfigure(1, weight=1)
bottom_frame.grid_rowconfigure(0, weight=1)

file_list_outer = tb.Frame(bottom_frame, width=290)
file_list_outer.grid(row=0, column=0, sticky="nsw")
file_list_outer.grid_propagate(False)
file_list_outer.grid_rowconfigure(0, weight=1)
file_list_outer.grid_columnconfigure(0, weight=1)
file_list_canvas = tk.Canvas(file_list_outer, highlightthickness=0, height=120)
file_list_canvas.grid(row=0, column=0, sticky="nsew")
file_list_inner_frame = tb.Frame(file_list_canvas)
file_list_canvas.create_window((0, 0), window=file_list_inner_frame, anchor="nw")

window_list_outer = tb.Frame(bottom_frame)
window_list_outer.grid(row=0, column=1, sticky="nsew")
window_list_outer.grid_rowconfigure(0, weight=1)
window_list_outer.grid_columnconfigure(0, weight=1)
window_list_canvas = tk.Canvas(window_list_outer, highlightthickness=0)
window_list_canvas.grid(row=0, column=0, sticky="nsew")
window_list_inner_frame = tb.Frame(window_list_canvas)
window_list_canvas.create_window((0, 0), window=window_list_inner_frame, anchor="nw")

def _on_file_frame_configure(event):
    file_list_canvas.configure(scrollregion=file_list_canvas.bbox("all"))
file_list_inner_frame.bind("<Configure>", _on_file_frame_configure)

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
        lbl = tb.Label(file_list_inner_frame, text=fname, anchor="w", width=15, font=tkfont.Font(family="Microsoft JhengHei", size=10))
        lbl.grid(row=row, column=0, sticky="ew", padx=2, pady=1)
        lbl.bind("<ButtonPress-1>", lambda e, t=fname: on_label_drag_start(e, t))
    file_list_inner_frame.update_idletasks()
    file_list_canvas.config(scrollregion=file_list_canvas.bbox("all"))

def update_window_list():
    # 清空現有視窗列表
    for widget in window_list_inner_frame.winfo_children():
        widget.destroy()
    titles = get_taskbar_window_titles()
    col_count = (len(titles) + 4) // 5  # 每欄5個，計算需要幾欄
    for col in range(col_count):
        for row in range(5):
            idx = col * 5 + row
            if idx >= len(titles):
                break
            title = titles[idx]
            lbl = tb.Label(window_list_inner_frame, text=title, anchor="w", width=22, font=tkfont.Font(family="Microsoft JhengHei", size=10))
            lbl.grid(row=row, column=col, sticky="ew", padx=2, pady=1)
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

log_history = []

def log(msg):
    timestamp = time.strftime("%H:%M:%S")
    full_msg = f"[{timestamp}] {msg}"
    log_history.append(full_msg)
    if log_text.winfo_exists():
        log_text.config(state="normal")
        log_text.insert("end", full_msg + "\n")
        log_text.see("end")
        log_text.config(state="disabled")

app.mainloop()