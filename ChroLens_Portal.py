### 2025/11/21 By Lucienwooo
### 視窗管理工具 - 分組啟動、快捷切換、智能布局記憶
### 需要管理者權限（會自動要求提升）
#
# ============================================================================
# === 專案文件清理規則（給 AI Agent 的重要指示）===
# ============================================================================
# 
# 🎯 核心原則：保持專案精簡，避免累積測試檔案和臨時文檔
#
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 【✅ 必須保留的檔案】
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 
# 📄 核心程式檔案：
#   - ChroLens_Portal.py (主程式)
#   - update_manager.py (更新管理器)
#   - update_dialog.py (更新對話框)
#   - lang.py (多語言支援)
#   - build_simple.py (打包腳本)
#
# 📄 設定與資料：
#   - chrolens_portal.json (使用者設定檔)
#   - user_config.json (使用者配置)
#   - requirements.txt (依賴套件清單)
#
# 📄 文檔檔案：
#   - README.md (專案說明文件 - 根目錄唯一)
#   - CHANGELOG.md (版本更新紀錄)
#   - UPDATE.md (版本更新說明)
#   - LICENSE (授權文件)
#
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 【❌ 必須刪除的檔案】（每次 AI 操作後清理）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#
# 🗑️ 測試腳本：
#   - *_test.py (例如: stability_test.py, multi_group_test.py)
#   - test_*.py (例如: test_update_feature.py)
#   - run_*.py (例如: run_test.py)
#
# 🗑️ 臨時文檔：
#   - TEST_*.md (例如: TEST_REPORT.md)
#   - REPAIR_*.md (例如: REPAIR_REPORT.md)
#   - *_NOTES.md (例如: DEVELOPMENT_NOTES.md)
#   - *_GUIDE.md (例如: LAYOUT_GUIDE.md, USER_GUIDE.md)
#   - CHANGELOG_*.md (臨時更新日誌)
#
# 🗑️ 備份檔案：
#   - backup_*.json (例如: backup_settings.json)
#   - *.bak (任何備份檔案)
#   - *_backup.* (備份檔案)
#
# 🗑️ 臨時檔案：
#   - __pycache__/ (Python 快取目錄)
#   - *.pyc, *.pyo (編譯後的 Python 檔案)
#   - .pytest_cache/ (pytest 快取)
#   - .coverage (測試覆蓋率檔案)
#
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 【🔧 自動清理命令】（PowerShell - 在完成任務後執行）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#
# 完整清理命令（複製貼上即可執行）：
#
# cd "c:\Users\Lucien\Documents\GitHub\ChroLens_Portal"
# Remove-Item *_test.py -Force -ErrorAction SilentlyContinue
# Remove-Item test_*.py -Force -ErrorAction SilentlyContinue
# Remove-Item run_*.py -Force -ErrorAction SilentlyContinue
# Remove-Item TEST_*.md -Force -ErrorAction SilentlyContinue
# Remove-Item REPAIR_*.md -Force -ErrorAction SilentlyContinue
# Remove-Item *_NOTES.md -Force -ErrorAction SilentlyContinue
# Remove-Item *_GUIDE.md -Force -ErrorAction SilentlyContinue
# Remove-Item CHANGELOG_*.md -Force -ErrorAction SilentlyContinue
# Remove-Item backup_*.json -Force -ErrorAction SilentlyContinue
# Remove-Item *.bak -Force -ErrorAction SilentlyContinue
# Remove-Item *_backup.* -Force -ErrorAction SilentlyContinue
# Remove-Item -Recurse -Force __pycache__ -ErrorAction SilentlyContinue
# Remove-Item *.pyc -Force -ErrorAction SilentlyContinue
# Remove-Item *.pyo -Force -ErrorAction SilentlyContinue
# Get-ChildItem *.md  # 顯示剩餘的 .md 檔案（應只有 README.md, CHANGELOG.md, UPDATE.md, LICENSE）
#
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 【⚠️ AI Agent 作業流程】
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#
# 1. 執行任務（開發、測試、修復等）
# 2. 測試完成後，立即執行上方的清理命令
# 3. 驗證只剩下必須保留的檔案
# 4. 完成任務回報
#
# ⚠️  重要：不要保留任何測試檔案或臨時文檔！
# ⚠️  保持專案清潔是維護品質的關鍵！
#
# ============================================================================
#
# === 版本更新紀錄 ===
# 詳見 CHANGELOG.md 檔案
#

CURRENT_VERSION = "2.5.3"
GITHUB_REPO = "Lucienwooo/ChroLens_Portal"
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
import ctypes
from lang import LANG_MAP, get_text

SETTINGS_FILE = "chrolens_portal.json"

# === DPI 感知設定 ===
def set_dpi_awareness():
    """設定 DPI 感知，支援高解析度和縮放顯示器"""
    try:
        # Windows 10 / 11 - Per Monitor DPI Awareness V2
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except:
        try:
            # Windows 8.1 / 10 - Per Monitor DPI Awareness
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            try:
                # Windows Vista / 7 / 8 - System DPI Awareness
                ctypes.windll.user32.SetProcessDPIAware()
            except:
                pass

def get_dpi_scale():
    """取得當前 DPI 縮放比例"""
    try:
        # 獲取主螢幕的 DPI
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, hdc)
        scale = dpi / 96.0  # 96 DPI 是 100% 縮放
        return scale
    except:
        return 1.0

# 在任何 GUI 建立之前設定 DPI 感知
set_dpi_awareness()

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

# === 啟動時清理測試檔案和多餘文檔 ===
def cleanup_test_files():
    """清理測試檔案和臨時文檔（保持專案精簡）"""
    import glob
    
    # 定義需要清理的檔案模式（依照上方清理規範）
    cleanup_patterns = [
        # 測試腳本
        "*_test.py",
        "test_*.py",
        "run_*.py",
        # 臨時文檔
        "TEST_*.md",
        "REPAIR_*.md",
        "*_NOTES.md",
        "*_GUIDE.md",
        "CHANGELOG_*.md",
        # 備份檔案
        "backup_*.json",
        "*.bak",
        "*_backup.*",
        # Python 快取
        "*.pyc",
        "*.pyo"
    ]
    
    # 需要保留的檔案（白名單）
    keep_files = [
        "README.md",
        "UPDATE.md",
        "LICENSE",
        "chrolens_portal.json",
        "user_config.json"
    ]
    
    cleaned_count = 0
    for pattern in cleanup_patterns:
        for file_path in glob.glob(pattern):
            # 檢查是否在保留清單中
            if os.path.basename(file_path) not in keep_files:
                try:
                    os.remove(file_path)
                    print(f"✓ 已清理：{file_path}")
                    cleaned_count += 1
                except Exception as e:
                    print(f"✗ 清理失敗：{file_path} - {e}")
    
    if cleaned_count > 0:
        print(f"🧹 清理完成，共移除 {cleaned_count} 個檔案")
    else:
        print("✓ 無需清理")

# 執行清理（靜默模式，不影響啟動速度）
try:
    cleanup_test_files()
except Exception as e:
    print(f"清理過程發生錯誤：{e}")

# === 介面區塊 ===
app = tb.Window(themename="darkly")
app.title(f"ChroLens_Portal {CURRENT_VERSION}")
try:
    ico_path = resource_path("冥想貓貓.ico")
    app.iconbitmap(ico_path)
except Exception as e:
    print(f"無法設定 icon: {e}")

# 取得 DPI 縮放比例並計算自適應尺寸
dpi_scale = get_dpi_scale()
print(f"檢測到 DPI 縮放: {dpi_scale:.2f}x ({int(dpi_scale * 100)}%)")

# 根據 DPI 縮放調整字型大小和間距
base_font_size = 10
if dpi_scale >= 1.5:
    scaled_font_size = 9  # 150% 時使用 9pt，保持清晰
elif dpi_scale >= 1.25:
    scaled_font_size = 9  # 125% 時使用 9pt
else:
    scaled_font_size = 10  # 100% 時使用標準 10pt
num_font = tkfont.Font(family="Microsoft JhengHei", size=scaled_font_size, weight="bold")

# 根據 DPI 縮放調整視窗大小和 padding
base_width = 1400  # 增加基礎寬度，確保能完整顯示 3 欄 15 格分組
base_height = 750  # 略微增加高度

# 針對高 DPI 的自適應調整策略
if dpi_scale >= 1.5:  # 150% 或更高
    # 高 DPI 時保持較大的視窗，避免內容被壓縮
    scaled_width = int(base_width * 0.95)  # 只略微縮小
    scaled_height = int(base_height * 0.95)
elif dpi_scale >= 1.25:  # 125% 縮放
    scaled_width = int(base_width * 0.92)
    scaled_height = int(base_height * 0.92)
else:  # 100% 或更低
    scaled_width = base_width
    scaled_height = base_height

# 設定固定視窗大小（強制固定寬度）
fixed_width = 1400  # 固定寬度
min_height = 800  # 增加高度以容納分組欄位
scaled_width = fixed_width
scaled_height = max(min_height, scaled_height)

app.geometry(f"{scaled_width}x{scaled_height}")
app.minsize(fixed_width, min_height)  # 設定最小視窗大小
app.maxsize(fixed_width, 9999)  # 強制固定寬度，只允許調整高度

# 監聽視窗大小變化，強制保持固定寬度
def on_window_resize(event=None):
    """視窗大小改變時的處理，強制保持固定寬度"""
    if event and event.widget == app:
        current_width = app.winfo_width()
        if current_width != fixed_width:
            app.geometry(f"{fixed_width}x{app.winfo_height()}")
    app.update_idletasks()

app.bind("<Configure>", on_window_resize)

# 根據 DPI 調整 padding（高 DPI 時使用更合適的 padding）
if dpi_scale >= 1.5:
    adaptive_padding = 2  # 150% 時仍使用適當的 padding，避免過於擁擠
elif dpi_scale >= 1.25:
    adaptive_padding = 2  # 125% 時使用標準 padding
else:
    adaptive_padding = 2  # 100% 時使用標準 padding

# --- 主 Frame ---
frm = tb.Frame(app, padding=adaptive_padding)
frm.pack(fill="both", expand=True)

# 設定響應式行配置
frm.grid_rowconfigure(0, weight=0)  # 頂部工具列 - 固定
frm.grid_rowconfigure(1, weight=0)  # 置頂切換區 - 固定
frm.grid_rowconfigure(2, weight=0)  # 分組檔案列 - 固定高度
frm.grid_rowconfigure(8, weight=0)  # 按鈕區 - 固定
frm.grid_rowconfigure(9, weight=0)  # 按鈕區 - 固定
frm.grid_rowconfigure(10, weight=1)  # 檔案/視窗列表 - 可擴展（填滿剩餘空間）

# 設定響應式欄配置
frm.grid_columnconfigure(0, weight=0, minsize=150)  # 日誌框固定寬度
for col_idx in range(1, 8):
    frm.grid_columnconfigure(col_idx, weight=1, uniform="column")  # 其他欄位平均分配

# --- 分組與快捷鍵（動態支援 2-10 組）---
all_group_codes = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
all_default_hotkeys = ["Alt+1", "Alt+2", "Alt+3", "Alt+q", "Alt+w", "Alt+e", "Alt+a", "Alt+s", "Alt+d", "Alt+f"]

# Dracula 配色方案（VS Code Dracula Theme）+ 自訂顏色
DRACULA_COLORS = {
    "粉紅": "#FF79C6",    # Pink
    "紫色": "#BD93F9",    # Purple
    "青色": "#8BE9FD",    # Cyan
    "綠色": "#50FA7B",    # Green
    "橙色": "#FFB86C",    # Orange
    "紅色": "#FF5555",    # Red
    "黃色": "#F1FA8C",    # Yellow
    "藍紫": "#6272A4",    # Comment (藍灰)
    "白色": "#F8F8F2",    # Foreground
    "無色": "",             # 不設定顏色
    "自訂": "custom"      # 自訂顏色（特殊標記）
}

# 當前使用的分組數量（預設 6 組）
active_group_count = tk.IntVar(value=6)
group_codes = all_group_codes[:active_group_count.get()]
group_display_names = {c: tk.StringVar(value=c) for c in all_group_codes}
group_colors = {c: tk.StringVar(value="無色") for c in all_group_codes}  # 分組顏色設定
group_custom_colors = {c: "" for c in all_group_codes}  # 儲存自訂顏色的十六進位值
group_hotkeys = [tk.StringVar(value=all_default_hotkeys[i]) for i in range(10)]
group_buttons = {}
close_buttons = {}

# --- Mini 模式狀態追踪（用於資源優化）---
mini_mode_active = False
mini_window = None  # Mini 模式視窗實例

# --- 語言設定（多語言功能）---
current_language = "繁體中文"  # 預設語言
lang_map = LANG_MAP["繁體中文"]  # 當前語言的翻譯字典

# --- 排程任務（需要在程式開頭定義以避免 save_settings 錯誤）---
schedule_tasks = []

# --- 視窗位置和大小記憶 (FancyZones 功能) ---
# 儲存格式: { "group_code": { "file_name": {"x": int, "y": int, "width": int, "height": int, "state": str} } }
window_layouts = {}

# --- row 0：頂部工具列（響應式增強）---
top_row_frame = tb.Frame(frm, padding=adaptive_padding)
top_row_frame.grid(row=0, column=0, columnspan=8, sticky="ew", pady=(adaptive_padding, adaptive_padding))
# 左側固定寬度，右側靠右對齊
for col_idx in range(3):
    top_row_frame.grid_columnconfigure(col_idx, weight=0)  # 左側固定
top_row_frame.grid_columnconfigure(3, weight=1)  # 中間可擴展，推動右側靠右
for col_idx in range(4, 14):
    top_row_frame.grid_columnconfigure(col_idx, weight=0)  # 右側固定

folder_var = tb.StringVar(value="")
interval_var = tb.StringVar(value="4")

def choose_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_var.set(folder_selected)
        if not mini_mode_active:
            update_file_list()  # 只在非 mini 模式下刷新檔案列表

folder_frame = tb.Frame(top_row_frame, padding=(adaptive_padding, adaptive_padding))
folder_frame.grid(row=0, column=0, sticky="w", padx=(0, 4))
# 隱藏路徑顯示框，只保留按鈕
# tb.Entry(folder_frame, textvariable=folder_var, width=30).grid(row=0, column=0, padx=(adaptive_padding, adaptive_padding))
choose_path_btn = tb.Button(folder_frame, text=lang_map["選擇開啟路徑"], command=lambda: choose_folder(), bootstyle=SECONDARY, width=12)
choose_path_btn.grid(row=0, column=0, padx=(adaptive_padding, 0))

# 啟動間隔（調換到前面）
interval_frame = tb.Frame(top_row_frame, padding=(adaptive_padding, adaptive_padding))
interval_frame.grid(row=0, column=1, sticky="w", padx=(0, 4))
interval_label = tb.Label(interval_frame, text="啟動間隔(秒):")
interval_label.grid(row=0, column=0, sticky="w")
tb.Entry(interval_frame, textvariable=interval_var, width=3).grid(row=0, column=1, padx=(adaptive_padding, 0), sticky="w")

# 分組數量選擇器（調換到後面）
group_count_frame = tb.Frame(top_row_frame, padding=(adaptive_padding, adaptive_padding))
group_count_frame.grid(row=0, column=2, sticky="w", padx=(0, 4))
group_count_label = tb.Label(group_count_frame, text="分組數量:")
group_count_label.grid(row=0, column=0, sticky="w")
group_count_combo = tb.Combobox(
    group_count_frame,
    textvariable=active_group_count,
    values=[2, 3, 4, 5, 6, 7, 8, 9, 10],
    width=3,
    state="readonly"
)
group_count_combo.grid(row=0, column=1, padx=(adaptive_padding, 0), sticky="w")

# === 視窗佈局記憶功能 (FancyZones) ===
def capture_window_layout(group_code):
    """捕獲指定分組所有視窗的位置和大小"""
    files = get_group_files(group_code)
    if not files:
        log(f"{lang_map['分組']} {group_display_names[group_code].get()} {lang_map['沒有檔案，無法捕獲佈局']}")
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
                    log(f"{lang_map['捕獲視窗佈局：']}{filename} ({rect[2]-rect[0]}x{rect[3]-rect[1]} at {rect[0]},{rect[1]})")
                except Exception as e:
                    log(f"{lang_map['捕獲視窗佈局失敗：']}{window_text} ({e})")
                break
    
    win32gui.EnumWindows(enum_handler, None)
    
    if captured:
        if group_code not in window_layouts:
            window_layouts[group_code] = {}
        window_layouts[group_code] = captured
        save_settings()
        log(f"{lang_map['已捕獲分組']} {group_display_names[group_code].get()} {lang_map['的']} {len(captured)} {lang_map['個視窗佈局']}")
    else:
        log(f"{lang_map['未找到分組']} {group_display_names[group_code].get()} {lang_map['的任何視窗']}")

def restore_window_layout(group_code, hwnd, window_text):
    """恢復指定視窗的位置和大小"""
    if group_code not in window_layouts:
        log(f"{lang_map['[佈局] 分組']} {group_code} {lang_map['沒有保存的佈局資料']}")
        return False
    
    layout = window_layouts[group_code]
    if not layout:
        log(f"{lang_map['[佈局] 分組']} {group_code} {lang_map['的佈局資料為空']}")
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
                    log(f"{lang_map['[佈局] 恢復最大化：']}{window_text}")
                elif state == win32con.SW_SHOWMINIMIZED:
                    # 最小化視窗
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)
                    log(f"{lang_map['[佈局] 恢復最小化：']}{window_text}")
                else:
                    # 正常視窗：直接設置位置和大小，不調用 ShowWindow
                    # 使用 SWP_NOZORDER 避免改變 Z-order
                    win32gui.SetWindowPos(
                        hwnd, 
                        win32con.HWND_TOP,
                        x, y, width, height,
                        win32con.SWP_SHOWWINDOW | win32con.SWP_NOZORDER
                    )
                    log(f"{lang_map['[佈局] 恢復位置：']}{window_text} -> {width}x{height} at ({x},{y})")
                return True
            except Exception as e:
                log(f"{lang_map['[佈局] 恢復失敗：']}{window_text} ({e})")
                return False
    
    log(f"{lang_map['[佈局] 未找到匹配的佈局：']}{window_text}")
    return False

# 新增「存檔」按鈕（捕獲當前視窗並選擇要儲存到哪些分組）
def manual_save():
    """彈出對話框讓使用者選擇要儲存當前視窗佈局到哪些分組"""
    
    # 根據分組數量計算對話框高度（每個分組約 35px，加上頂部和底部空間）
    count = active_group_count.get()
    base_height = 180  # 頂部標題、提示文字、底部按鈕的基本高度
    checkbox_height = count * 40  # 每個複選框的高度
    dialog_height = base_height + checkbox_height
    dialog_height = min(dialog_height, 700)  # 最大高度 700px
    dialog_height = max(dialog_height, 300)  # 最小高度 300px
    
    # 建立選擇對話框
    save_dialog = tk.Toplevel(app)
    save_dialog.title("選擇儲存目標分組")
    save_dialog.transient(app)
    save_dialog.grab_set()
    
    # 置中顯示
    save_dialog.update_idletasks()
    screen_width = save_dialog.winfo_screenwidth()
    screen_height = save_dialog.winfo_screenheight()
    x = (screen_width // 2) - 200
    y = (screen_height // 2) - (dialog_height // 2)
    save_dialog.geometry(f"400x{dialog_height}+{x}+{y}")
    
    # 標題
    title_label = tb.Label(
        save_dialog,
        text="將當前開啟的視窗佈局儲存到：",
        font=("Microsoft JhengHei", 12, "bold")
    )
    title_label.pack(pady=(20, 10))
    
    # 提示文字
    hint_label = tb.Label(
        save_dialog,
        text="（可選擇多個分組，當前視窗配置會套用到所選分組）",
        font=("Microsoft JhengHei", 9),
        foreground="#888"
    )
    hint_label.pack(pady=(0, 15))
    
    # 複選框區域
    checkbox_frame = tb.Frame(save_dialog)
    checkbox_frame.pack(fill="both", expand=True, padx=30, pady=10)
    
    # 為每個分組創建複選框
    group_check_vars = {}
    for idx, code in enumerate(group_codes):
        var = tk.BooleanVar(value=False)
        group_check_vars[code] = var
        
        # 檢查該分組是否有檔案
        files = get_group_files(code)
        has_files = len(files) > 0
        
        cb = tb.Checkbutton(
            checkbox_frame,
            text=f"{group_display_names[code].get()} ({len(files)} 個檔案)",
            variable=var,
            bootstyle="success-round-toggle",
            state="normal" if has_files else "disabled"
        )
        cb.pack(anchor="w", pady=5)
    
    # 按鈕區域
    button_frame = tb.Frame(save_dialog)
    button_frame.pack(pady=(10, 20))
    
    def do_save():
        """執行儲存"""
        selected_groups = [code for code, var in group_check_vars.items() if var.get()]
        
        if not selected_groups:
            messagebox.showwarning("未選擇分組", "請至少選擇一個分組來儲存視窗佈局", parent=save_dialog)
            return
        
        save_dialog.destroy()
        
        log("=" * 50)
        log(f"開始捕獲視窗佈局並儲存到 {len(selected_groups)} 個分組...")
        total_captured = 0
        
        # 先收集當前所有視窗的位置
        captured_windows = {}
        my_hwnd = app.winfo_id()
        
        def enum_handler(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            if hwnd == my_hwnd:
                return
            
            window_text = win32gui.GetWindowText(hwnd)
            if not window_text.strip():
                return
            
            try:
                rect = win32gui.GetWindowRect(hwnd)
                placement = win32gui.GetWindowPlacement(hwnd)
                
                captured_windows[window_text.lower()] = {
                    "text": window_text,
                    "x": rect[0],
                    "y": rect[1],
                    "width": rect[2] - rect[0],
                    "height": rect[3] - rect[1],
                    "state": placement[1]
                }
            except Exception as e:
                pass
        
        win32gui.EnumWindows(enum_handler, None)
        log(f"找到 {len(captured_windows)} 個可見視窗")
        
        # 將視窗配置套用到選中的分組
        for group_code in selected_groups:
            files = get_group_files(group_code)
            if not files:
                log(f"分組 {group_display_names[group_code].get()} 沒有檔案，跳過")
                continue
            
            # 為這個分組匹配視窗
            group_layout = {}
            for f in files:
                if f:
                    file_key = os.path.splitext(os.path.basename(f))[0].lower()
                    # 在捕獲的視窗中尋找匹配
                    for win_text_lower, win_data in captured_windows.items():
                        if file_key in win_text_lower:
                            group_layout[f] = {
                                "x": win_data["x"],
                                "y": win_data["y"],
                                "width": win_data["width"],
                                "height": win_data["height"],
                                "state": win_data["state"]
                            }
                            log(f"  [{group_display_names[group_code].get()}] {os.path.basename(f)} -> {win_data['width']}x{win_data['height']} at ({win_data['x']},{win_data['y']})")
                            total_captured += 1
                            break
            
            if group_layout:
                if group_code not in window_layouts:
                    window_layouts[group_code] = {}
                window_layouts[group_code] = group_layout
                log(f"已儲存 {len(group_layout)} 個視窗到分組 {group_display_names[group_code].get()}")
            else:
                log(f"分組 {group_display_names[group_code].get()} 沒有找到匹配的視窗")
        
        save_settings()
        log(f"已手動儲存設定檔，共儲存 {total_captured} 個視窗佈局到 {len(selected_groups)} 個分組")
        log("=" * 50)
    
    def cancel_save():
        """取消儲存"""
        save_dialog.destroy()
    
    # 確定和取消按鈕
    save_btn = tb.Button(
        button_frame,
        text="確定儲存",
        command=do_save,
        bootstyle="success",
        width=12
    )
    save_btn.pack(side="left", padx=5)
    
    cancel_btn = tb.Button(
        button_frame,
        text="取消",
        command=cancel_save,
        bootstyle="secondary",
        width=12
    )
    cancel_btn.pack(side="left", padx=5)
    
    # 綁定 ESC 鍵取消
    save_dialog.bind("<Escape>", lambda e: cancel_save())

save_btn = tb.Button(top_row_frame, text=lang_map["存檔"], command=manual_save, bootstyle="success")
save_btn.grid(row=0, column=5, padx=(4,2), sticky="e")

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
    
    # 更新顏色選擇器
    current_color = group_colors[code].get()
    if current_color in DRACULA_COLORS:
        group_color_combo.set(current_color)
    else:
        group_color_combo.set("無色")
    
    # 更新顏色預覽
    update_color_preview(code)

def update_color_preview(code=None):
    """更新顏色預覽框"""
    if code is None:
        code = group_name_edit_combo_var.get()
    color_name = group_colors[code].get()
    if color_name == "自訂":
        # 顯示自訂顏色
        custom_color = group_custom_colors.get(code, "")
        if custom_color:
            group_color_preview.config(background=custom_color)
        else:
            group_color_preview.config(background="#2b3e50")
    elif color_name in DRACULA_COLORS:
        color_hex = DRACULA_COLORS[color_name]
        if color_hex:
            group_color_preview.config(background=color_hex)
        else:
            group_color_preview.config(background="#2b3e50")
    else:
        group_color_preview.config(background="#2b3e50")

def open_color_chooser():
    """開啟顏色選擇器"""
    from tkinter import colorchooser
    
    code = group_name_edit_combo_var.get()
    # 獲取當前顏色作為預設值
    current_color_name = group_colors[code].get()
    if current_color_name == "自訂":
        initial_color = group_custom_colors.get(code, "#FF79C6")
    elif current_color_name in DRACULA_COLORS:
        initial_color = DRACULA_COLORS[current_color_name]
        if not initial_color:
            initial_color = "#FF79C6"
    else:
        initial_color = "#FF79C6"
    
    # 開啟顏色選擇器
    color = colorchooser.askcolor(title="選擇分組顏色", initialcolor=initial_color)
    if color[1]:  # color[1] 是十六進位顏色值
        # 儲存自訂顏色
        group_custom_colors[code] = color[1]
        group_colors[code].set("自訂")
        group_color_combo.set("自訂")
        update_color_preview(code)
        log(f"分組 {code} 自訂顏色已設為：{color[1]}")
        # 立即更新按鈕顏色
        update_button_colors()
        save_settings()

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
    
    # 保存顏色設定
    new_color = group_color_combo.get()
    if new_color == "自訂":
        # 如果選擇自訂，開啟顏色選擇器
        open_color_chooser()
    elif new_color in DRACULA_COLORS:
        group_colors[code].set(new_color)
        # 清除自訂顏色（如果有）
        if code in group_custom_colors:
            group_custom_colors[code] = ""
        update_color_preview(code)
        log(f"分組 {code} 顏色已設為：{new_color}")
        # 立即更新按鈕顏色
        update_button_colors()

group_name_edit_combo = tb.Combobox(
    top_row_frame, values=group_codes, textvariable=group_name_edit_combo_var, width=3, state="readonly"
)
group_name_edit_entry = tb.Entry(
    top_row_frame, textvariable=group_name_edit_var, width=12
)
group_name_edit_entry.placeholder = False

# 顏色選擇下拉選單（Dracula 配色 + 自訂）
group_color_combo = tb.Combobox(
    top_row_frame,
    values=list(DRACULA_COLORS.keys()),
    width=6,
    state="readonly"
)
group_color_combo.current(9)  # 預設選擇「無色」

# 自訂顏色按鈕（顯示當前顏色）
group_color_preview = tb.Label(
    top_row_frame,
    text="  ",
    width=3,
    relief="solid",
    borderwidth=1,
    background="#2b3e50",  # 預設背景
    cursor="hand2"  # 游標變手指
)
# 綁定點擊事件
group_color_preview.bind("<Button-1>", lambda e: open_color_chooser())

group_name_edit_combo.grid(row=0, column=4, padx=(8,2), sticky="e")
group_name_edit_entry.grid(row=0, column=5, padx=(2,2), sticky="e")
group_color_combo.grid(row=0, column=6, padx=(2,2), sticky="e")  # 顏色選擇器
group_color_preview.grid(row=0, column=7, padx=(2,2), sticky="e")  # 顏色預覽
save_btn.grid(row=0, column=8, padx=(8,2), sticky="e")

# 綁定顏色選擇器變化事件
group_color_combo.bind("<<ComboboxSelected>>", lambda e: update_color_preview())

# 定義更新 group_name_edit_combo 的函數（將在 rebuild_group_buttons 時調用）
def update_group_name_combo():
    """更新分組命名下拉選單的選項"""
    group_name_edit_combo.config(values=group_codes)
    # 如果當前選中的選項已經不在新的 group_codes 中，重設為第一個
    current = group_name_edit_combo_var.get()
    if current not in group_codes:
        group_name_edit_combo_var.set(group_codes[0] if group_codes else "A")
        on_group_name_combo_change()

# 綁定事件
group_name_edit_combo.bind("<<ComboboxSelected>>", on_group_name_combo_change)
group_name_edit_entry.bind("<FocusIn>", lambda e: (hide_placeholder(), group_name_edit_entry.config(foreground="#fff")))
group_name_edit_entry.bind("<FocusOut>", lambda e: (show_placeholder() if not group_name_edit_entry.get().strip() else None, on_group_name_edit_submit()))
group_name_edit_entry.bind("<Return>", on_group_name_edit_submit)
group_color_combo.bind("<<ComboboxSelected>>", lambda e: on_group_name_edit_submit())

# 初始化顯示 placeholder
show_placeholder()

# --- row 1：分組置頂顯示切換區（填滿整行）---
show_label_frames = []
second_row_frame = tb.Frame(frm)
second_row_frame.grid(row=1, column=0, columnspan=8, sticky="nsew", padx=0, pady=0)
show_label_font = tkfont.Font(family="Microsoft JhengHei", size=12)
# 置頂切換按鈕區域將平鈖整行寬度

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

def rebuild_topmost_buttons():
    """根據當前分組數量重建置頂切換按鈕（每行最多5個，最多2行，置中對齊）"""
    # 清除所有舊按鈕
    for widget in second_row_frame.winfo_children():
        widget.destroy()
    
    show_label_frames.clear()
    
    count = active_group_count.get()
    buttons_per_row = 5  # 每行最多 5 個
    
    # 清空所有欄位配置
    for i in range(buttons_per_row):
        second_row_frame.grid_columnconfigure(i, weight=0)
    
    # 配置行（最多 2 行）
    second_row_frame.grid_rowconfigure(0, weight=0)
    second_row_frame.grid_rowconfigure(1, weight=0)
    
    # 計算第一行和第二行各有幾個按鈕
    first_row_count = min(count, buttons_per_row)
    second_row_count = max(0, count - buttons_per_row)
    
    # 生成新按鈕
    for idx, code in enumerate(group_codes):
        row = idx // buttons_per_row  # 0-4 在第 0 行，5-9 在第 1 行
        col_in_row = idx % buttons_per_row
        
        # 計算實際欄位位置（置中對齊）
        if row == 0:
            # 第一行：居中顯示
            col_offset = (buttons_per_row - first_row_count) // 2
            actual_col = col_offset + col_in_row
        else:
            # 第二行：居中顯示
            col_offset = (buttons_per_row - second_row_count) // 2
            actual_col = col_offset + (idx - buttons_per_row)
        
        frame = tb.Frame(second_row_frame, borderwidth=max(1, int(2 / dpi_scale)), relief="groove", padding=2)
        frame.grid(row=row, column=actual_col, padx=adaptive_padding, pady=adaptive_padding, sticky="ew")
        
        # 讓 frame 內部平均分配空間
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=0)
        
        # 置頂切換按鈕（使用 wraplength 防止文字過長）
        show_btn = tb.Button(
            frame, 
            text=group_display_names[code].get(), 
            command=lambda c=code: focus_next_in_group(c),
            bootstyle="info-outline"
        )
        show_btn.grid(row=0, column=0, sticky="ew", padx=(0, adaptive_padding))
        # 確保按鈕文字不會影響寬度，超出部分截斷
        show_btn.config(compound="center")
        
        # 快捷鍵文字框（固定寬度）
        hotkey_entry = tb.Entry(
            frame, 
            textvariable=group_hotkeys[idx], 
            width=8, 
            state="readonly", 
            justify="center", 
            font=show_label_font,
            foreground="#ffffff",
            style="Dark.TEntry"
        )
        hotkey_entry.grid(row=0, column=1, sticky="e")
        
        def make_on_key(idx):
            return lambda event, i=idx: on_hotkey_entry_key(event, i)
        hotkey_entry.bind("<Key>", make_on_key(idx))
        hotkey_entry.bind("<Button-1>", lambda e, entry=hotkey_entry: entry.focus_set())
        show_label_frames.append((show_btn, hotkey_entry))
    
    # 配置有按鈕的欄位為平均分配
    if first_row_count > 0:
        for i in range(buttons_per_row):
            second_row_frame.grid_columnconfigure(i, weight=1, uniform="topmost_col")

# 注意：rebuild_topmost_buttons() 的初始化將在所有函數定義後執行

def update_show_labels(*args):
    for idx, code in enumerate(group_codes):
        show_label_frames[idx][0].config(text=group_display_names[code].get())  
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
        log(f"{lang_map['已設定分組']} {group_codes[idx]} {lang_map['的快捷鍵為：']}{hotkey}")

# --- 編號標籤點擊處理函數 ---
def on_num_label_click(event, entry):
    """點擊編號標籤時開啟對應的檔案"""
    open_entry_file(entry)

# --- row 2：分組檔案列 ---
group_frames = []
for col in range(3):
    group_frame = tb.Frame(frm, borderwidth=max(1, int(1 / dpi_scale)), relief="solid", padding=adaptive_padding)
    group_frame.grid(row=2, column=col, padx=adaptive_padding, pady=adaptive_padding, sticky="nsew")
    frm.grid_columnconfigure(col, weight=1, uniform="group_col")  # 使用 uniform 確保三欄等寬
    # 配置內部欄位 - 讓所有 entry 欄位等寬
    group_frame.grid_columnconfigure(0, weight=0, minsize=30)  # 編號欄固定寬度
    group_frame.grid_columnconfigure(1, weight=1, minsize=150)  # 檔案名稱欄可擴展，最小寬度 150
    for i in range(2, 6):  # combo 欄位固定寬度
        group_frame.grid_columnconfigure(i, weight=0, minsize=35)
    group_frames.append(group_frame)

combo_width = 3  # 原本是4，縮短1/3

checkbox_vars_entries = []

# 定義更新 row2 所有 combo 選項的函數
def update_row2_combo_values():
    """更新 row2 所有分組下拉選單的選項"""
    new_values = [""] + [group_display_names[c].get() for c in group_codes]
    for _, var1, var2, var3, var4, combo1, combo2, combo3, combo4 in checkbox_vars_entries:
        # 暫存原本的值
        v1, v2, v3, v4 = var1.get(), var2.get(), var3.get(), var4.get()
        combo1.config(values=new_values)
        combo2.config(values=new_values)
        combo3.config(values=new_values)
        combo4.config(values=new_values)
        # 如果原值還在新 values 裡，還原；否則清空
        if v1 not in new_values:
            var1.set("")
        if v2 not in new_values:
            var2.set("")
        if v3 not in new_values:
            var3.set("")
        if v4 not in new_values:
            var4.set("")

for i in range(15):  # 15行
    row = i % 5
    col = i // 5
    # 移除固定寬度，讓 entry 自動根據 grid_columnconfigure 配置調整
    entry = tb.Entry(group_frames[col], state="readonly")
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
    # 編號標籤
    num_label = tb.Label(group_frames[col], text=str(i+1), width=2, font=("Microsoft JhengHei", scaled_font_size, "bold"), background="#444", foreground="#fff", anchor="center", cursor="hand2")
    num_label.grid(row=row, column=0, sticky="ew", padx=0)
    num_label.bind("<Button-1>", lambda e, ent=entry: on_num_label_click(e, ent))
    # Entry 欄位 - 使用 sticky="ew" 讓它自動填滿空間
    entry.grid(row=row, column=1, padx=1, pady=adaptive_padding, sticky="ew")
    # Combo 欄位 - 固定寬度
    group_combo1.grid(row=row, column=2, padx=0, pady=adaptive_padding, sticky="ew")
    group_combo2.grid(row=row, column=3, padx=0, pady=adaptive_padding, sticky="ew")
    group_combo3.grid(row=row, column=4, padx=0, pady=adaptive_padding, sticky="ew")
    group_combo4.grid(row=row, column=5, padx=0, pady=adaptive_padding, sticky="ew")
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
style.configure("Num.TButton", font=("Microsoft JhengHei", scaled_font_size, "bold"))
# 配置深色 Entry 樣式（用於快捷鍵文字框）
style.configure("Dark.TEntry", 
                fieldbackground="#2b3e50",  # 深色背景（與 darkly 主題一致）
                foreground="#ffffff",  # 白色文字
                bordercolor="#4e5d6c",  # 邊框顏色
                lightcolor="#4e5d6c",
                darkcolor="#2b3e50",
                insertcolor="#ffffff")  # 游標顏色

# --- row 8~10 動態日誌區塊 ---
log_text = tb.Text(frm, height=12, width=18, state="disabled", wrap="word", font=tkfont.Font(family="Microsoft JhengHei", size=scaled_font_size))
log_text.grid(row=8, column=0, rowspan=3, sticky="nsew", padx=(0, int(8 / dpi_scale)), pady=(0, 0))

# --- row 8~9 啟動/關閉分組按鈕區域（動態生成，平均擴散）---
btns_outer_frame = tb.Frame(frm)
btns_outer_frame.grid(row=8, column=1, rowspan=2, columnspan=7, sticky="nsew", padx=(0, 4), pady=(0, 0))

def rebuild_group_buttons():
    """根據當前分組數量重建啟動/關閉按鈕（每行 5 個，最多 2 行，左邊啟動右邊關閉）"""
    # 清除舊按鈕
    for widget in btns_outer_frame.winfo_children():
        widget.destroy()
    
    group_buttons.clear()
    close_buttons.clear()
    
    # 更新 group_codes
    global group_codes
    count = active_group_count.get()
    group_codes = all_group_codes[:count]
    
    # 更新分組命名下拉選單
    update_group_name_combo()
    
    # 更新 row2 的所有 combo 選項
    update_row2_combo_values()
    
    # 計算按鈕佈局（每行最多 5 個，左半部啟動，右半部關閉）
    buttons_per_row = 5
    total_cols = buttons_per_row * 2  # 左 5 個啟動 + 右 5 個關閉
    
    # 配置所有欄位平均分配空間
    for i in range(total_cols):
        btns_outer_frame.grid_columnconfigure(i, weight=1, uniform="btn_col")
    
    # 配置行（最多 2 行）
    btns_outer_frame.grid_rowconfigure(0, weight=1)
    btns_outer_frame.grid_rowconfigure(1, weight=1)
    
    # 生成啟動按鈕（左半部，每行 5 個）
    for idx, code in enumerate(group_codes):
        row = idx // buttons_per_row  # 0-4 在第 0 行，5-9 在第 1 行
        col = idx % buttons_per_row  # 0-4
        start_btn = tb.Button(
            btns_outer_frame,
            text=f"{group_display_names[code].get()}",
            bootstyle="success-outline",
            command=lambda c=code: start_group_opening(c),
            compound="center"  # 確保文字不影響寬度
        )
        start_btn.grid(row=row, column=col, padx=adaptive_padding, pady=adaptive_padding, sticky="nsew")
        group_buttons[code] = start_btn
    
    # 生成關閉按鈕（右半部，每行 5 個）
    for idx, code in enumerate(group_codes):
        row = idx // buttons_per_row
        col = (idx % buttons_per_row) + buttons_per_row  # 偏移 5 欄到右半部
        close_btn = tb.Button(
            btns_outer_frame,
            text=f"{group_display_names[code].get()}",
            bootstyle="danger-outline",
            command=lambda c=code: close_group_windows(c),
            compound="center"  # 確保文字不影響寬度
        )
        close_btn.grid(row=row, column=col, padx=adaptive_padding, pady=adaptive_padding, sticky="nsew")
        close_buttons[code] = close_btn
    
    # 更新置頂切換區的按鈕
    rebuild_topmost_buttons()
    
    # 更新所有按鈕顏色
    update_button_colors()
    
    # 如果 mini 模式已開啟，也需要更新
    if mini_mode_active and mini_window:
        mini_window.rebuild_buttons()

def update_button_colors():
    """根據分組顏色設定更新所有按鈕的顏色"""
    # 為每個顏色創建自定義樣式
    for code in group_codes:
        color_name = group_colors[code].get()
        
        # 決定使用的顏色
        if color_name == "自訂":
            # 使用自訂顏色
            color_hex = group_custom_colors.get(code, "")
        else:
            # 使用預設 Dracula 顏色
            color_hex = DRACULA_COLORS.get(color_name, "")
        
        if color_hex:
            # 創建自定義樣式名稱
            success_style = f"Group{code}.success.TButton"
            danger_style = f"Group{code}.danger.TButton"
            info_style = f"Group{code}.info.TButton"
            
            # 配置自定義樣式（使用 ttkbootstrap 的 Style）
            try:
                style.configure(success_style, background=color_hex, foreground="#282A36", bordercolor=color_hex)
                style.configure(danger_style, background=color_hex, foreground="#282A36", bordercolor=color_hex)
                style.configure(info_style, background=color_hex, foreground="#282A36", bordercolor=color_hex)
                
                # 配置按鈕懸停和按下狀態
                style.map(success_style,
                    background=[('active', color_hex), ('pressed', color_hex)],
                    foreground=[('active', '#282A36'), ('pressed', '#282A36')])
                style.map(danger_style,
                    background=[('active', color_hex), ('pressed', color_hex)],
                    foreground=[('active', '#282A36'), ('pressed', '#282A36')])
                style.map(info_style,
                    background=[('active', color_hex), ('pressed', color_hex)],
                    foreground=[('active', '#282A36'), ('pressed', '#282A36')])
            except Exception as e:
                print(f"配置樣式失敗: {e}")
            
            # 更新啟動按鈕
            if code in group_buttons:
                group_buttons[code].configure(style=success_style)
            
            # 更新關閉按鈕
            if code in close_buttons:
                close_buttons[code].configure(style=danger_style)
            
            # 更新置頂切換按鈕
            for idx, btn_code in enumerate(group_codes):
                if btn_code == code and idx < len(show_label_frames):
                    show_label_frames[idx][0].configure(style=info_style)
        else:
            # 恢復預設樣式
            if code in group_buttons:
                group_buttons[code].configure(bootstyle="success-outline")
            if code in close_buttons:
                close_buttons[code].configure(bootstyle="danger-outline")
            for idx, btn_code in enumerate(group_codes):
                if btn_code == code and idx < len(show_label_frames):
                    show_label_frames[idx][0].configure(bootstyle="info-outline")

# 注意：group_count_combo 的綁定和按鈕初始化將在所有函數定義後執行

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
                log(f"{lang_map['拖移「']}{title}{lang_map['」到分組欄位']}")
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
            log(f"{lang_map['清空分組欄位內容（原內容：']}{old}{lang_map['）']}")
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
            log(f"{lang_map['處理視窗失敗：']}{window_text} ({e})")

def start_group_opening(group_code):
    folder = folder_var.get()
    try:
        interval = float(interval_var.get())
    except ValueError:
        log(lang_map["請輸入正確的間隔秒數"])
        return
    if not os.path.isdir(folder):
        log(lang_map["請選擇正確的資料夾"])
        return
    files = get_group_files(group_code)
    if not files:
        log(f"{lang_map['分組']} {group_display_names[group_code].get()} {lang_map['沒有檔案']}")
        return
    log(f"{lang_map['開始開啟分組']} {group_display_names[group_code].get()} {lang_map['的檔案於']} {folder}")
    
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
                log(f"{lang_map['找不到檔案: ']}{file_path}")
                continue
            
            file_key = os.path.splitext(os.path.basename(file))[0]
            
            try:
                if file_path.lower().endswith('.lnk'):
                    # 解析捷徑
                    target, args = open_lnk_target(file_path)
                    if target and os.path.exists(target):
                        log(f"{lang_map['開啟捷徑目標: ']}{target} {args}")
                        if args:
                            subprocess.Popen(f'"{target}" {args}', shell=True)
                        else:
                            subprocess.Popen(f'"{target}"', shell=True)
                    else:
                        # 如果解析失敗，直接用 Windows 開啟捷徑
                        log(f"{lang_map['直接開啟捷徑: ']}{file_path}")
                        os.startfile(file_path)
                elif file_path.lower().endswith('.exe'):
                    # 直接執行 .exe 檔案
                    log(f"{lang_map['執行程式: ']}{file_path}")
                    subprocess.Popen(f'"{file_path}"', shell=True)
                else:
                    # 其他檔案類型用系統預設程式開啟
                    log(f"{lang_map['開啟檔案: ']}{file_path}")
                    os.startfile(file_path)
                
                # 等待視窗出現並恢復佈局
                wait_for_window_and_restore(file_key)
                
            except Exception as e:
                log(f"{lang_map['無法開啟: ']}{file_path}{lang_map['，錯誤：']}{e}")
            
            time.sleep(interval)
    
    threading.Thread(target=open_files, daemon=True).start()

def close_group_windows(group_code):
    files = get_group_files(group_code)
    if not files:
        log(f"{lang_map['分組']} {group_display_names[group_code].get()} {lang_map['沒有檔案']}")
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
                    log(f"{lang_map['已關閉視窗：']}{window_text}")
                    nonlocal closed_any
                    closed_any = True
                except Exception as e:
                    log(f"{lang_map['關閉視窗失敗：']}{window_text} ({e})")
    win32gui.EnumWindows(enum_handler, None)
    if not closed_any:
        log(f"{lang_map['找不到分組']} {group_display_names[group_code].get()} {lang_map['的視窗可關閉']}")

def save_settings():
    try:
        data = {
            "folder": folder_var.get(),
            "interval": interval_var.get(),
            "active_group_count": active_group_count.get(),  # 儲存分組數量
            "group_display_names": {c: group_display_names[c].get() for c in all_group_codes},  # 修正：儲存所有分組（A-J）
            "group_colors": {c: group_colors[c].get() for c in all_group_codes},  # 儲存分組顏色
            "group_custom_colors": {c: group_custom_colors.get(c, "") for c in all_group_codes},  # 儲存自訂顏色
            "group_hotkeys": [v.get() for v in group_hotkeys],
            "checkbox_entries": [entry.get() for entry, *_ in checkbox_vars_entries],
            "group_var1": [var1.get() for _, var1, _, _, _, *_ in checkbox_vars_entries],
            "group_var2": [var2.get() for _, _, var2, _, _, *_ in checkbox_vars_entries],
            "group_var3": [var3.get() for _, _, _, var3, _, *_ in checkbox_vars_entries],
            "group_var4": [var4.get() for _, _, _, _, var4, *_ in checkbox_vars_entries],
            "schedule_tasks": schedule_tasks,  # 儲存排程任務
            "window_layouts": window_layouts,  # 儲存視窗佈局
            "language": current_language,  # 儲存當前語言
        }
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except NameError:
        # 忽略變數未定義的錯誤（通常發生在程式初始化時）
        pass
    except Exception as e:
        # 只在非初始化錯誤時記錄
        if "is not defined" not in str(e):
            log(f"{lang_map['儲存設定檔失敗: ']}{e}")

        
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

# ============================================
# 語言切換函數
# ============================================

def apply_language(new_lang):
    """切換語言並更新所有 UI 元件"""
    global current_language, lang_map
    current_language = new_lang
    lang_map = LANG_MAP.get(new_lang, LANG_MAP["繁體中文"])
    
    # 更新頂部按鈕
    mini_btn.config(text=lang_map["mini"])
    schedule_btn.config(text=lang_map["排程"])
    about_btn.config(text=lang_map["關於"])
    
    # 更新第一行按鈕和標籤
    choose_path_btn.config(text=lang_map["選擇開啟路徑"])
    interval_label.config(text=lang_map["間隔秒數:"])
    
    # 更新第二行按鈕
    save_btn.config(text=lang_map["存檔"])
    
    # 置頂切換標籤已移除
    
    # 注意：部分 UI 元件需要在程式啟動時就使用 lang_map 來建立
    # 目前已更新的元件會在切換語言時即時更新
    
    log(f"{lang_map['語言已切換為：']}{new_lang}")

def load_settings():
    global schedule_tasks, window_layouts
    if not os.path.exists(SETTINGS_FILE):
        return
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        folder_var.set(data.get("folder", folder_var.get()))
        interval_var.set(data.get("interval", interval_var.get()))
        
        # 載入分組數量（必須在載入分組名稱之前）
        saved_group_count = data.get("active_group_count", 6)
        active_group_count.set(saved_group_count)
        
        # 修正：載入所有分組（A-J）的名稱
        for c in all_group_codes:
            group_display_names[c].set(data.get("group_display_names", {}).get(c, c))
        # 載入所有分組的顏色
        for c in all_group_codes:
            group_colors[c].set(data.get("group_colors", {}).get(c, "無色"))
        # 載入自訂顏色
        for c in all_group_codes:
            group_custom_colors[c] = data.get("group_custom_colors", {}).get(c, "")
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
        
        # 載入語言設定（只設定變數，不更新 UI，因為 UI 還沒建立）
        global current_language, lang_map
        saved_lang = data.get("language", "繁體中文")
        if saved_lang in LANG_MAP:
            current_language = saved_lang
            lang_map = LANG_MAP.get(saved_lang, LANG_MAP["繁體中文"])
        
        update_show_labels()
        update_group_name()
        
        # 載入完成後重建按鈕（以反映正確的分組數量）
        try:
            rebuild_group_buttons()
            # 重建後更新按鈕顏色
            update_button_colors()
        except:
            # 如果 rebuild_group_buttons 還未定義（初次載入），忽略錯誤
            pass
    except Exception as e:
        # 只在 log 函數存在時才記錄錯誤
        try:
            log(f"{lang_map['設定檔讀取失敗: ']}{e}")
        except:
            print(f"設定檔讀取失敗: {e}")
# 在所有重要變動時呼叫 save_settings
def on_any_change(*args):
    save_settings()

# 當分組數量變更時，需要重建按鈕
def on_group_count_change(*args):
    save_settings()
    rebuild_group_buttons()

folder_var.trace_add("write", on_any_change)
interval_var.trace_add("write", on_any_change)
active_group_count.trace_add("write", on_group_count_change)  # 追蹤分組數量變更並重建按鈕
# 修正：追蹤所有分組（A-J）的名稱變更，而非只追蹤當前啟用的分組
for c in all_group_codes:
    group_display_names[c].trace_add("write", on_any_change)
    group_colors[c].trace_add("write", on_any_change)  # 追蹤顏色變更
for v in group_hotkeys:
    v.trace_add("write", on_any_change)
for entry, *_ in checkbox_vars_entries:
    entry.bind("<FocusOut>", lambda e: save_settings())
    entry.bind("<KeyRelease>", lambda e: save_settings())
# 修正：追蹤所有 4 個分組變數（var1, var2, var3, var4）
for _, var1, var2, var3, var4, *_ in checkbox_vars_entries:
    var1.trace_add("write", on_any_change)
    var2.trace_add("write", on_any_change)
    var3.trace_add("write", on_any_change)
    var4.trace_add("write", on_any_change)

# 關閉時自動儲存
atexit.register(save_settings)

# 啟動時自動載入
load_settings()

# --- 啟動時延遲0.5秒再讀取設定檔 ---
def delayed_load_settings():
    time.sleep(0.5)
    app.after(0, lambda: [load_settings(), apply_language(current_language), update_file_list(), update_window_list()])
threading.Thread(target=delayed_load_settings, daemon=True).start()

# 綺定分組名稱變動時自動更新（修正：追蹤所有分組 A-J）
for c in all_group_codes:
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
    about_win.title(lang_map["關於 ChroLens_Portal"])
    
    # 根據 DPI 調整對話框大小
    dialog_width = int(450 / dpi_scale)
    dialog_height = int(300 / dpi_scale)
    dialog_width = max(400, dialog_width)  # 最小寬度
    dialog_height = max(280, dialog_height)  # 最小高度
    
    about_win.geometry(f"{dialog_width}x{dialog_height}")
    about_win.resizable(False, False)
    about_win.grab_set()
    # 置中顯示
    app.update_idletasks()
    x = app.winfo_x() + (app.winfo_width() // 2) - (dialog_width // 2)
    y = app.winfo_y() + 80
    about_win.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

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

    # 根據 DPI 調整 padding
    dialog_padding = max(10, int(20 / dpi_scale))
    frm = tb.Frame(about_win, padding=dialog_padding)
    frm.pack(fill="both", expand=True)

    tb.Label(frm, text=f"ChroLens_Portal\n{lang_map['視窗管理工具']}\n{lang_map['分組啟動、快捷切換、智能布局記憶']}", font=("Microsoft JhengHei", max(9, int(11 / dpi_scale)),)).pack(anchor="w", pady=(0, int(6 / dpi_scale)))
    link = tk.Label(frm, text="ChroLens_模擬器討論區", font=("Microsoft JhengHei", max(8, int(10 / dpi_scale)), "underline"), fg="#5865F2", cursor="hand2")
    link.pack(anchor="w")
    link.bind("<Button-1>", lambda e: os.startfile("https://discord.gg/72Kbs4WPPn"))
    github = tk.Label(frm, text="查看更多工具(巴哈)", font=("Microsoft JhengHei", 10, "underline"), fg="#24292f", cursor="hand2")
    github.pack(anchor="w", pady=(8, 0))
    github.bind("<Button-1>", lambda e: os.startfile("https://home.gamer.com.tw/profile/index_creation.php?owner=umiwued&folder=524003"))
    
    # 作者資訊（包含贊助連結）
    author_frame = tb.Frame(frm)
    author_frame.pack(anchor="w", pady=(8, 6))
    tb.Label(author_frame, text="Created by Lucienwooo ", font=("Microsoft JhengHei", 11)).pack(side="left")
    kofi_link = tk.Label(author_frame, text="(好啦給你錢錢)", font=("Microsoft JhengHei", 10, "underline"), fg="#FF5E5B", cursor="hand2")
    kofi_link.pack(side="left")
    kofi_link.bind("<Button-1>", lambda e: os.startfile("https://ko-fi.com/B0B51FBVA8"))
    
    # 按鈕區域（包含語言選擇器）
    button_frame = tb.Frame(frm)
    button_frame.pack(anchor="e", pady=(16, 0))
    
    # 語言選擇器（放在左側，不顯示 "語言:" 標籤）
    lang_var = tk.StringVar(value=current_language)
    lang_combo = tb.Combobox(button_frame, textvariable=lang_var, values=["繁體中文", "日本語", "English"], 
                              state="readonly", width=10)
    lang_combo.pack(side="left", padx=(0, 10))
    
    def on_language_change(event):
        """語言切換事件"""
        new_lang = lang_var.get()
        if new_lang != current_language:
            apply_language(new_lang)
            save_settings()  # 儲存新語言設定
            # 關閉並重新開啟關於對話框以更新所有文字
            about_win.destroy()
            show_about_dialog()
    
    lang_combo.bind("<<ComboboxSelected>>", on_language_change)
    
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
    
    tb.Button(button_frame, text=lang_map["檢查更新"], command=check_for_updates, width=10, bootstyle=INFO).pack(side="left", padx=(0, 5))
    tb.Button(button_frame, text=lang_map["關閉"], command=about_win.destroy, width=8, bootstyle=SECONDARY).pack(side="left")

# --- row 0 新增「刷新視窗」按鈕（SVG 圖示版）---
def manual_refresh_window_list():
    if mini_mode_active:
        log(lang_map["Mini 模式下無需刷新視窗列表"])
        return
    update_window_list()
    log(lang_map["已刷新視窗列表"])

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
    schedule_win.title(lang_map["ChroLens_Portal 排程設定"])
    
    # 根據 DPI 調整對話框大小
    dialog_width = int(600 / dpi_scale)
    dialog_height = int(400 / dpi_scale)
    dialog_width = max(500, dialog_width)  # 最小寬度
    dialog_height = max(350, dialog_height)  # 最小高度
    
    schedule_win.geometry(f"{dialog_width}x{dialog_height}")
    schedule_win.resizable(True, True)
    schedule_win.minsize(500, 350)  # 設定最小尺寸
    
    # 置中顯示
    app.update_idletasks()
    x = app.winfo_x() + (app.winfo_width() // 2) - (dialog_width // 2)
    y = app.winfo_y() + 50
    schedule_win.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
    
    # 設定 icon
    try:
        schedule_win.iconbitmap(ico_path)
    except Exception:
        pass
    
    # 主框架（根據 DPI 調整 padding）
    dialog_padding = max(5, int(10 / dpi_scale))
    main_frame = tb.Frame(schedule_win, padding=dialog_padding)
    main_frame.pack(fill="both", expand=True)
    
    # 響應式配置
    main_frame.grid_rowconfigure(1, weight=1)
    main_frame.grid_columnconfigure(0, weight=1)
    
    # 上方控制區
    control_frame = tb.Frame(main_frame)
    control_frame.grid(row=0, column=0, sticky="ew", pady=(0, dialog_padding))
    
    # 分組選擇
    tb.Label(control_frame, text=lang_map["選擇分組:"]).grid(row=0, column=1, padx=(10, 5), sticky="w")
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
    tb.Label(control_frame, text=lang_map["時間 (HH:MM):"]).grid(row=0, column=3, padx=(10, 5), sticky="w")
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
                log(f"{lang_map['排程已存在：']}{group_display_names[group_code].get()} {lang_map['於']} {time_str}")
                return
        
        schedule_tasks.append({
            "group": group_code,
            "time": time_str,
            "enabled": True
        })
        update_schedule_list()
        save_settings()  # 儲存排程到設定檔
        log(f"{lang_map['已新增排程：']}{group_display_names[group_code].get()} {lang_map['於']} {time_str}")
    
    add_btn = tb.Button(control_frame, text="新增", command=add_schedule, bootstyle="success", width=8)
    add_btn.grid(row=0, column=5, padx=(10, 0), sticky="w")
    
    # 排程列表區域
    list_frame = tb.Frame(main_frame, borderwidth=max(1, int(1 / dpi_scale)), relief="solid")
    list_frame.grid(row=1, column=0, sticky="nsew", pady=(0, dialog_padding))
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
                log(f"{lang_map['排程']} {schedule_tasks[index]['time']} {group_display_names[schedule_tasks[index]['group']].get()} {lang_map['已啟用'] if var.get() else lang_map['已停用']}")
            
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
                log(f"{lang_map['已刪除排程：']}{group_display_names[removed['group']].get()} {lang_map['於']} {removed['time']}")
            
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
            log(lang_map["已開啟 Windows 工作排程器"])
        except Exception as e:
            log(f"{lang_map['無法開啟工作排程器：']}{e}")
    
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
    """從 mini 模式還原到一般模式 - 已不需要，mini mode 改用獨立視窗"""
    global mini_mode_active, mini_window
    if mini_window:
        mini_window.close()

class MiniMode:
    """獨立的 Mini Mode 視窗類別 - 動態響應分組數量"""
    def __init__(self, parent):
        self.parent = parent
        self.win = tb.Toplevel(parent)
        self.win.title("ChroLens Portal - Mini")
        self.win.attributes("-alpha", 0.9)
        self.win.wm_attributes("-topmost", True)
        self.win.protocol("WM_DELETE_WINDOW", self.close)
        
        # 設定 icon 與主程式一致
        try:
            if hasattr(sys, '_MEIPASS'):
                ico_path = os.path.join(sys._MEIPASS, "冥想貓貓.ico")
            else:
                ico_path = "冥想貓貓.ico"
            if os.path.exists(ico_path):
                self.win.iconbitmap(ico_path)
        except Exception as e:
            print(f"Mini 模式無法設定 icon: {e}")
        
        # 可拖曳視窗
        self.offset_x = 0
        self.offset_y = 0
        self.win.bind("<Button-1>", self._start_move)
        self.win.bind("<B1-Motion>", self._do_move)
        
        # 建立主框架，使用最小 padding
        self.frm = tb.Frame(self.win, padding=5)
        self.frm.pack(fill="both", expand=True)
        
        # 初始化按鈕
        self.rebuild_buttons()
    
    def show_desktop(self):
        """顯示桌面（最小化所有視窗）"""
        try:
            # 使用 Windows API 顯示桌面
            import ctypes
            # 模擬 Win+D 快捷鍵
            ctypes.windll.user32.keybd_event(0x5B, 0, 0, 0)  # Win key down
            ctypes.windll.user32.keybd_event(0x44, 0, 0, 0)  # D key down
            ctypes.windll.user32.keybd_event(0x44, 0, 2, 0)  # D key up
            ctypes.windll.user32.keybd_event(0x5B, 0, 2, 0)  # Win key up
            log("已顯示桌面")
        except Exception as e:
            log(f"顯示桌面失敗：{e}")
    
    def rebuild_buttons(self):
        """根據當前分組數量重建 mini 模式按鈕"""
        # 清除所有舊按鈕
        for widget in self.frm.winfo_children():
            widget.destroy()
        
        # 獲取當前分組數量
        count = active_group_count.get()
        current_group_codes = all_group_codes[:count]
        
        # 計算視窗大小（響應式）
        buttons_per_row = 5
        rows_needed = (count + buttons_per_row - 1) // buttons_per_row
        total_rows = rows_needed + 1  # 置頂行 + 啟動/關閉行（同一行）
        
        # 強制固定視窗寬度，不允許調整
        # 寬度需要容納：1個特殊按鈕 + count個置頂按鈕（第一行最寬）
        # 或：1個特殊按鈕 + buttons_per_row個啟動 + buttons_per_row個關閉
        button_width = 70
        # 取第一行和第二行中較寬的那個
        first_row_width = (1 + count) * button_width + 100  # 特殊按鈕 + 所有置頂按鈕
        second_row_width = (1 + buttons_per_row * 2) * button_width + 100  # 特殊按鈕 + 啟動 + 關閉
        mini_fixed_width = max(600, first_row_width, second_row_width)
        total_height = 60 + total_rows * 38  # 減少高度
        self.win.geometry(f"{mini_fixed_width}x{total_height}")
        # 強制固定 mini 視窗寬度
        self.win.minsize(mini_fixed_width, total_height)
        self.win.maxsize(mini_fixed_width, 9999)
        
        # 配置欄位（返回按鈕1欄 + 最多buttons_per_row*2欄用於啟動/關閉按鈕）
        # 取第一行和第二行需要的最大欄位數
        max_cols = max(1 + count, 1 + buttons_per_row * 2)
        for i in range(max_cols):
            self.frm.grid_columnconfigure(i, weight=1, uniform="mini_col")
        
        # 配置行高（3行）
        self.frm.grid_rowconfigure(0, weight=1, uniform="mini_row")  # 置頂行
        self.frm.grid_rowconfigure(1, weight=1, uniform="mini_row")  # 啟動/關閉行1
        self.frm.grid_rowconfigure(2, weight=1, uniform="mini_row")  # 啟動/關閉行2
        
        # === 第 1 行：置頂切換按鈕（平鋪寬度）===
        # 置頂切換按鈕
        for idx, code in enumerate(current_group_codes):
            color_name = group_colors[code].get()
            color_hex = DRACULA_COLORS.get(color_name, "")
            
            if color_hex:
                # 使用自定義樣式
                info_style = f"Group{code}.info.TButton"
                btn = tb.Button(
                    self.frm,
                    text=group_display_names[code].get(),
                    command=lambda c=code: focus_next_in_group(c),
                    style=info_style
                )
            else:
                btn = tb.Button(
                    self.frm,
                    text=group_display_names[code].get(),
                    command=lambda c=code: focus_next_in_group(c),
                    bootstyle="info-outline"
                )
            btn.grid(row=0, column=idx+1, padx=1, pady=1, sticky="nsew")
        
        # === 第 2-3 行：返回按鈕（跨3行，佔1.5行高度）===
        restore_btn = tb.Button(
            self.frm,
            text="↩️",
            command=self.close,
            bootstyle="secondary"
        )
        restore_btn.grid(row=0, column=0, rowspan=2, padx=1, pady=1, sticky="nsew")
        
        # === 第 2-3 行：顯示桌面按鈕（跨3行，佔1.5行高度）===
        show_desktop_btn = tb.Button(
            self.frm,
            text="🖥️",
            command=self.show_desktop,
            bootstyle="warning"
        )
        show_desktop_btn.grid(row=2, column=0, rowspan=2, padx=1, pady=1, sticky="nsew")
        
        # === 第 2-3 行：啟動按鈕（左邊，每行最多 5 個）===
        for idx, code in enumerate(current_group_codes):
            color_name = group_colors[code].get()
            color_hex = DRACULA_COLORS.get(color_name, "")
            
            row = 1 + (idx // buttons_per_row)
            col = (idx % buttons_per_row) + 1
            
            if color_hex:
                # 使用自定義樣式
                success_style = f"Group{code}.success.TButton"
                btn = tb.Button(
                    self.frm,
                    text=group_display_names[code].get(),
                    command=lambda c=code: start_group_opening(c),
                    style=success_style
                )
            else:
                btn = tb.Button(
                    self.frm,
                    text=group_display_names[code].get(),
                    command=lambda c=code: start_group_opening(c),
                    bootstyle="success-outline"
                )
            btn.grid(row=row, column=col, padx=1, pady=1, sticky="nsew")
        
        # === 第 2-3 行：關閉按鈕（右邊，緊接著啟動按鈕）===
        offset_col = 1 + buttons_per_row  # 啟動按鈕右邊
        for idx, code in enumerate(current_group_codes):
            color_name = group_colors[code].get()
            color_hex = DRACULA_COLORS.get(color_name, "")
            
            row = 1 + (idx // buttons_per_row)
            col = offset_col + (idx % buttons_per_row)
            
            if color_hex:
                # 使用自定義樣式
                danger_style = f"Group{code}.danger.TButton"
                btn = tb.Button(
                    self.frm,
                    text=group_display_names[code].get(),
                    command=lambda c=code: close_group_windows(c),
                    style=danger_style
                )
            else:
                btn = tb.Button(
                    self.frm,
                    text=group_display_names[code].get(),
                    command=lambda c=code: close_group_windows(c),
                    bootstyle="danger-outline"
                )
            btn.grid(row=row, column=col, padx=1, pady=1, sticky="nsew")
    
    def _start_move(self, event):
        """開始拖曳"""
        self.offset_x = event.x
        self.offset_y = event.y
    
    def _do_move(self, event):
        """執行拖曳"""
        x = self.win.winfo_x() + (event.x - self.offset_x)
        y = self.win.winfo_y() + (event.y - self.offset_y)
        self.win.geometry(f"+{x}+{y}")
    
    def close(self):
        """關閉 mini 視窗並顯示主程式"""
        self.win.destroy()
        global mini_mode_active, mini_window
        mini_mode_active = False
        mini_window = None
        mini_btn.config(text=lang_map["mini"])
        # 顯示主程式視窗
        app.deiconify()
        log(lang_map["已還原一般模式"])

def toggle_mini_mode():
    """切換 mini 模式 - 隱藏主程式視窗，顯示獨立的 mini 視窗"""
    global mini_mode_active, mini_window
    
    if not mini_mode_active:
        mini_mode_active = True
        mini_window = MiniMode(app)
        mini_btn.config(text=lang_map["還原"])
        # 隱藏主程式視窗
        app.withdraw()
        log(lang_map["已進入 mini 模式"])
    else:
        if mini_window:
            mini_window.close()

# 舊的 mini mode 函數已移除，改用 MiniMode 類別

# --- 排程按鈕 ---
schedule_btn = tb.Button(top_row_frame, text=lang_map["排程"], command=show_schedule_dialog, bootstyle="warning")
schedule_btn.grid(row=0, column=12, padx=(2,2), sticky="e")

mini_btn = tb.Button(top_row_frame, text=lang_map["mini"], command=toggle_mini_mode, bootstyle=INFO)
mini_btn.grid(row=0, column=11, padx=(2,2), sticky="e")

# --- 刷新按鈕 ---
def refresh_lists():
    """刷新檔案列表和視窗列表"""
    update_file_list()
    update_window_list()
    log(lang_map["已刷新檔案和視窗列表"])

refresh_btn = tb.Button(top_row_frame, text="🔄", command=refresh_lists, bootstyle="info", width=3)
refresh_btn.grid(row=0, column=10, padx=(2,2), sticky="e")

# --- 關於按鈕（原本已存在） ---
about_btn = tb.Button(top_row_frame, text=lang_map["關於"], command=show_about_dialog, bootstyle=SECONDARY)
about_btn.grid(row=0, column=9, padx=(4,2), sticky="e")


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
    log(f"{lang_map['[快捷鍵] 觸發分組']} {group_display_names[group_code].get()}")
    
    update_group_hwnd_list(group_code)
    hwnds = group_hwnd_lists[group_code]
    
    if not hwnds:
        log(f"{lang_map['[快捷鍵] 分組']} {group_display_names[group_code].get()} {lang_map['沒有視窗']}")
        return
    
    log(f"{lang_map['[快捷鍵] 開始置頂並恢復分組']} {group_display_names[group_code].get()} {lang_map['的']} {len(hwnds)} {lang_map['個視窗佈局']}")
    
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
                app.after(0, lambda rc=restored_count: log(f"{lang_map['已恢復']} {rc} {lang_map['個視窗的佈局']}"))
            
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
            
            app.after(0, lambda gn=group_display_names[group_code].get(), hwnd_count=len(hwnds): log(f"{lang_map['已完成置頂分組']} {gn} {lang_map['的所有']} {hwnd_count} {lang_map['個視窗']}"))
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
                        log(f"{lang_map['[快捷鍵]']} {hotkey} {lang_map['已觸發 -> 分組']} {group_display_names[group_code].get()}")
                        focus_next_in_group(group_code)
                    except Exception as e:
                        log(f"{lang_map['[快捷鍵] 執行錯誤: ']}{e}")
                return callback
            
            # 註冊快捷鍵（按下時立即觸發）
            handler = keyboard.add_hotkey(
                hotkey_str,
                create_hotkey_callback(code),
                suppress=False
            )
            
            hotkey_handlers.append(handler)
            success_count += 1
            # log(f"✓ {hotkey} → 分組 {group_display_names[code].get()}")  # 隱藏快捷鍵註冊訊息
            
        except Exception as e:
            log(f"✗ 註冊失敗 {hotkey}: {e}")
    
    # if success_count > 0:
    #     log(f"共註冊 {success_count} 個快捷鍵")  # 隱藏快捷鍵註冊訊息
    # else:
    #     log("警告：快捷鍵註冊失敗，請以管理員權限執行")
    
    # 只在失敗時顯示警告
    if success_count == 0:
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
    """解析 .lnk 捷徑檔案，回傳 (目標路徑, 參數字串)
    使用多種方法嘗試解析，確保兼容性"""
    
    # 方法 1: 使用 win32com.client (更穩定的方式)
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)
        target_path = shortcut.Targetpath
        arguments = shortcut.Arguments
        if target_path:
            log(f"{lang_map['[捷徑] 解析成功: ']}{target_path}")
            return target_path, arguments
    except Exception as e:
        log(f"{lang_map['[捷徑] 方法1失敗: ']}{e}")
    
    # 方法 2: 使用 pythoncom (備援)
    try:
        import pythoncom
        from win32com.shell import shell as win32_shell
        
        pythoncom.CoInitialize()
        try:
            shell_link = pythoncom.CoCreateInstance(
                win32_shell.CLSID_ShellLink, None,
                pythoncom.CLSCTX_INPROC_SERVER, win32_shell.IID_IShellLink
            )
            persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
            persist_file.Load(lnk_path)
            target_path, _ = shell_link.GetPath(win32_shell.SLGP_UNCPRIORITY)
            arguments = shell_link.GetArguments()
            if target_path:
                log(f"[捷徑] 解析成功(方法2): {target_path}")
                return target_path, arguments
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        log(f"[捷徑] 方法2失敗: {e}")
    
    # 方法 3: 使用 PowerShell (最可靠的備援方法)
    try:
        import subprocess

# 版本管理模組
try:
    from version_manager import VersionManager
    from version_info_dialog import VersionInfoDialog
    VERSION_MANAGER_AVAILABLE = True
except ImportError:
    print("版本管理模組未安裝，版本檢查功能將停用")
    VERSION_MANAGER_AVAILABLE = False

        ps_cmd = f'''
        $ws = New-Object -ComObject WScript.Shell;
        $shortcut = $ws.CreateShortcut('{lnk_path}');
        Write-Output $shortcut.TargetPath;
        Write-Output "|||";
        Write-Output $shortcut.Arguments
        '''
        result = subprocess.run(
            ['powershell', '-Command', ps_cmd],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            output = result.stdout.strip().split('|||')
            target_path = output[0].strip() if output else ''
            arguments = output[1].strip() if len(output) > 1 else ''
            if target_path and os.path.exists(target_path):
                log(f"[捷徑] 解析成功(PowerShell): {target_path}")
                return target_path, arguments
    except Exception as e:
        log(f"[捷徑] PowerShell方法失敗: {e}")
    
    # 所有方法都失敗
    log(f"[捷徑] 無法解析: {lnk_path}，將使用直接開啟方式")
    return None, None

def open_entry_file(entry):
    file_path = entry.get().strip()
    folder = folder_var.get()
    if not file_path:
        log(lang_map["此欄位無檔案名稱"])
        return
    full_path = os.path.join(folder, file_path)
    if not os.path.exists(full_path):
        log(f"{lang_map['找不到檔案: ']}{full_path}")
        return
    try:
        if full_path.lower().endswith('.lnk'):
            # 解析捷徑
            target, args = open_lnk_target(full_path)
            if target and os.path.exists(target):
                log(f"{lang_map['開啟捷徑目標: ']}{target} {args}")
                # 使用 subprocess 執行目標程式
                if args:
                    subprocess.Popen(f'"{target}" {args}', shell=True)
                else:
                    subprocess.Popen(f'"{target}"', shell=True)
            else:
                # 如果解析失敗，直接用 Windows 開啟捷徑
                log(f"直接開啟捷徑: {full_path}")
                os.startfile(full_path)
        elif full_path.lower().endswith('.exe'):
            # 直接執行 .exe 檔案
            log(f"執行程式: {full_path}")
            subprocess.Popen(f'"{full_path}"', shell=True)
        else:
            # 其他檔案類型用系統預設程式開啟
            log(f"已開啟檔案: {full_path}")
            os.startfile(full_path)
    except Exception as e:
        log(f"開啟檔案失敗: {full_path}，錯誤：{e}")

# --- 延遲初始化：在所有函數定義完成後執行 ---
# 綁定分組數量選擇器事件並初始化按鈕
group_count_combo.bind("<<ComboboxSelected>>", on_group_count_change)
rebuild_topmost_buttons()  # 先初始化置頂切換按鈕
rebuild_group_buttons()    # 再初始化啟動/關閉按鈕

app.mainloop()