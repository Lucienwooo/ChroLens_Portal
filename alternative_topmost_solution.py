"""
方案2：批次同時置頂的實作方式
這個檔案展示如何使用 HWND_TOPMOST 來同時置頂所有分組視窗

使用方式：將此函數替換主程式中的 focus_next_in_group 函數
"""

import win32gui
import win32con
import time

def focus_next_in_group_batch(group_code):
    """批次同時置頂分組中的所有視窗（方案2）"""
    update_group_hwnd_list(group_code)
    hwnds = group_hwnd_lists[group_code]
    if not hwnds:
        log(f"分組 {group_display_names[group_code].get()} 沒有視窗")
        return
    
    log(f"開始置頂分組 {group_display_names[group_code].get()} 的 {len(hwnds)} 個視窗")
    
    # 第一步：還原所有最小化的視窗
    for hwnd in hwnds:
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        except Exception:
            pass
    
    # 第二步：將所有視窗設為永久置頂（TOPMOST）
    for hwnd in hwnds:
        try:
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
            )
        except Exception as e:
            log(f"設定 TOPMOST 失敗: {e}")
    
    # 第三步：立即取消永久置頂狀態（但視窗會保持在最上層）
    for hwnd in hwnds:
        try:
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
            )
        except Exception as e:
            log(f"取消 TOPMOST 失敗: {e}")
    
    # 第四步：將第一個視窗設為前景
    if hwnds:
        try:
            win32gui.SetForegroundWindow(hwnds[0])
        except Exception:
            pass
    
    log(f"已完成置頂分組 {group_display_names[group_code].get()} 的所有 {len(hwnds)} 個視窗")


# 方案3：結合兩者的優點 - 使用執行緒 + TOPMOST
def focus_next_in_group_hybrid(group_code):
    """混合方案：使用執行緒執行批次置頂"""
    update_group_hwnd_list(group_code)
    hwnds = group_hwnd_lists[group_code]
    if not hwnds:
        log(f"分組 {group_display_names[group_code].get()} 沒有視窗")
        return
    
    log(f"開始置頂分組 {group_display_names[group_code].get()} 的 {len(hwnds)} 個視窗")
    
    def topmost_windows():
        try:
            # 還原最小化視窗
            for hwnd in hwnds:
                try:
                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                except Exception:
                    pass
            
            # 設為永久置頂
            for hwnd in hwnds:
                try:
                    win32gui.SetWindowPos(
                        hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
                    )
                except Exception:
                    pass
            
            # 短暫延遲確保置頂完成
            time.sleep(0.05)
            
            # 取消永久置頂
            for hwnd in hwnds:
                try:
                    win32gui.SetWindowPos(
                        hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
                    )
                except Exception:
                    pass
            
            # 設為前景視窗
            if hwnds:
                try:
                    win32gui.SetForegroundWindow(hwnds[0])
                except Exception:
                    pass
            
            app.after(0, lambda: log(f"已完成置頂分組 {group_display_names[group_code].get()} 的所有 {len(hwnds)} 個視窗"))
        except Exception as e:
            app.after(0, lambda msg=f"置頂處理異常: {e}": log(msg))
    
    threading.Thread(target=topmost_windows, daemon=True).start()
