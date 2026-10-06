# -*- coding: utf-8 -*-
"""
版本資訊對話視窗 - ChroLens_Portal
顯示當前版本、更新日誌和檢查更新

作者: Lucien
"""

import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox, scrolledtext
import threading
import os
import sys
import webbrowser

try:
    from utils import get_icon_path, set_window_icon
except ImportError:
    def get_icon_path():
        if getattr(sys, 'frozen', False):
            return os.path.join(sys._MEIPASS, "冥想貓貓.ico")
        return "冥想貓貓.ico"
        
    def set_window_icon(w, icon_path=None):
        try:
            ico = icon_path or get_icon_path()
            if os.path.exists(ico):
                w.iconbitmap(ico)
        except:
            pass


class VersionInfoDialog(tb.Toplevel):
    """版本資訊與更新對話視窗（合併版）"""
    
    def __init__(
        self,
        parent,
        version_manager,
        current_version,
        on_update_callback=None,
        on_restore_defaults=None,
        on_language_change=None,
        current_language="繁體中文",
        initial_update_info=None,
        initial_tab="about",
        icon_path=None,
    ):
        """
        初始化版本資訊對話視窗
        
        Args:
            parent: 父視窗
            version_manager: 版本管理器實例
            current_version: 當前版本號
            on_update_callback: 更新完成後的回調函數
        """
        super().__init__(parent)
        
        self.parent = parent
        self.version_manager = version_manager
        self.current_version = current_version
        self.on_update_callback = on_update_callback
        self.on_restore_defaults = on_restore_defaults
        self.on_language_change = on_language_change
        self.current_language = current_language
        self.initial_update_info = initial_update_info
        self.initial_tab = initial_tab
        self.icon_path = icon_path

        self.title("關於")
        self.geometry("820x835")
        self.resizable(True, True)
        self.minsize(700, 675)
        
        # 設定圖示
        set_window_icon(self, self.icon_path)
        
        # 置中顯示
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")
        
        # 設定為模態對話視窗
        self.transient(parent)
        self.grab_set()
        
        # 創建 UI
        self._create_ui()
        if self.initial_tab == "version":
            self.show_version_tab()
        
        if initial_update_info is not None:
            self.after(100, lambda: self._update_version_status(initial_update_info))
        else:
            self.after(100, self._load_content)
    
    def _create_ui(self):
        """創建 UI 元件"""
        # 主框架
        main_frame = tb.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=YES)

        self.notebook = tb.Notebook(main_frame)
        self.notebook.pack(fill=BOTH, expand=YES)
        self.about_tab = tb.Frame(self.notebook, padding=18)
        self.version_tab = tb.Frame(self.notebook, padding=15)
        self.notebook.add(self.about_tab, text="關於")
        self.notebook.add(self.version_tab, text="版本資訊")
        self._create_about_tab()
        
        # ========== 標題區域 ==========
        title_frame = tb.Frame(self.version_tab)
        title_frame.pack(fill=X, pady=(0, 15))
        
        title_label = tb.Label(
            title_frame,
            text="ChroLens_Portal",
            font=("Microsoft JhengHei", 18, "bold"),
            bootstyle="primary"
        )
        title_label.pack(side=LEFT)
        
        # ========== 版本資訊區域 ==========
        info_frame = tb.Labelframe(
            self.version_tab,
            text="版本資訊",
            padding=15,
            bootstyle="info"
        )
        info_frame.pack(fill=X, pady=(0, 15))
        
        # 當前版本
        current_version_frame = tb.Frame(info_frame)
        current_version_frame.pack(fill=X, pady=5)
        
        tb.Label(
            current_version_frame,
            text="目前版本：",
            font=("Microsoft JhengHei", 11),
            width=12,
            anchor=W
        ).pack(side=LEFT)
        
        tb.Label(
            current_version_frame,
            text=f"v{self.current_version}",
            font=("Microsoft JhengHei", 11, "bold"),
            bootstyle="info"
        ).pack(side=LEFT)
        
        # 最新版本
        latest_version_frame = tb.Frame(info_frame)
        latest_version_frame.pack(fill=X, pady=5)
        
        tb.Label(
            latest_version_frame,
            text="最新版本：",
            font=("Microsoft JhengHei", 11),
            width=12,
            anchor=W
        ).pack(side=LEFT)
        
        self.latest_version_label = tb.Label(
            latest_version_frame,
            text="檢查中...",
            font=("Microsoft JhengHei", 11),
            bootstyle="secondary"
        )
        self.latest_version_label.pack(side=LEFT)
        
        # 更新狀態
        status_frame = tb.Frame(info_frame)
        status_frame.pack(fill=X, pady=5)
        
        tb.Label(
            status_frame,
            text="更新狀態：",
            font=("Microsoft JhengHei", 11),
            width=12,
            anchor=W
        ).pack(side=LEFT)
        
        self.update_status_label = tb.Label(
            status_frame,
            text="正在檢查更新...",
            font=("Microsoft JhengHei", 11),
            bootstyle="secondary"
        )
        self.update_status_label.pack(side=LEFT)
        
        # ========== 更新說明區域 ==========
        update_notes_frame = tb.Labelframe(
            self.version_tab,
            text="更新說明",
            padding=15,
            bootstyle="success"
        )
        update_notes_frame.pack(fill=BOTH, expand=YES, pady=(0, 10))
        
        # 更新說明文字框
        self.update_notes_text = scrolledtext.ScrolledText(
            update_notes_frame,
            wrap=tk.WORD,
            font=("Microsoft JhengHei", 9),
            height=8
        )
        self.update_notes_text.pack(fill=BOTH, expand=YES)
        self.update_notes_text.insert("1.0", "檢查中...")
        self.update_notes_text.config(state=tk.DISABLED)
        
        # ========== 進度區域（初始隱藏）==========
        self.progress_frame = tb.Labelframe(
            self.version_tab,
            text="更新進度",
            padding=10,
            bootstyle="info"
        )
        
        self.progress_label = tb.Label(
            self.progress_frame,
            text="準備下載...",
            font=("Microsoft JhengHei", 10)
        )
        self.progress_label.pack(pady=(0, 5))
        
        self.progress_bar = tb.Progressbar(
            self.progress_frame,
            mode='determinate',
            length=500,
            bootstyle="success"
        )
        self.progress_bar.pack(pady=(0, 5))
        
        self.progress_detail_label = tb.Label(
            self.progress_frame,
            text="",
            font=("Microsoft JhengHei", 9),
            bootstyle="secondary"
        )
        self.progress_detail_label.pack()
        
        # ========== 按鈕區域 ==========
        button_frame = tb.Frame(self.version_tab)
        button_frame.pack(fill=X, pady=(10, 0))
        self.version_button_frame = button_frame
        
        # 更新按鈕
        self.update_btn = tb.Button(
            button_frame,
            text="立即更新",
            bootstyle="success",
            width=15,
            command=self._start_update,
            state=DISABLED
        )
        self.update_btn.pack(side=LEFT, padx=5)
        
        # 關閉按鈕
        close_btn = tb.Button(
            button_frame,
            text="關閉",
            bootstyle="secondary",
            width=15,
            command=self.destroy
        )
        close_btn.pack(side=RIGHT, padx=5)

    def _create_about_tab(self):
        title = tb.Label(
            self.about_tab,
            text="ChroLens_Portal",
            font=("Microsoft JhengHei", 20, "bold"),
            bootstyle="primary",
        )
        title.pack(anchor=W, pady=(0, 8))
        tb.Label(
            self.about_tab,
            text="視窗管理工具\n分組啟動、快捷切換、智能布局記憶",
            font=("Microsoft JhengHei", 11),
            justify=LEFT,
        ).pack(anchor=W, pady=(0, 12))
        tb.Label(self.about_tab, text="By Lucienwooo").pack(anchor=W)

        links = tb.Frame(self.about_tab)
        links.pack(fill=X, pady=(10, 0))
        for label, url in (
            ("Discord 支援伺服器", "https://discord.gg/72Kbs4WPPn"),
            ("更多工具（巴哈）", "https://home.gamer.com.tw/profile/index_creation.php?owner=umiwued&folder=524003"),
        ):
            link = tb.Label(links, text=label, bootstyle="info", cursor="hand2")
            link.pack(anchor=W, pady=2)
            link.bind("<Button-1>", lambda event, target=url: webbrowser.open(target))

        qr_frame = tb.Labelframe(
            self.about_tab,
            text="贊助",
            padding=8,
            bootstyle="secondary",
        )
        qr_frame.pack(fill=BOTH, expand=YES, pady=(14, 10))
        tb.Label(qr_frame, text="若不嫌棄，歡迎請我喝飲料（支援 LINE Pay）").pack(pady=(0, 6))
        resource_dirs = [
            getattr(sys, "_MEIPASS", ""),
            os.path.dirname(sys.executable),
            os.path.dirname(os.path.abspath(__file__)),
        ]
        qr_path = next(
            (
                os.path.join(directory, "images", "qr_code.png")
                for directory in resource_dirs
                if directory and os.path.exists(os.path.join(directory, "images", "qr_code.png"))
            ),
            "",
        )
        if os.path.exists(qr_path):
            try:
                self.qr_image = tk.PhotoImage(file=qr_path)
                while self.qr_image.width() > 320 or self.qr_image.height() > 320:
                    self.qr_image = self.qr_image.subsample(2, 2)
                tb.Label(qr_frame, image=self.qr_image).pack(expand=YES)
            except tk.TclError as error:
                tb.Label(qr_frame, text=f"贊助圖片載入失敗：{error}").pack(expand=YES)
        else:
            tb.Label(qr_frame, text="未提供贊助圖片").pack(expand=YES)

        controls = tb.Frame(self.about_tab)
        controls.pack(fill=X, pady=(4, 0))
        self.language_var = tk.StringVar(value=self.current_language)
        language_combo = tb.Combobox(
            controls,
            textvariable=self.language_var,
            values=["繁體中文", "日本語", "English"],
            state="readonly",
            width=12,
        )
        language_combo.pack(side=LEFT, padx=(0, 8))
        language_combo.bind("<<ComboboxSelected>>", self._change_language)

        if self.on_restore_defaults:
            tb.Button(
                controls,
                text="恢復預設",
                command=self.on_restore_defaults,
                bootstyle="danger",
            ).pack(side=LEFT, padx=(0, 8))
        tb.Button(
            controls,
            text="版本資訊與更新",
            command=self.show_version_tab,
            bootstyle="info",
        ).pack(side=LEFT)
        tb.Button(
            controls,
            text="關閉",
            command=self.destroy,
            bootstyle="secondary",
        ).pack(side=RIGHT)

    def _change_language(self, event=None):
        if self.on_language_change:
            self.on_language_change(self.language_var.get())

    def show_version_tab(self):
        self.notebook.select(self.version_tab)

    def show_about_tab(self):
        self.notebook.select(self.about_tab)
    
    def _load_content(self):
        """載入內容"""
        threading.Thread(target=self._fetch_data, daemon=True).start()
    
    def _perform_update(self):
        """執行更新流程"""
        try:
            # 1. 下載更新
            self.after(0, lambda: self.progress_label.config(text="階段 1/3: 下載更新檔案"))
            self.after(0, lambda: self.progress_detail_label.config(text="正在連接伺服器..."))
            
            def progress_callback(downloaded, total):
                if total > 0:
                    download_percent = (downloaded / total) * 40
                    self.after(0, lambda p=download_percent: self.progress_bar.config(value=p))
                    
                    downloaded_mb = downloaded / (1024 * 1024)
                    total_mb = total / (1024 * 1024)
                    
                    self.after(0, lambda d=downloaded_mb, t=total_mb: 
                              self.progress_detail_label.config(
                                  text=f"已下載: {d:.2f}MB / {t:.2f}MB ({(d/t)*100:.1f}%)"
                              ))
            
            zip_path = self.version_manager.download_update(
                self.update_info['download_url'],
                progress_callback
            )
            
            if not zip_path:
                raise Exception("下載更新檔案失敗，請檢查網路連線")
            
            # 2. 解壓縮
            self.after(0, lambda: self.progress_bar.config(value=40))
            self.after(0, lambda: self.progress_label.config(text="階段 2/3: 解壓縮檔案"))
            self.after(0, lambda: self.progress_detail_label.config(text="正在解壓縮更新檔案..."))
            
            for i in range(40, 70, 5):
                self.after(0, lambda p=i: self.progress_bar.config(value=p))
                import time
                time.sleep(0.1)
            
            extract_dir = self.version_manager.extract_update(zip_path)
            
            if not extract_dir:
                raise Exception("解壓縮更新檔案失敗，壓縮檔可能已毀損")
            
            self.after(0, lambda: self.progress_bar.config(value=70))
            
            # 3. 應用更新
            self.after(0, lambda: self.progress_label.config(text="階段 3/3: 安裝更新"))
            self.after(0, lambda: self.progress_detail_label.config(text="正在準備更新與安全性備份..."))
            self.after(0, lambda: self.progress_bar.config(value=80))
            
            success = self.version_manager.apply_update(extract_dir, restart_after=True)
            
            if success:
                self.after(0, lambda: self.progress_bar.config(value=100))
                self.after(0, lambda: self.progress_detail_label.config(text="更新安裝程式已啟動!"))
                self.after(0, self._show_success)
            else:
                raise Exception("無法套用更新，寫入更新腳本失敗")
                
        except Exception as e:
            self.after(0, lambda err=str(e): self._show_error(err))
    
    def _show_success(self):
        """顯示更新成功"""
        self.progress_bar.config(value=100)
        self.progress_label.config(text="✓ 更新已就緒!")
        self.progress_detail_label.config(text="安全性備份已完成，準備安裝...")
        
        result = messagebox.askyesno(
            "更新完成",
            "ChroLens_Portal 新版本已成功下載並備份完畢！\n\n是否立即關閉程式並啟動更新覆寫？\n\n(這將自動彈出系統管理員 UAC 要求，更新完成後將自動重新啟動)",
            parent=self
        )
        
        if result:
            if self.on_update_callback:
                self.on_update_callback()
            self.destroy()
        else:
            messagebox.showinfo(
                "提示",
                "已取消立即更新，您可在下次重啟程式時手動點選關於進行更新。",
                parent=self
            )
            self.destroy()
    
    def _show_error(self, error_msg: str):
        """顯示更新失敗"""
        try:
            self.progress_bar.stop()
        except:
            pass
        
        self.progress_bar.config(bootstyle="danger")
        self.progress_label.config(text="✗ 更新失敗")
        self.progress_detail_label.config(text=f"錯誤: {error_msg}")
        self.update_btn.config(state=NORMAL)
        
        messagebox.showerror(
            "更新失敗",
            f"更新過程中發生錯誤：\n\n{error_msg}\n\n請稍後再試，或前往 GitHub Releases 頁面手動下載新版本。",
            parent=self
        )
    
    def _fetch_data(self):
        """獲取更新資訊和檢查更新"""
        try:
            update_info = self.version_manager.check_for_updates()
        except Exception as error:
            self.after(0, lambda message=str(error): self._show_check_error(message))
            return
        self.after(0, lambda: self._update_version_status(update_info))

    def _show_check_error(self, error_message):
        if not self.winfo_exists():
            return
        self.latest_version_label.config(text="無法取得版本資訊", bootstyle="danger")
        self.update_status_label.config(text="更新檢查失敗", bootstyle="danger")
        self.update_notes_text.config(state=tk.NORMAL)
        self.update_notes_text.delete("1.0", tk.END)
        self.update_notes_text.insert("1.0", f"檢查更新時發生錯誤：\n{error_message}")
        self.update_notes_text.config(state=tk.DISABLED)
        self.update_btn.config(state=DISABLED)
    
    def _update_version_status(self, update_info):
        """更新版本狀態和更新說明"""
        if update_info:
            latest_version = update_info['version']
            self.latest_version_label.config(text=f"v{latest_version}", bootstyle="success")
            self.update_status_label.config(text="發現新版本可用！", bootstyle="success")
            self.update_btn.config(state=NORMAL)
            self.update_info = update_info
            
            release_notes = update_info.get('release_notes', '無更新說明')
            self.update_notes_text.config(state=tk.NORMAL)
            self.update_notes_text.delete("1.0", tk.END)
            self.update_notes_text.insert("1.0", release_notes)
            self.update_notes_text.config(state=tk.DISABLED)
        else:
            self.latest_version_label.config(text=f"v{self.current_version}", bootstyle="info")
            self.update_status_label.config(text="目前已是最新版本", bootstyle="info")
            self.update_btn.config(state=DISABLED)
            self.update_info = None
            
            self.update_notes_text.config(state=tk.NORMAL)
            self.update_notes_text.delete("1.0", tk.END)
            self.update_notes_text.insert("1.0", "您目前使用的是最新版本，無需更新。")
            self.update_notes_text.config(state=tk.DISABLED)
    
    def _start_update(self):
        """開始更新流程"""
        if not hasattr(self, 'update_info') or not self.update_info:
            return
        
        result = messagebox.askyesno(
            "確認更新",
            f"即將更新到 v{self.update_info['version']}\n\n更新將會創建系統安全備份與秒還原防護。\n\n確定要啟動更新下載嗎？",
            parent=self
        )
        
        if not result:
            return
        
        self.update_btn.config(state=DISABLED)
        
        # 顯示進度框
        self.progress_frame.pack(
            fill=X,
            pady=(0, 10),
            before=self.version_button_frame,
        )
        
        self.progress_bar.config(value=0, mode='determinate')
        self.progress_label.config(text="準備下載...")
        self.progress_detail_label.config(text="")
        
        threading.Thread(target=self._perform_update, daemon=True).start()
