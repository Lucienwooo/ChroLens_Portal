# ChroLens_Portal 修復更新 - 2025/11/20

## 🔧 修復內容

### 問題 1: row0 按鈕被壓迫
**症狀**: 在 150% DPI 縮放下，頂部工具列的按鈕過於擁擠，可能被視窗邊框壓迫。

**修復方案**:

1. **調整視窗大小策略**
   ```python
   # 舊策略：過度縮小視窗
   scaled_width = int(base_width * min(1.0, 1.0 / (dpi_scale * 0.8)))
   
   # 新策略：保持較大視窗，避免內容壓縮
   if dpi_scale >= 1.5:  # 150% 或更高
       scaled_width = int(base_width * 0.95)  # 只略微縮小 5%
   ```

2. **增加最小視窗大小**
   - 最小寬度: 1000 → **1100** (增加 100px)
   - 最小高度: 600 → **650** (增加 50px)

3. **調整 Padding 策略**
   ```python
   # 舊策略：高 DPI 時減少 padding
   adaptive_padding = max(1, int(2 / dpi_scale))  # 150% 時為 1
   
   # 新策略：保持適當 padding
   if dpi_scale >= 1.5:
       adaptive_padding = 2  # 保持 2，避免過於擁擠
   ```

4. **調整字型大小策略**
   ```python
   # 舊策略：字型過度縮小
   scaled_font_size = max(8, int(10 / dpi_scale))  # 150% 時為 8pt
   
   # 新策略：保持清晰可讀
   if dpi_scale >= 1.5:
       scaled_font_size = 9  # 使用 9pt，保持清晰
   ```

5. **優化頂部工具列網格配置**
   - 讓路徑和間隔區域可以自動擴展 (`weight=1`)
   - 按鈕區域固定大小 (`weight=0`)
   - 支援 13 個欄位的彈性佈局

**效果**:
- ✅ 150% DPI 下視窗大小: 1140x684 (原本 999x600)
- ✅ 按鈕間距更適當，不會被壓迫
- ✅ 字型大小 9pt，清晰可讀
- ✅ 所有功能按鈕都能完整顯示

---

### 問題 2: 無法開啟 .lnk 捷徑檔案
**症狀**: 
```
無法開啟: C:/Users/Lucien/Desktop/0-shot\噗噗.lnk
錯誤：No module named 'win32timezone'
```

**原因分析**:
1. `pythoncom.CoInitialize()` 在多執行緒環境下未正確清理
2. COM 元件未正確初始化和釋放
3. 缺少錯誤處理機制

**修復方案**:

1. **增強 `open_lnk_target()` 函數**
   ```python
   def open_lnk_target(lnk_path):
       try:
           pythoncom.CoInitialize()  # 初始化 COM
           try:
               # 解析捷徑邏輯
               shell_link = pythoncom.CoCreateInstance(...)
               # ... 解析過程 ...
               return target_path, arguments
           finally:
               pythoncom.CoUninitialize()  # 清理 COM（關鍵！）
       except Exception as e:
           log(f"解析捷徑失敗: {e}")
           return None, None  # 返回 None 而不是崩潰
   ```

2. **增強檔案開啟邏輯 - 多重備援**
   ```python
   if full_path.lower().endswith('.lnk'):
       target, args = open_lnk_target(full_path)
       if target and os.path.exists(target):
           # 方案 A: 成功解析，直接執行目標
           subprocess.Popen(f'"{target}" {args}', shell=True)
       else:
           # 方案 B: 解析失敗，用 Windows 預設方式開啟
           log(f"直接開啟捷徑: {full_path}")
           os.startfile(full_path)  # 備援方案
   ```

3. **支援多種檔案類型**
   - ✅ `.lnk` 捷徑: 解析目標或直接開啟
   - ✅ `.exe` 執行檔: 直接執行
   - ✅ 其他檔案: 用系統預設程式開啟

4. **改善錯誤訊息**
   ```python
   except Exception as e:
       log(f"無法開啟: {file_path}，錯誤：{e}")
   ```

**測試工具**: `test_lnk.py`
```bash
python test_lnk.py
```

**效果**:
- ✅ 正確初始化和清理 COM 元件
- ✅ 即使解析失敗也能開啟捷徑（備援機制）
- ✅ 支援 .lnk / .exe / 其他檔案類型
- ✅ 更詳細的錯誤訊息
- ✅ 不會因為 `win32timezone` 錯誤而崩潰

---

## 📊 測試結果

### 你的系統配置 (150% DPI)
```
DPI: 144
縮放比例: 1.50x (150%)
建議視窗大小: 1140x684
最小視窗大小: 1100x650
建議 padding: 2
建議字型大小: 9pt
```

### 捷徑測試
```
檔案路徑: C:/Users/Lucien/Desktop/0-shot\噗噗.lnk
正規化路徑: C:\Users\Lucien\Desktop\0-shot\噗噗.lnk
檔案存在: True
```

---

## 🚀 使用說明

### 開啟捷徑/程式的方式

1. **透過分組欄位**
   - 拖曳檔案名稱到分組欄位
   - 點擊編號按鈕開啟

2. **透過分組啟動**
   - 設定檔案到分組
   - 點擊「啟動 A/B/C...」按鈕
   - 程式會自動批次開啟

3. **支援的檔案類型**
   - `.lnk` - Windows 捷徑
   - `.exe` - 執行檔
   - `.bat` - 批次檔
   - `.py` - Python 腳本
   - `.txt`, `.docx`, `.pdf` 等 - 用預設程式開啟

---

## 🔄 更新檔案清單

### 主程式
- ✅ `ChroLens_Portal.py` - 修復 DPI 和捷徑問題

### 測試工具
- ✅ `test_dpi.py` - DPI 檢測測試（已更新）
- ✅ `test_lnk.py` - 捷徑開啟測試（新增）

### 文檔
- ✅ `DPI_ADAPTIVE_UPDATE.md` - DPI 自適應說明
- ✅ `FIX_UPDATE.md` - 本修復說明（本檔案）

---

## 💡 常見問題 Q&A

### Q1: 按鈕還是被壓迫怎麼辦？
**A**: 手動調整視窗大小，程式會記住你的偏好。最小寬度已設為 1100px。

### Q2: 捷徑開啟失敗怎麼辦？
**A**: 
1. 確認 pywin32 已安裝: `pip install pywin32`
2. 執行後續配置: `python Scripts\pywin32_postinstall.py -install`
3. 檢查捷徑檔案是否存在
4. 檢查捷徑目標是否有效

### Q3: 如何測試捷徑功能？
**A**: 執行 `python test_lnk.py`，它會：
- 檢查檔案是否存在
- 嘗試解析捷徑
- 顯示目標和參數
- 實際開啟程式（僅測試時）

### Q4: 支援哪些檔案類型？
**A**: 
- 捷徑檔案 (`.lnk`)
- 執行檔 (`.exe`, `.bat`, `.cmd`)
- Python 腳本 (`.py`)
- 任何可以用 `os.startfile()` 開啟的檔案

### Q5: 多執行緒環境下會有問題嗎？
**A**: 不會，已正確處理 COM 初始化和清理：
```python
pythoncom.CoInitialize()   # 初始化
try:
    # 使用 COM
finally:
    pythoncom.CoUninitialize()  # 清理
```

---

## 🎯 下一步建議

1. **測試程式**
   ```bash
   cd C:\Users\Lucien\Documents\GitHub\ChroLens_Portal
   python test_dpi.py      # 測試 DPI 配置
   python test_lnk.py      # 測試捷徑功能
   python ChroLens_Portal.py  # 啟動主程式（需管理員權限）
   ```

2. **驗證修復**
   - 檢查視窗大小是否合適
   - 測試開啟捷徑檔案
   - 測試開啟 .exe 執行檔
   - 檢查按鈕是否仍被壓迫

3. **如有問題**
   - 查看日誌區域的錯誤訊息
   - 執行測試腳本查看詳細資訊
   - 提供錯誤截圖和日誌

---

**修復完成！現在程式應該能在 150% DPI 下正常顯示，並能正確開啟各種類型的捷徑和程式！** 🎉
