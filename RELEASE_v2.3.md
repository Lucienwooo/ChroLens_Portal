# ChroLens_Portal 2.3 發布說明

## 📦 打包完成檔案

### 1. **ChroLens_Portal2.3_Release.zip** (26.52 MB) ⭐推薦
   - 完整發布版本
   - 包含：
     - ChroLens_Portal2.3.exe (主程式)
     - _internal/ (相依檔案)
     - 版本說明_v2.3.txt

### 2. **ChroLens_Portal2.3.exe**
   - 位於 `dist\ChroLens_Portal2.3\` 資料夾內
   - 需要與 _internal 資料夾一起執行

### 3. **ChroLens_Portal2.3.py**
   - Python 原始碼檔案
   - 開發者可用此檔案修改或學習

---

## 🆕 版本 2.3 更新內容

### 主要改進
✅ **修復分組置頂功能**
- 按下快捷鍵時，分組內所有視窗會**同時置頂**
- 不再需要多次按壓快捷鍵才能依序置頂每個視窗
- 使用背景執行緒處理，不會阻塞使用者介面

### 技術細節
- 採用 `HWND_TOPMOST` + `HWND_NOTOPMOST` 技術
- 優化置頂流程：
  1. 還原所有最小化視窗
  2. 批次設為永久置頂
  3. 延遲 0.05 秒確保完成
  4. 取消永久置頂狀態（但視窗保持在最上層）
  5. 將第一個視窗設為前景

---

## 📋 版本歷史

### v2.3 (2025/11/12)
- 修復分組置頂功能，所有視窗同時置頂
- 優化執行緒處理機制

### v2.2 (2025/05/26)
- 初始版本功能

---

## 🚀 使用方式

1. 解壓縮 `ChroLens_Portal2.3_Release.zip`
2. 執行 `ChroLens_Portal2.3.exe`（需要管理員權限）
3. 選擇檔案資料夾
4. 設定檔案分組
5. 使用快捷鍵置頂視窗：
   - Alt+1, Alt+2, Alt+3, Alt+4
   - Alt+Q, Alt+W

---

## 📁 檔案結構

```
ChroLens_Portal2.3/
├── ChroLens_Portal2.3.exe          # 主程式
├── 版本說明_v2.3.txt               # 版本說明
└── _internal/                      # 相依檔案資料夾
    ├── python310.dll
    ├── base_library.zip
    └── ... (其他相依檔案)
```

---

## ⚙️ 系統需求

- **作業系統**: Windows 10/11
- **權限**: 需要管理員權限（UAC）
- **記憶體**: 建議 4GB 以上
- **磁碟空間**: 約 100MB

---

## 📞 聯絡資訊

- **Discord**: ChroLens_模擬器討論區 (https://discord.gg/72Kbs4WPPn)
- **巴哈姆特**: https://home.gamer.com.tw/profile/index_creation.php?owner=umiwued&folder=523848
- **作者**: Lucienwooo

---

## 📄 授權資訊

請參閱 LICENSE 檔案

---

**感謝使用 ChroLens_Portal！**
