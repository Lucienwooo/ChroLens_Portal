# ChroLens_Portal 2.3

> 視窗管理工具 - 分組啟動、快捷切換、布局管理

![ChroLens_Portal_Basic_Operations](clp2.2.png)
[[ChroLens_Portal_基本操作]](https://player.vimeo.com/video/1087659485?h=83487a7ea9)

---

## ✨ 核心功能

### 🚀 分組啟動
將常用的程式、文件或捷徑分組，一鍵啟動所有應用，支援 `.lnk` 捷徑檔案。

### 🔄 快捷切換  
預設熱鍵 `Alt+1~4/Q/W` 快速切換分組視窗，提升工作效率。

### 🎯 視窗管理
一鍵關閉分組視窗，保持桌面整潔。

### 💾 布局管理 (v2.3 新增)
儲存並還原每個分組的視窗位置與大小，快速恢復工作環境。

### 🔐 管理員權限 (v2.3 新增)
自動請求管理員權限，確保所有功能正常運作。

---

## 📦 快速開始

### 安裝依賴
```bash
pip install -r requirements.txt
```

### 啟動程式
```bash
python main.py
# 或雙擊 啟動.bat
```

### 打包成 EXE
```bash
# 使用提供的打包腳本
build.bat
```

---

## 🎮 使用場景

### 工作模式
- **分組 A**: VSCode + Chrome + Terminal  
  → 儲存布局：左側編輯器、右側瀏覽器、底部終端

### 遊戲模式  
- **分組 B**: 遊戲 + Discord + 攻略網站  
  → 儲存布局：全螢幕遊戲、右上 Discord、右下攻略

### 學習模式
- **分組 C**: Word + PDF閱讀器 + 瀏覽器  
  → 儲存布局：左側筆記、中間教材、右側搜尋

---

## 🔧 使用方式

1. **選擇資料夾** - 點擊「選擇開啟路徑」設定檔案來源
2. **新增檔案** - 從清單拖曳檔案或視窗標題到分組欄位  
3. **選擇分組** - 每個檔案可選擇 A~F 任一分組（可多選）
4. **啟動分組** - 點擊「啟動 X」按鈕或使用熱鍵
5. **關閉分組** - 點擊「關閉 X」按鈕

---

## 📁 專案結構

```
ChroLens_Portal/
├── main.py              # 程式入口（含管理員權限檢查）
├── ui/                  # 使用者介面模組
│   ├── main_window.py   # 主視窗
│   └── file_list.py     # 檔案/視窗列表
├── core/                # 核心功能模組
│   ├── window_manager.py    # 視窗管理 + 布局管理
│   ├── hotkey_handler.py    # 全域熱鍵
│   └── file_opener.py       # 檔案開啟
├── utils/               # 工具模組
│   ├── config.py        # 設定檔管理
│   └── logger.py        # 日誌系統
└── requirements.txt     # 依賴套件
```

詳細架構請參閱 `ARCHITECTURE.txt`

---

## ⚙️ 設定檔

- `chrolens_portal.json` - 主要設定（路徑、分組、熱鍵）
- `window_layouts.json` - 視窗布局資料

---

## 📝 更新日誌

### v2.3 (2025/01/11)
- ✨ 模組化架構重構（拆分為 ui/core/utils 模組）
- ✨ 新增管理員權限檢查（自動請求 UAC 提升）
- ✨ 新增視窗布局管理（儲存/還原視窗位置）
- ⚡ 優化視窗切換效能（減少閃爍）
- 📚 完整文件與測試腳本

### v2.2 (2025/05/26)  
- 🎉 首次發布
- 分組啟動/關閉功能
- 熱鍵切換支援

---

## 💬 支援與反饋

- **Discord**: [https://discord.gg/72Kbs4WPPn](https://discord.gg/72Kbs4WPPn)
- **巴哈姆特**: [https://home.gamer.com.tw/profile/index_creation.php?owner=umiwued&folder=523848](https://home.gamer.com.tw/profile/index_creation.php?owner=umiwued&folder=523848)

### 💸 支持作者 / Support the Creator / 作者を応援する💸
[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/B0B51FBVA8)

**這些程式幫你省下的時間，分一點來抖內吧！給我錢錢！**  
**These scripts saved you time—share a bit and donate. Give me money!**  
**このツールで浮いた時間、ちょっとだけ投げ銭して？お金ちょうだい！**

---

## 📄 授權資訊

詳見 [LICENSE](LICENSE)

---

## 🔗 ChroLens 系列專案

- [ChroLens_Portal](https://github.com/Lucienwooo/ChroLens_Portal) - 視窗管理工具
- [ChroLens_Mimic](https://github.com/Lucienwooo/ChroLens_Mimic) - 巨集錄製工具  
- [ChroLens_Clear](https://github.com/Lucienwooo/ChroLens_Clear) - 批次關閉工具
- [ChroLens_Orbit](https://github.com/Lucienwooo/ChroLens_Orbit) - 排程工具
