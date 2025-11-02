"""測試所有模組是否正常載入"""
import sys
import os

def test_imports():
    """測試模組匯入"""
    print("=" * 50)
    print("ChroLens_Portal 2.3 模組測試")
    print("=" * 50)
    
    tests = []
    
    # 測試 utils 模組
    print("\n[1/3] 測試 utils 模組...")
    try:
        from utils import ConfigManager, Logger
        print("  ✓ ConfigManager")
        print("  ✓ Logger")
        tests.append(True)
    except Exception as e:
        print(f"  ✗ 錯誤: {e}")
        tests.append(False)
    
    # 測試 core 模組
    print("\n[2/3] 測試 core 模組...")
    try:
        from core import WindowManager, HotkeyHandler, FileOpener
        print("  ✓ WindowManager")
        print("  ✓ HotkeyHandler")
        print("  ✓ FileOpener")
        tests.append(True)
    except Exception as e:
        print(f"  ✗ 錯誤: {e}")
        tests.append(False)
    
    # 測試 ui 模組
    print("\n[3/3] 測試 ui 模組...")
    try:
        from ui import MainWindow
        from ui.file_list import FileListPanel, WindowListPanel
        print("  ✓ MainWindow")
        print("  ✓ FileListPanel")
        print("  ✓ WindowListPanel")
        tests.append(True)
    except Exception as e:
        print(f"  ✗ 錯誤: {e}")
        tests.append(False)
    
    # 總結
    print("\n" + "=" * 50)
    passed = sum(tests)
    total = len(tests)
    
    if passed == total:
        print(f"✓ 所有測試通過 ({passed}/{total})")
        print("模組載入成功！可以正常執行程式。")
        return True
    else:
        print(f"✗ 部分測試失敗 ({passed}/{total})")
        print("請檢查錯誤訊息並修正問題。")
        return False

def test_dependencies():
    """測試依賴套件"""
    print("\n" + "=" * 50)
    print("依賴套件檢查")
    print("=" * 50)
    
    deps = [
        ('ttkbootstrap', 'ttkbootstrap'),
        ('pywin32', 'win32gui'),
        ('keyboard', 'keyboard'),
    ]
    
    all_ok = True
    for name, import_name in deps:
        try:
            __import__(import_name)
            print(f"  ✓ {name}")
        except ImportError:
            print(f"  ✗ {name} (未安裝)")
            all_ok = False
    
    if not all_ok:
        print("\n請執行以下指令安裝缺少的套件：")
        print("pip install -r requirements.txt")
    
    return all_ok

if __name__ == "__main__":
    print("開始測試 ChroLens_Portal 2.3...\n")
    
    deps_ok = test_dependencies()
    
    if deps_ok:
        modules_ok = test_imports()
        
        if modules_ok:
            print("\n✓ 所有檢查通過！")
            print("\n可以執行以下指令啟動程式：")
            print("python main.py")
        else:
            print("\n✗ 模組載入失敗，請檢查錯誤訊息")
    else:
        print("\n✗ 依賴套件不完整，請先安裝缺少的套件")
    
    print("\n" + "=" * 50)
    input("按 Enter 鍵退出...")
