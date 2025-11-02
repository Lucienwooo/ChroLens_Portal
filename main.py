"""ChroLens_Portal 主程式入口
強制以管理員權限執行
"""
import sys
import os
import ctypes

def is_admin():
    """檢查是否以管理員權限執行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """請求管理員權限重新啟動程式"""
    if sys.platform == 'win32':
        try:
            # 取得當前 Python 執行檔路徑
            if getattr(sys, 'frozen', False):
                # 打包後的執行檔
                script = sys.executable
                params = ' '.join([f'"{arg}"' for arg in sys.argv[1:]])
            else:
                # 開發環境
                script = sys.executable
                params = f'"{os.path.abspath(__file__)}"'
            
            # 使用 ShellExecute 以管理員權限執行
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", script, params, None, 1
            )
            return True
        except Exception as e:
            print(f"無法以管理員權限執行: {e}")
            return False
    return False

def main():
    """主程式進入點"""
    # 檢查管理員權限
    if not is_admin():
        print("正在請求管理員權限...")
        if run_as_admin():
            sys.exit(0)
        else:
            print("需要管理員權限才能執行此程式")
            input("按 Enter 鍵退出...")
            sys.exit(1)
    
    # 已有管理員權限，啟動主程式
    print("以管理員權限執行中...")
    
    # 匯入主視窗（延遲匯入以加快啟動速度）
    from ui.main_window import MainWindow
    
    try:
        app = MainWindow()
        app.run()
    except Exception as e:
        print(f"程式執行錯誤: {e}")
        import traceback
        traceback.print_exc()
        input("按 Enter 鍵退出...")
        sys.exit(1)

if __name__ == "__main__":
    main()
