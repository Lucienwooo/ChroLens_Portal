"""測試 .lnk 捷徑開啟功能"""
import os
import subprocess
import pythoncom
from win32com.shell import shell

def open_lnk_target(lnk_path):
    """解析 .lnk 捷徑檔案，回傳 (目標路徑, 參數字串)
    使用多種方法嘗試解析，確保兼容性"""
    
    print("\n嘗試方法 1: win32com.client.Dispatch (WScript.Shell)")
    # 方法 1: 使用 win32com.client (更穩定的方式)
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)
        target_path = shortcut.Targetpath
        arguments = shortcut.Arguments
        if target_path:
            print(f"✓ 成功: {target_path}")
            return target_path, arguments
    except Exception as e:
        print(f"✗ 失敗: {e}")
    
    print("\n嘗試方法 2: pythoncom (傳統方法)")
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
                print(f"✓ 成功: {target_path}")
                return target_path, arguments
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        print(f"✗ 失敗: {e}")
    
    print("\n嘗試方法 3: PowerShell (最穩定)")
    # 方法 3: 使用 PowerShell (最可靠的備援方法)
    try:
        import subprocess
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
                print(f"✓ 成功: {target_path}")
                return target_path, arguments
    except Exception as e:
        print(f"✗ 失敗: {e}")
    
    # 所有方法都失敗
    print("\n✗ 所有解析方法都失敗")
    return None, None

def test_open_file(file_path):
    """測試開啟檔案"""
    print(f"\n{'='*60}")
    print(f"測試檔案: {file_path}")
    print(f"{'='*60}")
    
    if not os.path.exists(file_path):
        print(f"❌ 檔案不存在")
        return False
    
    try:
        if file_path.lower().endswith('.lnk'):
            # 解析捷徑
            print("檔案類型: .lnk 捷徑")
            target, args = open_lnk_target(file_path)
            if target and os.path.exists(target):
                print(f"✓ 捷徑目標: {target}")
                print(f"✓ 參數: {args if args else '(無)'}")
                print(f"✓ 目標存在: 是")
                
                # 執行目標
                if args:
                    cmd = f'"{target}" {args}'
                else:
                    cmd = f'"{target}"'
                print(f"執行命令: {cmd}")
                subprocess.Popen(cmd, shell=True)
                print("✓ 已啟動程式")
                return True
            else:
                print(f"⚠ 捷徑解析失敗，嘗試直接開啟")
                os.startfile(file_path)
                print("✓ 已使用 Windows 預設方式開啟捷徑")
                return True
                
        elif file_path.lower().endswith('.exe'):
            # 直接執行 .exe 檔案
            print("檔案類型: .exe 執行檔")
            subprocess.Popen(f'"{file_path}"', shell=True)
            print("✓ 已啟動程式")
            return True
        else:
            # 其他檔案類型
            print(f"檔案類型: 其他 ({os.path.splitext(file_path)[1]})")
            os.startfile(file_path)
            print("✓ 已使用系統預設程式開啟")
            return True
            
    except Exception as e:
        print(f"❌ 開啟失敗: {e}")
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("捷徑開啟功能測試")
    print("=" * 60)
    
    # 測試檔案路徑
    test_file = r"C:/Users/Lucien/Desktop/0-shot\噗噗.lnk"
    
    print(f"\n原始路徑: {test_file}")
    
    # 正規化路徑（將 / 轉換為 \）
    normalized_path = os.path.normpath(test_file)
    print(f"正規化路徑: {normalized_path}")
    
    # 測試開啟
    success = test_open_file(normalized_path)
    
    print("\n" + "=" * 60)
    if success:
        print("✅ 測試成功")
    else:
        print("❌ 測試失敗")
    print("=" * 60)
    
    # 額外提示
    print("\n💡 提示:")
    print("1. 確保 pywin32 已正確安裝並配置")
    print("2. 捷徑檔案路徑中的反斜線會被自動處理")
    print("3. 如果解析失敗，會嘗試使用 Windows 預設方式開啟")
    print("4. .exe 檔案會直接執行")
    print("5. 其他檔案會用系統預設程式開啟")
