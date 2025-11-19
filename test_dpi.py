"""測試 DPI 縮放檢測"""
import ctypes

def set_dpi_awareness():
    """設定 DPI 感知"""
    try:
        # Windows 10 / 11 - Per Monitor DPI Awareness V2
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        print("✓ 已設定 Per Monitor DPI Awareness V2")
    except:
        try:
            # Windows 8.1 / 10 - Per Monitor DPI Awareness
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
            print("✓ 已設定 Per Monitor DPI Awareness")
        except:
            try:
                # Windows Vista / 7 / 8 - System DPI Awareness
                ctypes.windll.user32.SetProcessDPIAware()
                print("✓ 已設定 System DPI Awareness")
            except:
                print("✗ 無法設定 DPI Awareness")

def get_dpi_scale():
    """取得當前 DPI 縮放比例"""
    try:
        # 獲取主螢幕的 DPI
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, hdc)
        scale = dpi / 96.0  # 96 DPI 是 100% 縮放
        return dpi, scale
    except:
        return 96, 1.0

if __name__ == "__main__":
    print("=" * 50)
    print("DPI 縮放檢測測試")
    print("=" * 50)
    
    set_dpi_awareness()
    dpi, scale = get_dpi_scale()
    
    print(f"\n目前 DPI: {dpi}")
    print(f"縮放比例: {scale:.2f}x ({int(scale * 100)}%)")
    
    # 計算建議的視窗大小（新策略 - 加寬以顯示完整 3 欄）
    base_width = 1400  # 增加基礎寬度
    base_height = 750
    
    if scale >= 1.5:  # 150% 或更高
        scaled_width = int(base_width * 0.95)
        scaled_height = int(base_height * 0.95)
    elif scale >= 1.25:  # 125% 縮放
        scaled_width = int(base_width * 0.92)
        scaled_height = int(base_height * 0.92)
    else:  # 100% 或更低
        scaled_width = base_width
        scaled_height = base_height
    
    min_width = 1300  # 增加最小寬度以確保 3 欄完整顯示
    min_height = 680
    scaled_width = max(min_width, scaled_width)
    scaled_height = max(min_height, scaled_height)
    
    print(f"\n原始視窗大小: {base_width}x{base_height}")
    print(f"建議視窗大小: {scaled_width}x{scaled_height}")
    print(f"最小視窗大小: {min_width}x{min_height}")
    
    # 計算建議的 padding（新策略：保持適當間距）
    if scale >= 1.5:
        adaptive_padding = 2
    elif scale >= 1.25:
        adaptive_padding = 2
    else:
        adaptive_padding = 2
    print(f"\n原始 padding: 2")
    print(f"建議 padding: {adaptive_padding}")
    
    # 計算建議的字型大小（新策略）
    base_font_size = 10
    if scale >= 1.5:
        scaled_font_size = 9
    elif scale >= 1.25:
        scaled_font_size = 9
    else:
        scaled_font_size = 10
    print(f"\n原始字型大小: {base_font_size}")
    print(f"建議字型大小: {scaled_font_size}")
    
    print("\n" + "=" * 50)
    
    if scale > 1.2:
        print("⚠ 偵測到高 DPI 縮放（>120%）")
        print("✓ 已自動調整介面元素以避免按鈕被壓縮")
    elif scale > 1.0:
        print("✓ 偵測到標準 DPI 縮放")
    else:
        print("✓ 偵測到標準 DPI（100%）")
    
    print("=" * 50)
