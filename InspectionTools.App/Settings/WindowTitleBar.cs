using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using MaterialDesignThemes.Wpf;

namespace InspectionTools.App.Settings;

/// <summary>ウィンドウのタイトルバー(Windows の枠)を、いまの配色に合わせる。Material Design のテーマでは変わらないため、Windows の機能で切り替える</summary>
public static class WindowTitleBar {

    // DWMWA_USE_IMMERSIVE_DARK_MODE(Windows 10 のバージョン 20H1 以降、Windows 11)
    private const int UseImmersiveDarkMode = 20;

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

    /// <summary>
    /// 現在の配色に合わせて、タイトルバーをダーク/ライトにする。ウィンドウが作られたあと(SourceInitialized 以降)に呼ぶ。
    /// 対応していない Windows では、何も起きない(タイトルバーは、ライトのまま)
    /// </summary>
    public static void ApplyCurrentTheme(Window window) {
        var handle = new WindowInteropHelper(window).Handle;
        if (handle == IntPtr.Zero) {
            return;
        }

        var isDark = new PaletteHelper().GetTheme().GetBaseTheme() == BaseTheme.Dark;
        var value = isDark ? 1 : 0;
        DwmSetWindowAttribute(handle, UseImmersiveDarkMode, ref value, sizeof(int));
    }
}
