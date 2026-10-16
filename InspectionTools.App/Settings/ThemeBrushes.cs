using System.Windows;
using System.Windows.Media;

namespace InspectionTools.App.Settings;

/// <summary>
/// 配色(ライト/ダーク)ごとに変える、決め打ちの色。Material Design のテーマには含まれない、このアプリ独自の色を、
/// アプリのリソースに入れる。画面は、DynamicResource で参照するので、配色を切り替えると、すぐに変わる
/// </summary>
public static class ThemeBrushes {

    /// <summary>表のマス: 合格</summary>
    public const string CellOk = "CellOkBrush";

    /// <summary>表のマス: 不合格</summary>
    public const string CellNg = "CellNgBrush";

    /// <summary>表のマス: 完了(作業項目)</summary>
    public const string CellDone = "CellDoneBrush";

    /// <summary>通信中に、画面を覆う膜</summary>
    public const string BusyOverlay = "BusyOverlayBrush";

    /// <summary>配色に合わせた色を、アプリのリソースに入れる(画面の部品を持つので、UI スレッドから呼ぶ)</summary>
    public static void Apply(ThemeSetting theme) {
        var resources = Application.Current.Resources;
        var isDark = theme == ThemeSetting.Dark;
        resources[CellOk] = Create(isDark ? Color.FromRgb(0x2E, 0x5E, 0x34) : Color.FromRgb(0xC8, 0xE6, 0xC9));
        resources[CellNg] = Create(isDark ? Color.FromRgb(0x6D, 0x2F, 0x36) : Color.FromRgb(0xFF, 0xCD, 0xD2));
        resources[CellDone] = Create(isDark ? Color.FromRgb(0x2B, 0x4C, 0x75) : Color.FromRgb(0xBB, 0xDE, 0xFB));
        resources[BusyOverlay] = Create(isDark ? Color.FromArgb(0x88, 0x00, 0x00, 0x00) : Color.FromArgb(0x66, 0xFF, 0xFF, 0xFF));
    }

    private static Brush Create(Color color) {
        var brush = new SolidColorBrush(color);
        brush.Freeze();
        return brush;
    }
}
