using MaterialDesignThemes.Wpf;

namespace InspectionTools.App.Settings;

/// <summary>Material Design のテーマを切り替える(画面の部品を持つので、UI スレッドから呼ぶ)</summary>
public sealed class MaterialThemeApplier : IThemeApplier {

    public void Apply(ThemeSetting theme) {
        var paletteHelper = new PaletteHelper();
        var materialTheme = paletteHelper.GetTheme();
        materialTheme.SetBaseTheme(theme == ThemeSetting.Dark ? BaseTheme.Dark : BaseTheme.Light);
        paletteHelper.SetTheme(materialTheme);
        // このアプリ独自の色(表の合否の色など)も、配色に合わせる
        ThemeBrushes.Apply(theme);
    }
}
