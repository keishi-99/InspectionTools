using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using InspectionTools.App.ViewModels;

namespace InspectionTools.App.Settings;

/// <summary>画面の配色を実際に切り替える(テストで差し替えられるように、窓口だけを決めておく)</summary>
public interface IThemeApplier {

    /// <summary>配色を切り替える</summary>
    void Apply(ThemeSetting theme);
}

/// <summary>
/// 画面の配色(ライト/ダーク)を、設定ファイルと結びつけて扱う。
/// 起動時に設定を読んで適用し、切り替えたら適用して、設定ファイルに保存する
/// </summary>
public sealed partial class ThemeController(string settingsPath, IThemeApplier applier, IDialogService dialog) : ObservableObject {

    /// <summary>いまの配色</summary>
    public ThemeSetting Theme { get; private set; } = ThemeSetting.Light;

    /// <summary>いま、ダークか(スイッチの表示に使う)</summary>
    public bool IsDark => Theme == ThemeSetting.Dark;

    /// <summary>起動時に呼ぶ。設定ファイルを読んで、配色を適用する(保存はしない)。読めなかったときは、ライトにして、理由を通知する</summary>
    public void Initialize() {
        var result = AppSettingsStore.Load(settingsPath);
        Theme = result.Settings.Theme;
        applier.Apply(Theme);
        if (result.Warning is not null) {
            dialog.ShowError($"設定ファイルを読み込めなかったため、ライトで起動しました。\n{result.Warning}");
        }
    }

    /// <summary>配色を切り替える。適用してから、設定ファイルに保存する(保存に失敗しても、切り替えは有効のまま。次の起動で引き継げないことを通知する)</summary>
    public void SetTheme(ThemeSetting theme) {
        if (theme == Theme) {
            return;
        }

        Theme = theme;
        applier.Apply(theme);
        OnPropertyChanged(nameof(Theme));
        OnPropertyChanged(nameof(IsDark));

        try {
            AppSettingsStore.Save(settingsPath, new AppSettings(theme));
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            dialog.ShowError($"配色の設定を保存できませんでした。次の起動では、元の配色に戻ります。\n{ex.Message}");
        }
    }
}
