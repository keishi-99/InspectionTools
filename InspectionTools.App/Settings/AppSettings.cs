namespace InspectionTools.App.Settings;

/// <summary>画面の配色(ライト／ダーク)</summary>
public enum ThemeSetting {
    Light,
    Dark,
}

/// <summary>アプリの設定(実行ファイルのフォルダの settings.json に保存し、次の起動で引き継ぐ)</summary>
/// <param name="Theme">画面の配色</param>
public sealed record AppSettings(ThemeSetting Theme) {

    /// <summary>設定ファイルがない、または読めないときの設定</summary>
    public static AppSettings Default { get; } = new(ThemeSetting.Light);
}

/// <summary>設定の読み込み結果。読めなかったときは、既定の設定と、その理由(Warning)を返す</summary>
/// <param name="Warning">読めなかった理由(問題がなければ null)</param>
public sealed record AppSettingsLoadResult(AppSettings Settings, string? Warning);
