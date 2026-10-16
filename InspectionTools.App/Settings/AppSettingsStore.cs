using System.IO;
using System.Text.Json;

namespace InspectionTools.App.Settings;

/// <summary>設定ファイル(settings.json)の読み込みと保存</summary>
public static class AppSettingsStore {

    /// <summary>設定ファイルの名前(実行ファイルのフォルダに置く)</summary>
    public const string FileName = "settings.json";

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web) {
        WriteIndented = true,
    };

    // JSON の形(読み間違いを見つけられるように、配色は文字列で受ける)
    private sealed record SettingsFile(string? Theme);

    /// <summary>設定を読み込む。ファイルがなければ、既定の設定(警告なし)。読めないときは、既定の設定と理由を返す(アプリは止めない)</summary>
    public static AppSettingsLoadResult Load(string path) {
        if (!File.Exists(path)) {
            return new AppSettingsLoadResult(AppSettings.Default, null);
        }

        string json;
        try {
            json = File.ReadAllText(path);
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            return Failed($"{FileName} を読み込めません: {ex.Message}");
        }
        return Parse(json);
    }

    /// <summary>JSON の文字列から設定を作る</summary>
    public static AppSettingsLoadResult Parse(string json) {
        SettingsFile? file;
        try {
            file = JsonSerializer.Deserialize<SettingsFile>(json, JsonOptions);
        } catch (JsonException ex) {
            return Failed($"{FileName} の JSON が正しくありません: {ex.Message}");
        }
        if (file is null) {
            return Failed($"{FileName} の内容が空です。");
        }

        // 配色が書かれていないときは、既定の配色にする(警告は出さない。ほかの設定を足したときに、古いファイルでも読めるように)
        if (file.Theme is null) {
            return new AppSettingsLoadResult(AppSettings.Default, null);
        }
        if (!Enum.TryParse<ThemeSetting>(file.Theme, ignoreCase: true, out var theme) || !Enum.IsDefined(theme)) {
            return Failed($"{FileName} の theme は、light か dark で指定してください(現在: '{file.Theme}')。");
        }
        return new AppSettingsLoadResult(new AppSettings(theme), null);
    }

    /// <summary>設定を保存する。書き込めないときは IOException か UnauthorizedAccessException</summary>
    public static void Save(string path, AppSettings settings) {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        // 書き込み中にアプリが落ちても、古い完全なファイルが残るように、別名に最後まで書いてから置き換える
        var temporaryPath = path + ".tmp";
        var file = new SettingsFile(settings.Theme.ToString().ToLowerInvariant());
        File.WriteAllText(temporaryPath, JsonSerializer.Serialize(file, JsonOptions));
        File.Move(temporaryPath, path, overwrite: true);
    }

    private static AppSettingsLoadResult Failed(string warning) => new(AppSettings.Default, warning);
}
