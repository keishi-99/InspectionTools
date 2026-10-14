using System.IO;
using System.Text.Json;

namespace InspectionTools.App.Input;

/// <summary>キー割り当ての読み込み結果</summary>
/// <param name="Bindings">使うキー割り当て</param>
/// <param name="Warnings">読み込めなかった部分や、キーのない操作の案内（なければ空）</param>
public sealed record KeyBindingLoadResult(KeyBindings Bindings, IReadOnlyList<string> Warnings);

/// <summary>
/// キー割り当ての設定ファイル（実行ファイルのフォルダの keybindings.json）を読み込む。割り当ては、このファイルだけで決まる
/// （書かれていない操作は、キーなし）。書き間違いがあってもアプリは止めず、そのキーだけを無効にして知らせる。
/// ファイルがない、または JSON として読めないときだけ、緊急用の割り当て（KeyBindings.Default）を使う。
/// <code>{ "keys": { "AcquireValue": ["Space", "F2"], "Confirm": ["Enter"] }, "globalHotkey": "]" }</code>
/// </summary>
public static class KeyBindingLoader {

    /// <summary>設定ファイルの名前（実行ファイルのフォルダに置く）</summary>
    public const string FileName = "keybindings.json";

    private sealed record SettingsFile(Dictionary<string, List<string>>? Keys, string? GlobalHotkey);

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    /// <summary>ファイルから読み込む。ファイルがないときは、緊急用の割り当てを使い、その旨を知らせる</summary>
    public static KeyBindingLoadResult Load(string path) {
        if (!File.Exists(path)) {
            return Fallback($"キー割り当ての設定ファイル（{FileName}）が見つかりません。実行ファイルのフォルダに置いてください。");
        }

        try {
            return Parse(File.ReadAllText(path));
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            return Fallback($"キー割り当ての設定ファイル（{FileName}）を読み込めません: {ex.Message}");
        }
    }

    /// <summary>JSON の文字列から読み込む</summary>
    public static KeyBindingLoadResult Parse(string json) {
        SettingsFile? file;
        try {
            file = JsonSerializer.Deserialize<SettingsFile>(json, JsonOptions);
        } catch (JsonException ex) {
            return Fallback($"キー割り当ての設定が JSON として正しくありません: {ex.Message}");
        }
        if (file is null) {
            return Fallback("キー割り当ての設定が空です。");
        }

        var warnings = new List<string>();

        // 全体のホットキーを先に読む（操作のキーが、これと重ならないように確認するため）
        var globalHotkey = ReadGlobalHotkey(file.GlobalHotkey, warnings);

        var chords = new Dictionary<KeyAction, IReadOnlyList<KeyChord>>();
        var usedBy = new Dictionary<KeyChord, KeyAction>();
        foreach (var (name, texts) in file.Keys ?? []) {
            if (!Enum.TryParse<KeyAction>(name, ignoreCase: true, out var action) || !Enum.IsDefined(action)) {
                warnings.Add($"未対応の操作です（読み飛ばしました）: {name}（使える操作: {string.Join("、", Enum.GetValues<KeyAction>())}）");
                continue;
            }
            if (chords.ContainsKey(action)) {
                warnings.Add($"{action} が複数回書かれています（2回目以降は読み飛ばしました）。");
                continue;
            }
            chords[action] = ReadChords(action, texts ?? [], globalHotkey, usedBy, warnings);
        }

        // 書かれていない操作、またはキーがすべて無効だった操作は、キーなし（マウスでは操作できる）
        foreach (var action in Enum.GetValues<KeyAction>()) {
            if (!chords.TryGetValue(action, out var list) || list.Count == 0) {
                chords[action] = [];
                warnings.Add($"{action}: 割り当てるキーがありません（この操作は、キーでは行えません）。");
            }
        }

        return new KeyBindingLoadResult(new KeyBindings(chords, globalHotkey), warnings);
    }

    // 全体のホットキーの書き方を読む。書かれていない、または読めないときは null（ホットキーなし）
    private static KeyChord? ReadGlobalHotkey(string? text, List<string> warnings) {
        if (text is null) {
            warnings.Add("globalHotkey が書かれていないため、新アプリを前面に出すホットキーは使えません。");
            return null;
        }
        if (!KeyChord.TryParse(text, out var chord, out var error)) {
            warnings.Add($"globalHotkey: {error}（新アプリを前面に出すホットキーは使えません）");
            return null;
        }
        return chord;
    }

    // 1つの操作のキーの書き方を読む。使えないキーだけを無効にして知らせ、使えるキーは生かす
    private static List<KeyChord> ReadChords(
        KeyAction action,
        List<string> texts,
        KeyChord? globalHotkey,
        Dictionary<KeyChord, KeyAction> usedBy,
        List<string> warnings) {

        var accepted = new List<KeyChord>();
        foreach (var text in texts) {
            if (!KeyChord.TryParse(text, out var chord, out var error)) {
                warnings.Add($"{action}: {error}（このキーは無効にしました）");
                continue;
            }
            if (chord.IsNumericEntryKey) {
                warnings.Add($"{action}: {chord} は数値の入力に使うキーなので、割り当てられません（このキーは無効にしました）。");
                continue;
            }
            if (globalHotkey is { } hotkey && chord == hotkey) {
                warnings.Add($"{action}: {chord} は、新アプリを前面に出す全体のホットキーと同じなので、割り当てられません（このキーは無効にしました）。");
                continue;
            }
            if (usedBy.TryGetValue(chord, out var other)) {
                // 同じ操作の中の重複は黙って1つにまとめる。別の操作と重なるときは、先に書かれた操作を優先する
                if (other != action) {
                    warnings.Add($"{action}: {chord} は、すでに {other} に割り当てられているので、無効にしました。");
                }
                continue;
            }
            usedBy[chord] = action;
            accepted.Add(chord);
        }
        return accepted;
    }

    private static KeyBindingLoadResult Fallback(string warning) => new(KeyBindings.Default, [$"{warning}（緊急用の割り当てを使います）"]);
}
