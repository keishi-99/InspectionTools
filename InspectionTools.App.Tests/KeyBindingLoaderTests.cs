using System.Text.Json.Nodes;
using System.Windows.Input;
using InspectionTools.App.Input;

namespace InspectionTools.App.Tests;

/// <summary>キー割り当ての設定ファイル（keybindings.json）の読み込みを確認する。割り当てはこのファイルだけで決まり、書き間違いは、そのキーだけを無効にして知らせる</summary>
public sealed class KeyBindingLoaderTests : IDisposable {

    // 全操作のキーを書いた、正しい設定（付属の keybindings.json と同じ内容）
    private const string FullJson = """
        {
          "keys": {
            "AcquireValue": ["Space", "F2", "NumAdd"],
            "Confirm": ["Enter", "F4"],
            "Reapply": ["F3", "NumMultiply"],
            "Pass": ["F5"],
            "Fail": ["F6"],
            "Next": ["Down", "PageDown"],
            "Previous": ["Up", "PageUp"]
          },
          "globalHotkey": "]"
        }
        """;

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private static IReadOnlyList<KeyChord> ChordsOf(KeyBindingLoadResult result, KeyAction action) => result.Bindings.ChordsOf(action);

    // 全操作を書いた設定のうち、指定した操作のキーだけを差し替える（書き間違いを1つだけ試すため）
    private static string JsonWith(KeyAction action, string keysJson) {
        var root = JsonNode.Parse(FullJson)!.AsObject();
        root["keys"]![action.ToString()] = JsonNode.Parse(keysJson);
        return root.ToJsonString();
    }

    // 全操作を書いた設定のうち、全体のホットキーだけを差し替える（null なら、書かない）
    private static string JsonWithGlobalHotkey(string? hotkey) {
        var root = JsonNode.Parse(FullJson)!.AsObject();
        if (hotkey is null) {
            root.Remove("globalHotkey");
        } else {
            root["globalHotkey"] = hotkey;
        }
        return root.ToJsonString();
    }

    // ===== 付属の設定ファイル =====

    [Fact]
    public void ShippedSettingsFileEqualsTheEmergencyDefaultsWithoutWarnings() {
        var path = RepositoryFiles.Find("InspectionTools.App", KeyBindingLoader.FileName);

        var result = KeyBindingLoader.Load(path);

        Assert.Empty(result.Warnings);
        Assert.Equal(KeyBindings.Default.GlobalHotkey, result.Bindings.GlobalHotkey);
        foreach (var action in Enum.GetValues<KeyAction>()) {
            Assert.Equal(KeyBindings.Default.ChordsOf(action), ChordsOf(result, action));
        }
    }

    [Fact]
    public void FullJsonReadsEveryActionWithoutWarnings() {
        var result = KeyBindingLoader.Parse(FullJson);

        Assert.Empty(result.Warnings);
        Assert.Equal([new KeyChord(Key.Space), new KeyChord(Key.F2), new KeyChord(Key.Add)], ChordsOf(result, KeyAction.AcquireValue));
        Assert.Equal(new KeyChord(Key.OemCloseBrackets), result.Bindings.GlobalHotkey);
    }

    // ===== ファイル =====

    [Fact]
    public void MissingFileGivesTheEmergencyDefaultsAndTellsTheUser() {
        var result = KeyBindingLoader.Load(Path.Combine(_directory, KeyBindingLoader.FileName));

        Assert.Same(KeyBindings.Default, result.Bindings);
        var warning = Assert.Single(result.Warnings);
        Assert.Contains("keybindings.json", warning);
    }

    [Fact]
    public void LoadReadsAFile() {
        Directory.CreateDirectory(_directory);
        var path = Path.Combine(_directory, KeyBindingLoader.FileName);
        File.WriteAllText(path, JsonWith(KeyAction.Fail, """["F8"]"""));

        var result = KeyBindingLoader.Load(path);

        Assert.Equal([new KeyChord(Key.F8)], ChordsOf(result, KeyAction.Fail));
        Assert.Empty(result.Warnings);
    }

    [Theory]
    [InlineData("{ これはJSONではない")]
    [InlineData("null")]
    [InlineData("")]
    public void BrokenJsonGivesTheEmergencyDefaultsWithAWarning(string json) {
        var result = KeyBindingLoader.Parse(json);

        Assert.Same(KeyBindings.Default, result.Bindings);
        Assert.Single(result.Warnings);
    }

    // ===== 書かれていない操作（割り当ては、ファイルだけで決まる） =====

    [Fact]
    public void ActionsNotWrittenHaveNoKeysAndAreReported() {
        var result = KeyBindingLoader.Parse("""{ "keys": { "Confirm": ["Enter"] }, "globalHotkey": "]" }""");

        Assert.Equal([new KeyChord(Key.Return)], ChordsOf(result, KeyAction.Confirm));
        Assert.Empty(ChordsOf(result, KeyAction.Fail));
        Assert.False(result.Bindings.TryFind(new KeyChord(Key.F6), out _));
        Assert.Equal(6, result.Warnings.Count);
        Assert.Contains(result.Warnings, warning => warning.StartsWith("Fail:"));
    }

    [Fact]
    public void WrittenKeysAreNotMergedWithTheEmergencyDefaults() {
        var result = KeyBindingLoader.Parse(JsonWith(KeyAction.Confirm, """["Enter"]"""));

        Assert.Equal([new KeyChord(Key.Return)], ChordsOf(result, KeyAction.Confirm));
        Assert.False(result.Bindings.TryFind(new KeyChord(Key.F4), out _));
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void MissingKeysSectionReportsEveryAction() {
        var result = KeyBindingLoader.Parse("""{ "globalHotkey": "]" }""");

        Assert.Equal(Enum.GetValues<KeyAction>().Length, result.Warnings.Count);
    }

    [Fact]
    public void EmptyKeyListMeansNoKeysAndIsReported() {
        var result = KeyBindingLoader.Parse(JsonWith(KeyAction.Fail, "[]"));

        Assert.Empty(ChordsOf(result, KeyAction.Fail));
        Assert.Contains("Fail", Assert.Single(result.Warnings));
    }

    // ===== 書き方 =====

    [Fact]
    public void ActionNamesAreCaseInsensitive() {
        var result = KeyBindingLoader.Parse(FullJson.Replace("\"Fail\"", "\"fail\""));

        Assert.Equal([new KeyChord(Key.F6)], ChordsOf(result, KeyAction.Fail));
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void ModifiersAndSymbolKeysCanBeUsed() {
        var result = KeyBindingLoader.Parse(JsonWith(KeyAction.Fail, """["Ctrl+Shift+F6", "@"]"""));

        Assert.Equal([new KeyChord(Key.F6, ModifierKeys.Control | ModifierKeys.Shift), new KeyChord(Key.Oem3)], ChordsOf(result, KeyAction.Fail));
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void DuplicateKeysInOneActionAreCollapsedWithoutWarning() {
        var result = KeyBindingLoader.Parse(JsonWith(KeyAction.Fail, """["F8", "f8"]"""));

        Assert.Equal([new KeyChord(Key.F8)], ChordsOf(result, KeyAction.Fail));
        Assert.Empty(result.Warnings);
    }

    // ===== 書き間違い =====

    [Fact]
    public void UnknownActionIsSkippedWithAWarning() {
        var json = FullJson.Replace("\"keys\": {", "\"keys\": { \"Jump\": [\"F8\"],");

        var result = KeyBindingLoader.Parse(json);

        Assert.Contains("Jump", Assert.Single(result.Warnings));
        Assert.Equal([new KeyChord(Key.F6)], ChordsOf(result, KeyAction.Fail));
    }

    [Fact]
    public void UnreadableKeyIsDisabledButTheOtherKeysOfTheActionStay() {
        var result = KeyBindingLoader.Parse(JsonWith(KeyAction.Fail, """["F7", "Foo"]"""));

        Assert.Equal([new KeyChord(Key.F7)], ChordsOf(result, KeyAction.Fail));
        Assert.Contains("Fail", Assert.Single(result.Warnings));
        Assert.Contains("Foo", result.Warnings[0]);
    }

    [Theory]
    [InlineData("5")]
    [InlineData("NumPad5")]
    [InlineData(".")]
    [InlineData("-")]
    [InlineData("NumSubtract")]
    public void KeysUsedToTypeNumbersAreDisabled(string key) {
        var result = KeyBindingLoader.Parse(JsonWith(KeyAction.Fail, $"[\"{key}\"]"));

        // 「5」は数字だけの書き方として読めない、それ以外は数値の入力に使うキーとして断られる。どちらも無効になり、キーのない操作として知らせる
        Assert.Empty(ChordsOf(result, KeyAction.Fail));
        Assert.Equal(2, result.Warnings.Count);
    }

    [Fact]
    public void KeyUsedByAnEarlierActionIsDisabledForTheLaterOne() {
        // F2 は、先に書かれた AcquireValue のキー
        var result = KeyBindingLoader.Parse(JsonWith(KeyAction.Fail, """["F2", "F7"]"""));

        Assert.Equal([new KeyChord(Key.F7)], ChordsOf(result, KeyAction.Fail));
        Assert.Contains("AcquireValue", Assert.Single(result.Warnings));
        Assert.Equal(KeyAction.AcquireValue, FindAction(result, Key.F2));
    }

    [Fact]
    public void ActionKeyEqualToTheGlobalHotkeyIsDisabled() {
        var result = KeyBindingLoader.Parse(JsonWith(KeyAction.Fail, """["]", "F7"]"""));

        Assert.Equal([new KeyChord(Key.F7)], ChordsOf(result, KeyAction.Fail));
        Assert.Single(result.Warnings);
        Assert.Equal(new KeyChord(Key.OemCloseBrackets), result.Bindings.GlobalHotkey);
    }

    // ===== 全体のホットキー =====

    [Fact]
    public void GlobalHotkeyCanBeChanged() {
        var result = KeyBindingLoader.Parse(JsonWithGlobalHotkey("F12"));

        Assert.Equal(new KeyChord(Key.F12), result.Bindings.GlobalHotkey);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void GlobalHotkeyEqualToAnActionKeyMakesThatActionKeyDisabledNotTheHotkey() {
        // 全体のホットキーにしたキーは、システム全体で奪われるので、操作のキーとしては使えない
        var result = KeyBindingLoader.Parse(JsonWithGlobalHotkey("F6"));

        Assert.Equal(new KeyChord(Key.F6), result.Bindings.GlobalHotkey);
        Assert.Empty(ChordsOf(result, KeyAction.Fail));
        Assert.Equal(2, result.Warnings.Count);
    }

    [Fact]
    public void MissingGlobalHotkeyMeansNoHotkeyAndIsReported() {
        var result = KeyBindingLoader.Parse(JsonWithGlobalHotkey(null));

        Assert.Null(result.Bindings.GlobalHotkey);
        Assert.Contains("globalHotkey", Assert.Single(result.Warnings));
    }

    [Fact]
    public void UnreadableGlobalHotkeyMeansNoHotkeyAndIsReported() {
        var result = KeyBindingLoader.Parse(JsonWithGlobalHotkey("Foo"));

        Assert.Null(result.Bindings.GlobalHotkey);
        Assert.Contains("globalHotkey", Assert.Single(result.Warnings));
    }

    private static KeyAction? FindAction(KeyBindingLoadResult result, Key key) =>
        result.Bindings.TryFind(new KeyChord(key), out var action) ? action : null;
}
