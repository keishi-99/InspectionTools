using System.Windows.Input;
using InspectionTools.App.Input;

namespace InspectionTools.App.Tests;

/// <summary>キーと操作の対応（既定の割り当て、重なりの拒否、画面に出す案内）を確認する</summary>
public class KeyBindingsTests {

    [Theory]
    [InlineData(Key.Space, KeyAction.AcquireValue)]
    [InlineData(Key.F2, KeyAction.AcquireValue)]
    [InlineData(Key.Add, KeyAction.AcquireValue)]
    [InlineData(Key.Return, KeyAction.Confirm)]
    [InlineData(Key.F4, KeyAction.Confirm)]
    [InlineData(Key.F3, KeyAction.Reapply)]
    [InlineData(Key.Multiply, KeyAction.Reapply)]
    [InlineData(Key.F5, KeyAction.Pass)]
    [InlineData(Key.F6, KeyAction.Fail)]
    [InlineData(Key.Down, KeyAction.Next)]
    [InlineData(Key.Next, KeyAction.Next)]
    [InlineData(Key.Up, KeyAction.Previous)]
    [InlineData(Key.Prior, KeyAction.Previous)]
    public void DefaultBindingsFindTheExpectedAction(Key key, KeyAction expected) {
        Assert.True(KeyBindings.Default.TryFind(new KeyChord(key), out var action));
        Assert.Equal(expected, action);
    }

    [Fact]
    public void KeysWithOtherModifiersOrUnassignedKeysAreNotFound() {
        Assert.False(KeyBindings.Default.TryFind(new KeyChord(Key.Return, ModifierKeys.Control), out _));
        Assert.False(KeyBindings.Default.TryFind(new KeyChord(Key.F12), out _));
        Assert.False(KeyBindings.Default.TryFind(new KeyChord(Key.D5), out _));
    }

    [Fact]
    public void EveryActionHasAtLeastOneDefaultKeyThatWorksWithoutTheNumericKeypad() {
        Key[] keypadOnly = [Key.Add, Key.Multiply, Key.Divide, Key.Subtract];

        foreach (var action in Enum.GetValues<KeyAction>()) {
            Assert.Contains(KeyBindings.Default.ChordsOf(action), chord => !keypadOnly.Contains(chord.Key));
        }
    }

    [Fact]
    public void NoDefaultKeyIsUsedToTypeNumbers() {
        var all = Enum.GetValues<KeyAction>().SelectMany(KeyBindings.Default.ChordsOf);

        Assert.All(all, chord => Assert.False(chord.IsNumericEntryKey, chord.ToString()));
    }

    [Fact]
    public void DefaultGlobalHotkeyIsTheClosingBracketKeyAndDiffersFromEveryActionKey() {
        Assert.Equal(new KeyChord(Key.OemCloseBrackets), KeyBindings.Default.GlobalHotkey);
        Assert.False(KeyBindings.Default.TryFind(KeyBindings.Default.GlobalHotkey!.Value, out _));
    }

    [Fact]
    public void SameKeyForTwoActionsIsRejected() {
        var chords = new Dictionary<KeyAction, IReadOnlyList<KeyChord>> {
            [KeyAction.Confirm] = [new(Key.F2)],
            [KeyAction.AcquireValue] = [new(Key.F2)],
        };

        var exception = Assert.Throws<ArgumentException>(() => new KeyBindings(chords, new KeyChord(Key.OemCloseBrackets)));

        Assert.Contains("F2", exception.Message);
    }

    [Fact]
    public void GlobalHotkeyEqualToAnActionKeyIsRejected() {
        var chords = new Dictionary<KeyAction, IReadOnlyList<KeyChord>> { [KeyAction.Confirm] = [new(Key.F2)] };

        Assert.Throws<ArgumentException>(() => new KeyBindings(chords, new KeyChord(Key.F2)));
    }

    [Fact]
    public void GlobalHotkeyCanBeAbsent() {
        var bindings = new KeyBindings(new Dictionary<KeyAction, IReadOnlyList<KeyChord>>(), globalHotkey: null);

        Assert.Null(bindings.GlobalHotkey);
    }

    [Fact]
    public void ChordsOfReturnsEmptyForAnActionWithoutKeys() {
        var bindings = new KeyBindings(new Dictionary<KeyAction, IReadOnlyList<KeyChord>>(), new KeyChord(Key.F12));

        Assert.Empty(bindings.ChordsOf(KeyAction.Confirm));
        Assert.False(bindings.TryFind(new KeyChord(Key.Return), out _));
    }

    [Fact]
    public void HintTextListsTheFirstKeyOfEachAction() {
        var hint = KeyBindings.Default.HintText;

        Assert.Contains("値を取得 [Space]", hint);
        Assert.Contains("確定 [Enter]", hint);
        Assert.Contains("再適用 [F3]", hint);
        Assert.Contains("合格 [F5]", hint);
        Assert.Contains("不合格 [F6]", hint);
        Assert.Contains("次へ [Down]", hint);
        Assert.Contains("前へ [Up]", hint);
    }

    [Fact]
    public void HintTextSkipsActionsWithoutKeys() {
        var chords = new Dictionary<KeyAction, IReadOnlyList<KeyChord>> { [KeyAction.Confirm] = [new(Key.Return)] };

        var hint = new KeyBindings(chords, new KeyChord(Key.F12)).HintText;

        Assert.Equal("確定 [Enter]", hint);
    }
}
