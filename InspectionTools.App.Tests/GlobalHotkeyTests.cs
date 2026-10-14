using System.Windows.Input;
using InspectionTools.App.Input;

namespace InspectionTools.App.Tests;

/// <summary>全体のホットキーのうち、画面なしで確かめられる部分（キーを Windows の値に直す）を確認する</summary>
public sealed class GlobalHotkeyTests {

    [Fact]
    public void CloseBracketConvertsToItsVirtualKeyWithoutModifiers() {
        Assert.True(GlobalHotkey.TryConvert(new KeyChord(Key.OemCloseBrackets), out var modifiers, out var virtualKey));

        Assert.Equal(0u, modifiers);
        Assert.Equal(0xDDu, virtualKey);
    }

    [Fact]
    public void ModifiersAreConvertedToWindowsFlags() {
        Assert.True(GlobalHotkey.TryConvert(new KeyChord(Key.F5, ModifierKeys.Control | ModifierKeys.Shift), out var modifiers, out var virtualKey));

        // Ctrl = 2、Shift = 4
        Assert.Equal(6u, modifiers);
        Assert.Equal(0x74u, virtualKey);
    }

    [Fact]
    public void AltAndWindowsModifiersAreConverted() {
        Assert.True(GlobalHotkey.TryConvert(new KeyChord(Key.F1, ModifierKeys.Alt | ModifierKeys.Windows), out var modifiers, out _));

        Assert.Equal(9u, modifiers);
    }

    [Fact]
    public void KeyWithoutVirtualKeyCannotBeConverted() {
        Assert.False(GlobalHotkey.TryConvert(new KeyChord(Key.None), out _, out _));
    }
}
