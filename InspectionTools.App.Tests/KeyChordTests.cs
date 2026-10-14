using System.Windows.Input;
using InspectionTools.App.Input;

namespace InspectionTools.App.Tests;

/// <summary>キーの書き方（設定ファイルの文字）と、キーの組み合わせの扱いを確認する</summary>
public class KeyChordTests {

    [Theory]
    [InlineData("F2", Key.F2, ModifierKeys.None)]
    [InlineData("Enter", Key.Return, ModifierKeys.None)]
    [InlineData("Space", Key.Space, ModifierKeys.None)]
    [InlineData("Down", Key.Down, ModifierKeys.None)]
    [InlineData("PageDown", Key.Next, ModifierKeys.None)]
    [InlineData("PageUp", Key.Prior, ModifierKeys.None)]
    [InlineData("NumAdd", Key.Add, ModifierKeys.None)]
    [InlineData("NumMultiply", Key.Multiply, ModifierKeys.None)]
    [InlineData("]", Key.OemCloseBrackets, ModifierKeys.None)]
    [InlineData("[", Key.OemOpenBrackets, ModifierKeys.None)]
    [InlineData("@", Key.Oem3, ModifierKeys.None)]
    [InlineData("A", Key.A, ModifierKeys.None)]
    [InlineData("Ctrl+F5", Key.F5, ModifierKeys.Control)]
    [InlineData("Ctrl+Shift+F5", Key.F5, ModifierKeys.Control | ModifierKeys.Shift)]
    [InlineData("Alt+Enter", Key.Return, ModifierKeys.Alt)]
    [InlineData("Win+]", Key.OemCloseBrackets, ModifierKeys.Windows)]
    public void ParseReadsKeyAndModifiers(string text, Key key, ModifierKeys modifiers) {
        Assert.Equal(new KeyChord(key, modifiers), KeyChord.Parse(text));
    }

    [Theory]
    [InlineData("enter")]
    [InlineData("SPACE")]
    [InlineData("ctrl+f5")]
    [InlineData(" Ctrl + F5 ")]
    [InlineData("CONTROL+F5")]
    public void ParseIgnoresCaseAndSurroundingSpaces(string text) {
        Assert.True(KeyChord.TryParse(text, out _, out _));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    [InlineData("Foo")]
    [InlineData("Ctrl")]
    [InlineData("Ctrl+")]
    [InlineData("Ctrl+Foo")]
    [InlineData("Hyper+F2")]
    [InlineData("12")]
    [InlineData("+")]
    [InlineData("LeftCtrl")]
    [InlineData("Shift+LeftShift")]
    [InlineData("None")]
    public void ParseRejectsUnreadableKeysWithAReason(string? text) {
        var isParsed = KeyChord.TryParse(text, out _, out var error);

        Assert.False(isParsed);
        Assert.False(string.IsNullOrEmpty(error));
        Assert.Throws<FormatException>(() => KeyChord.Parse(text));
    }

    [Fact]
    public void ToStringCanBeParsedBack() {
        KeyChord[] chords = [
            new(Key.Return), new(Key.Space), new(Key.F2), new(Key.Next), new(Key.Prior), new(Key.Add), new(Key.Multiply),
            new(Key.OemCloseBrackets), new(Key.Oem3), new(Key.A), new(Key.F5, ModifierKeys.Control | ModifierKeys.Shift | ModifierKeys.Alt),
        ];

        Assert.All(chords, chord => Assert.Equal(chord, KeyChord.Parse(chord.ToString())));
    }

    [Fact]
    public void ToStringUsesTheFriendlyNames() {
        Assert.Equal("Enter", new KeyChord(Key.Return).ToString());
        Assert.Equal("Ctrl+Shift+F5", new KeyChord(Key.F5, ModifierKeys.Control | ModifierKeys.Shift).ToString());
        Assert.Equal("]", new KeyChord(Key.OemCloseBrackets).ToString());
    }

    [Theory]
    [InlineData(Key.D0, ModifierKeys.None, true)]
    [InlineData(Key.D9, ModifierKeys.None, true)]
    [InlineData(Key.NumPad5, ModifierKeys.None, true)]
    [InlineData(Key.Decimal, ModifierKeys.None, true)]
    [InlineData(Key.OemPeriod, ModifierKeys.None, true)]
    [InlineData(Key.OemMinus, ModifierKeys.None, true)]
    [InlineData(Key.Subtract, ModifierKeys.None, true)]
    [InlineData(Key.D5, ModifierKeys.Shift, true)]
    [InlineData(Key.D5, ModifierKeys.Control, false)]
    [InlineData(Key.OemMinus, ModifierKeys.Alt, false)]
    [InlineData(Key.F2, ModifierKeys.None, false)]
    [InlineData(Key.Space, ModifierKeys.None, false)]
    [InlineData(Key.Return, ModifierKeys.None, false)]
    [InlineData(Key.Add, ModifierKeys.None, false)]
    [InlineData(Key.OemCloseBrackets, ModifierKeys.None, false)]
    public void IsNumericEntryKeyIdentifiesKeysUsedToTypeNumbers(Key key, ModifierKeys modifiers, bool expected) {
        Assert.Equal(expected, new KeyChord(key, modifiers).IsNumericEntryKey);
    }
}
