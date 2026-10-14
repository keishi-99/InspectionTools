using System.Diagnostics.CodeAnalysis;
using System.Windows.Input;

namespace InspectionTools.App.Input;

/// <summary>キー1つと修飾キー（Ctrl、Shift、Alt、Windows）の組み合わせ。WPF の KeyGesture は Space や Enter を修飾キーなしでは使えないため、自前で持つ</summary>
public readonly record struct KeyChord(Key Key, ModifierKeys Modifiers = ModifierKeys.None) {

    // 設定ファイルに書く名前（別名）と実際のキー。ToString はここの最初の名前を使う（Parse と往復できるようにするため）。
    // 記号キーは日本語配列のキー位置に合わせる（WPF のキー名は米国配列の名前なので、そのままでは分かりにくい）
    private static readonly (string Name, Key Key)[] Aliases = [
        ("Enter", Key.Return),
        ("Space", Key.Space),
        ("Esc", Key.Escape),
        ("PageUp", Key.Prior),
        ("PageDown", Key.Next),
        ("NumAdd", Key.Add),
        ("NumSubtract", Key.Subtract),
        ("NumMultiply", Key.Multiply),
        ("NumDivide", Key.Divide),
        ("]", Key.OemCloseBrackets),
        ("[", Key.OemOpenBrackets),
        ("\\", Key.OemPipe),
        (":", Key.Oem1),
        (";", Key.OemPlus),
        ("@", Key.Oem3),
        ("^", Key.Oem7),
        (",", Key.OemComma),
        (".", Key.OemPeriod),
        ("/", Key.OemQuestion),
        ("-", Key.OemMinus),
    ];

    /// <summary>設定ファイルの書き方（例: F5、Ctrl+Shift+F5、Enter、]）から作る。読めないときは理由を error に入れて false を返す</summary>
    public static bool TryParse(string? text, out KeyChord chord, [NotNullWhen(false)] out string? error) {
        chord = default;
        if (string.IsNullOrWhiteSpace(text)) {
            error = "キーの指定が空です。";
            return false;
        }

        var modifiers = ModifierKeys.None;
        var tokens = text.Trim().Split('+');
        // 最後の要素がキー、その前がすべて修飾キー。「Ctrl++」のように + 自体を指定することはできない（NumAdd と書く）
        for (var index = 0; index < tokens.Length - 1; index++) {
            var modifier = ParseModifier(tokens[index].Trim());
            if (modifier is null) {
                error = $"修飾キーとして読めません: '{tokens[index].Trim()}'（Ctrl、Shift、Alt、Win が使えます。キー自体に + を使うときは NumAdd と書きます）";
                return false;
            }
            modifiers |= modifier.Value;
        }

        var keyName = tokens[^1].Trim();
        var key = ParseKey(keyName);
        if (key is null) {
            error = $"キーとして読めません: '{keyName}'";
            return false;
        }

        chord = new KeyChord(key.Value, modifiers);
        error = null;
        return true;
    }

    /// <summary>設定ファイルの書き方から作る。読めないときは FormatException</summary>
    public static KeyChord Parse(string? text) {
        return TryParse(text, out var chord, out var error) ? chord : throw new FormatException(error);
    }

    /// <summary>数値の入力欄に打つキー（数字、小数点、マイナス）か。これらを操作に割り当てると、値を入力できなくなる（Ctrl や Alt を伴うときは別）</summary>
    public bool IsNumericEntryKey {
        get {
            if ((Modifiers & (ModifierKeys.Control | ModifierKeys.Alt | ModifierKeys.Windows)) != ModifierKeys.None) {
                return false;
            }
            return Key is >= Key.D0 and <= Key.D9
                or >= Key.NumPad0 and <= Key.NumPad9
                or Key.Decimal or Key.OemPeriod or Key.Subtract or Key.OemMinus;
        }
    }

    /// <summary>設定ファイルに書く名前に直す（例: Ctrl+F5）</summary>
    public override string ToString() {
        var parts = new List<string>();
        if (Modifiers.HasFlag(ModifierKeys.Control)) {
            parts.Add("Ctrl");
        }
        if (Modifiers.HasFlag(ModifierKeys.Shift)) {
            parts.Add("Shift");
        }
        if (Modifiers.HasFlag(ModifierKeys.Alt)) {
            parts.Add("Alt");
        }
        if (Modifiers.HasFlag(ModifierKeys.Windows)) {
            parts.Add("Win");
        }

        // 構造体の中のラムダ式は this を使えないので、キーを変数に取り出す
        var key = Key;
        var alias = Aliases.FirstOrDefault(entry => entry.Key == key);
        parts.Add(alias.Name ?? key.ToString());
        return string.Join("+", parts);
    }

    private static ModifierKeys? ParseModifier(string name) {
        return name.ToLowerInvariant() switch {
            "ctrl" or "control" => ModifierKeys.Control,
            "shift" => ModifierKeys.Shift,
            "alt" => ModifierKeys.Alt,
            "win" or "windows" => ModifierKeys.Windows,
            _ => null,
        };
    }

    private static Key? ParseKey(string name) {
        if (name.Length == 0) {
            return null;
        }

        var alias = Aliases.FirstOrDefault(entry => string.Equals(entry.Name, name, StringComparison.OrdinalIgnoreCase));
        if (alias.Name is not null) {
            return alias.Key;
        }

        // 別名にないものは、WPF のキー名（F2、Down、Home、A など）をそのまま使う。数字だけの文字列は、キーの番号として読まれてしまうので除く
        if (int.TryParse(name, out _)) {
            return null;
        }
        return Enum.TryParse<Key>(name, ignoreCase: true, out var key) && key != Key.None && !IsNotUsableAlone(key) ? key : null;
    }

    // 操作のキーにできないキー（修飾キーそのもの、日本語入力などが内部で使うキー）
    private static bool IsNotUsableAlone(Key key) => key is Key.LeftCtrl or Key.RightCtrl or Key.LeftShift or Key.RightShift
        or Key.LeftAlt or Key.RightAlt or Key.LWin or Key.RWin or Key.System or Key.ImeProcessed or Key.DeadCharProcessed;
}
