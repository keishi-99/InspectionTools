namespace InspectionTools.App.Input;

/// <summary>画面に出すキー1つ(枠で囲んで表示する)。記号1文字のキーは、枠の中でつぶれないように大きく太く出す</summary>
public sealed record KeyBadge(string Text) {

    /// <summary>記号1文字か(例: ]、↓)</summary>
    public bool IsSymbol => Text.Length == 1 && !char.IsLetterOrDigit(Text[0]);
}

/// <summary>画面下部のキーの案内の1項目(操作の名前と、その操作に割り当てたキーすべて)</summary>
/// <param name="IsAvailable">いまの項目で使える操作か(使えないときは、薄く出す)</param>
public sealed record KeyHintItem(string Label, IReadOnlyList<KeyBadge> Keys, KeyAction Action, bool IsAvailable = true);
