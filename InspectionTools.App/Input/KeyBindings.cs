using System.Windows.Input;

namespace InspectionTools.App.Input;

/// <summary>キーと操作の対応（1つの操作に複数のキーを割り当てられるが、同じキーを複数の操作に割り当てることはできない）と、全体で効くホットキー</summary>
public sealed class KeyBindings {

    // 画面に出す操作の名前（ヒントの文字）
    private static readonly Dictionary<KeyAction, string> Labels = new() {
        [KeyAction.AcquireValue] = "値を取得",
        [KeyAction.Confirm] = "確定",
        [KeyAction.Pass] = "合格",
        [KeyAction.Fail] = "不合格",
        [KeyAction.Reapply] = "再適用",
        [KeyAction.Next] = "次へ",
        [KeyAction.Previous] = "前へ",
    };

    private readonly Dictionary<KeyAction, IReadOnlyList<KeyChord>> _chords;

    /// <param name="chords">操作ごとのキー</param>
    /// <param name="globalHotkey">新アプリを前面に出してフォーカスする、全体で効くホットキー（使わないときは null）</param>
    /// <exception cref="ArgumentException">同じキーが複数の操作に割り当てられているとき、ホットキーが操作のキーと同じとき</exception>
    public KeyBindings(IReadOnlyDictionary<KeyAction, IReadOnlyList<KeyChord>> chords, KeyChord? globalHotkey) {
        var all = chords.SelectMany(pair => pair.Value.Select(chord => (pair.Key, chord))).ToList();
        var duplicate = all.GroupBy(entry => entry.chord).FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null) {
            var actions = string.Join("、", duplicate.Select(entry => entry.Key));
            throw new ArgumentException($"同じキー（{duplicate.Key}）が複数の操作（{actions}）に割り当てられています。", nameof(chords));
        }
        if (all.Any(entry => entry.chord == globalHotkey)) {
            throw new ArgumentException($"全体のホットキー（{globalHotkey}）が、操作のキーと同じです。", nameof(globalHotkey));
        }

        _chords = chords.ToDictionary(pair => pair.Key, pair => pair.Value);
        GlobalHotkey = globalHotkey;
    }

    /// <summary>
    /// 緊急用の割り当て（設定ファイルがない、または読めないときだけ使う）。通常は、実行ファイルのフォルダの keybindings.json が割り当てを決める。
    /// 付属の keybindings.json と同じ内容にしてあり、テストで食い違いを確かめている。
    /// テンキーがなくても使えるように、主なキーはメインのキーボードに置き、テンキーのキーは別名として足す
    /// </summary>
    public static KeyBindings Default { get; } = new(
        new Dictionary<KeyAction, IReadOnlyList<KeyChord>> {
            [KeyAction.AcquireValue] = [new(Key.Space), new(Key.F2), new(Key.Add)],
            [KeyAction.Confirm] = [new(Key.Return), new(Key.F4)],
            [KeyAction.Reapply] = [new(Key.F3), new(Key.Multiply)],
            [KeyAction.Pass] = [new(Key.F5)],
            [KeyAction.Fail] = [new(Key.F6)],
            [KeyAction.Next] = [new(Key.Down), new(Key.Next)],
            [KeyAction.Previous] = [new(Key.Up), new(Key.Prior)],
        },
        globalHotkey: new KeyChord(Key.OemCloseBrackets));

    /// <summary>新アプリを前面に出してフォーカスする、全体で効くホットキー（使わないときは null）</summary>
    public KeyChord? GlobalHotkey { get; }

    /// <summary>操作に割り当てられたキー（割り当てがなければ空）</summary>
    public IReadOnlyList<KeyChord> ChordsOf(KeyAction action) => _chords.TryGetValue(action, out var chords) ? chords : [];

    /// <summary>キーに割り当てられた操作を探す。なければ false</summary>
    public bool TryFind(KeyChord chord, out KeyAction action) {
        foreach (var (candidate, chords) in _chords) {
            if (chords.Contains(chord)) {
                action = candidate;
                return true;
            }
        }
        action = default;
        return false;
    }

    /// <summary>画面に出す、キーの案内(キーのある操作ごとに、操作の名前と、割り当てたキーすべて)</summary>
    public IReadOnlyList<KeyHintItem> Hints => Labels
        .Where(label => ChordsOf(label.Key).Count > 0)
        .Select(label => new KeyHintItem(label.Value, ChordsOf(label.Key).Select(chord => new KeyBadge(chord.DisplayText)).ToList(), label.Key))
        .ToList();
}
