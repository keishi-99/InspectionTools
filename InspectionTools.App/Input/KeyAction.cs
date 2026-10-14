namespace InspectionTools.App.Input;

/// <summary>キーで行える操作（新アプリにフォーカスがあるときだけ効く）</summary>
public enum KeyAction {
    /// <summary>現在の項目の測定器から値を取得して、入力欄に出す</summary>
    AcquireValue,
    /// <summary>値の項目は確定、作業項目は作業完了にする（確認項目では何もしない。合否は Pass / Fail で明示する）</summary>
    Confirm,
    /// <summary>確認項目を合格にする</summary>
    Pass,
    /// <summary>確認項目を不合格にする</summary>
    Fail,
    /// <summary>現在の項目の測定器の状態を、記録を無視して送り直す</summary>
    Reapply,
    /// <summary>次の項目へ移る</summary>
    Next,
    /// <summary>前の項目へ移る</summary>
    Previous,
}
