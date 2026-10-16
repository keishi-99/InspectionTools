namespace InspectionTools.Inspection;

/// <summary>1つのセル（項目 × 台）の結果。変更は InspectionSession からだけ行う</summary>
public sealed class ItemResult {

    /// <summary>測定値または入力値（確認項目と作業項目は null）</summary>
    public decimal? Value { get; internal set; }

    /// <summary>入力した文字（文字入力の項目のみ。それ以外は null）</summary>
    public string? Text { get; internal set; }

    /// <summary>セルの状態</summary>
    public ResultState State { get; internal set; } = ResultState.NotDone;

    /// <summary>手動で入力・編集した値か（内部の記録用で、Excel には出さない）</summary>
    public bool IsManual { get; internal set; }
}
