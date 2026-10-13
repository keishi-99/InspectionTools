namespace InspectionTools.Inspection;

/// <summary>1つのセル（項目 × 台）の状態</summary>
public enum ResultState {
    /// <summary>未実施（Excel では空欄。NG と見分けるために別の状態にしている）</summary>
    NotDone,
    /// <summary>合格</summary>
    Ok,
    /// <summary>不合格</summary>
    Ng,
    /// <summary>作業項目が完了した</summary>
    Done,
}
