namespace InspectionTools.Inspection;

/// <summary>検査項目の種類</summary>
public enum ItemKind {
    /// <summary>測定器または別ソフトの画面から値を取得し、合格範囲で自動判定する</summary>
    Measurement,
    /// <summary>製品の表示を見て値を入力し、合格範囲で自動判定する</summary>
    Input,
    /// <summary>値はなく、人が合格/不合格を決める</summary>
    Confirmation,
    /// <summary>測定器の操作や製品側の作業だけを行う（Excel に記録しない）</summary>
    Task,
}
