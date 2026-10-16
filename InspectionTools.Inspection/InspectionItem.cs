namespace InspectionTools.Inspection;

/// <summary>検査項目の定義（製品ごとに固定で、検査中に変わらない）</summary>
public sealed class InspectionItem {

    public InspectionItem(string name, ItemKind kind, int? excelRow = null, decimal? lowerLimit = null, decimal? upperLimit = null, string? actionId = null, string? description = null, decimal? defaultValue = null) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("項目名が空です。", nameof(name));
        }
        if (actionId is not null && string.IsNullOrWhiteSpace(actionId)) {
            throw new ArgumentException($"動作ID（actionId）が空です: {name}", nameof(actionId));
        }
        if (description is not null && string.IsNullOrWhiteSpace(description)) {
            throw new ArgumentException($"説明（description）が空です: {name}", nameof(description));
        }

        // 作業項目は Excel の行に対応しない（記録しないため）。それ以外は転記先の行が必要
        if (kind == ItemKind.Task && excelRow is not null) {
            throw new ArgumentException($"作業項目には Excel の行を指定できません: {name}", nameof(excelRow));
        }
        if (kind != ItemKind.Task && excelRow is null) {
            throw new ArgumentException($"Excel の行が必要です: {name}", nameof(excelRow));
        }

        // 値で自動判定する項目は、合格範囲（下限と上限）が必要
        if (kind is ItemKind.Measurement or ItemKind.Input) {
            if (lowerLimit is null || upperLimit is null) {
                throw new ArgumentException($"合格範囲（下限と上限）が必要です: {name}");
            }
            if (lowerLimit > upperLimit) {
                throw new ArgumentException($"下限が上限より大きくなっています: {name}");
            }
        }

        // 文字入力の項目は、合格範囲で判定しない
        if (kind == ItemKind.Text && (lowerLimit is not null || upperLimit is not null)) {
            throw new ArgumentException($"文字入力の項目には合格範囲を指定できません: {name}");
        }

        // 初期値を入れる欄があるのは、値を入力する項目だけ
        if (defaultValue is not null && kind is not (ItemKind.Measurement or ItemKind.Input)) {
            throw new ArgumentException($"初期値（defaultValue）は、測定項目と入力項目にだけ指定できます: {name}", nameof(defaultValue));
        }

        Name = name;
        Kind = kind;
        ExcelRow = excelRow;
        LowerLimit = lowerLimit;
        UpperLimit = upperLimit;
        ActionId = actionId;
        Description = description;
        DefaultValue = defaultValue;
    }

    /// <summary>画面に出す項目名</summary>
    public string Name { get; }

    /// <summary>項目の種類</summary>
    public ItemKind Kind { get; }

    /// <summary>転記先の Excel の行（作業項目は null）</summary>
    public int? ExcelRow { get; }

    /// <summary>合格範囲の下限（測定項目と入力項目のみ）</summary>
    public decimal? LowerLimit { get; }

    /// <summary>合格範囲の上限（測定項目と入力項目のみ）</summary>
    public decimal? UpperLimit { get; }

    /// <summary>この項目で測定器に行わせる動作の ID（製品ごとの手順表を引くための鍵。測定器を使わない項目は null）</summary>
    public string? ActionId { get; }

    /// <summary>画面に出す、項目の説明（作業項目の手順など。なければ null）</summary>
    public string? Description { get; }

    /// <summary>入力欄に最初から出しておく値（固定の値を入れる項目など。なければ null）</summary>
    public decimal? DefaultValue { get; }
}
