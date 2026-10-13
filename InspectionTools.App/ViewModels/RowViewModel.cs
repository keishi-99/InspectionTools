namespace InspectionTools.App.ViewModels;

/// <summary>表の1行（1つの検査項目）。台ごとのマスを持つ</summary>
public sealed class RowViewModel(string itemName, IReadOnlyList<CellViewModel> cells) {

    /// <summary>項目名</summary>
    public string ItemName { get; } = itemName;

    /// <summary>台ごとのマス（台の順）</summary>
    public IReadOnlyList<CellViewModel> Cells { get; } = cells;
}
