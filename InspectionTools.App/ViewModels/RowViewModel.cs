using CommunityToolkit.Mvvm.ComponentModel;

namespace InspectionTools.App.ViewModels;

/// <summary>表の1行（1つの検査項目）。台ごとのマスを持つ</summary>
public sealed partial class RowViewModel(string itemName, IReadOnlyList<CellViewModel> cells) : ObservableObject {

    /// <summary>項目名</summary>
    public string ItemName { get; } = itemName;

    /// <summary>台ごとのマス（台の順）</summary>
    public IReadOnlyList<CellViewModel> Cells { get; } = cells;

    /// <summary>いま検査している項目の行か（表で目立たせる）</summary>
    [ObservableProperty]
    private bool _isCurrent;
}
