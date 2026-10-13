namespace InspectionTools.Inspection;

/// <summary>途中保存する検査中の状態（項目の定義は含めない。定義は製品ごとの JSON から毎回読み込めるため）</summary>
public sealed record SessionSnapshot(
    string Product,
    int ItemCount,
    int UnitCount,
    int CurrentItemIndex,
    int CurrentUnitIndex,
    IReadOnlyList<CellSnapshot> Cells);

/// <summary>実施済みの1セル（項目 × 台）の保存内容（未実施のセルは保存しない）</summary>
public sealed record CellSnapshot(int Unit, int Item, decimal? Value, ResultState State, bool IsManual);
