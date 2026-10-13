namespace InspectionTools.Excel;

/// <summary>Excel のセルの位置（行と列。どちらも 1 から数える）</summary>
public readonly record struct CellPosition(int Row, int Column);
