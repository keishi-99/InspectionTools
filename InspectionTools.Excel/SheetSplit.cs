namespace InspectionTools.Excel;

/// <summary>台を数台ずつ別のシートに書く製品の、シートの分け方（1枚目に最初の数台、2枚目に次の数台、のように書く）</summary>
/// <param name="Sheets">シート名の一覧（台の順）</param>
/// <param name="UnitsPerSheet">1枚のシートに書く台数</param>
public sealed record SheetSplit(IReadOnlyList<string> Sheets, int UnitsPerSheet);
