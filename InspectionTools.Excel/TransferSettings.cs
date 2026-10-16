using InspectionTools.Inspection;

namespace InspectionTools.Excel;

/// <summary>転記の設定（製品ごとに決まる）</summary>
/// <param name="DataSheet">結果を書き込むシート（EL5000AC では ACdata）。台を複数のシートに分ける製品では、最初のシート</param>
/// <param name="UnitLayout">台ごとの結果を書き込む列・行の決め方（分けたシートの中での位置）</param>
/// <param name="Check">書き込む前に、正しい製品のテンプレートかを確かめる目印（確かめない製品は null）。シートを分ける製品では、使うシートごとに確かめる</param>
/// <param name="ConfirmationOkText">確認項目が合格のときに書き込む文字（テンプレートの判定式が見る文字）。確認の列を分ける製品では使わない</param>
/// <param name="Split">台を数台ずつ別のシートに書く製品の、シートの分け方（1枚のシートに書く製品は null）</param>
/// <param name="UnitCountCell">各シートに、そのシートで検査した台数を書くセル（書かない製品は null）</param>
public sealed record TransferSettings(
    string DataSheet,
    UnitLayout UnitLayout,
    TemplateCheck? Check,
    string ConfirmationOkText = "良",
    SheetSplit? Split = null,
    string? UnitCountCell = null) {

    /// <summary>製品の定義（JSON の transfer）から、転記の設定を作る</summary>
    public static TransferSettings From(TransferDefinition definition) => new(
        definition.DataSheet,
        new UnitLayout(definition.FirstUnitColumn, definition.UnitsPerBlock, definition.BlockRowOffset, definition.ConfirmationColumn),
        definition.VerifyCell is not null && definition.VerifyValue is not null
            ? new TemplateCheck(definition.DataSheet, definition.VerifyCell, definition.VerifyValue)
            : null,
        Split: definition.DataSheets is not null && definition.UnitsPerSheet is { } unitsPerSheet
            ? new SheetSplit(definition.DataSheets, unitsPerSheet)
            : null,
        UnitCountCell: definition.UnitCountCell);

    /// <summary>unitCount 台の結果を書くのに使うシートの名前を、台の順に返す(使わないシートは含めない)</summary>
    public IReadOnlyList<string> GetSheetNames(int unitCount) {
        if (Split is null) {
            return [DataSheet];
        }

        var sheetCount = (unitCount + Split.UnitsPerSheet - 1) / Split.UnitsPerSheet;
        ArgumentOutOfRangeException.ThrowIfGreaterThan(sheetCount, Split.Sheets.Count, nameof(unitCount));
        return [.. Split.Sheets.Take(sheetCount)];
    }

    /// <summary>台(0から数える)を書くシートの位置(0から数える)</summary>
    public int GetSheetIndex(int unitIndex) => Split is null ? 0 : unitIndex / Split.UnitsPerSheet;

    /// <summary>台(0から数える)の、書くシートの中での位置(0から数える)。UnitLayout には、この位置を渡す</summary>
    public int GetUnitIndexInSheet(int unitIndex) => Split is null ? unitIndex : unitIndex % Split.UnitsPerSheet;
}
