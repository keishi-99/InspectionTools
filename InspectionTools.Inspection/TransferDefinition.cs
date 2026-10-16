namespace InspectionTools.Inspection;

/// <summary>
/// 検査結果をExcelのテンプレートへ転記するための、製品ごとの設定（JSON の transfer）。
/// Excel の部品（ClosedXML など）に依存しないよう、文字列と数値だけで持つ
/// </summary>
/// <param name="DataSheet">結果を書き込むシート名（製品の確認用のセルもこのシートにある）。台を複数のシートに分ける製品では、最初のシート</param>
/// <param name="SettingSheet">合格範囲が入っているシート名（合格範囲のシートがないテンプレートは null）</param>
/// <param name="FirstRow">読み取る最初の行</param>
/// <param name="LastRow">読み取る最後の行</param>
/// <param name="VerifyCell">正しい製品のテンプレートか確認するセル（例: A3）。確認しない製品は null</param>
/// <param name="VerifyValue">確認用のセルに含まれているはずの値。確認しない製品は null（VerifyCell と、両方ある、または両方ない）</param>
/// <param name="FirstUnitColumn">1台目を書き込む列（例: F）</param>
/// <param name="UnitsPerBlock">折り返すまでの台数（折り返さない製品は null）</param>
/// <param name="BlockRowOffset">折り返したときに下へ進める行数（折り返さない製品は 0）</param>
/// <param name="ConfirmationColumn">確認項目の1台目を書き込む列（チェックボックスのリンク先など。値の列と同じセルに書く製品は null）</param>
/// <param name="DataSheets">台を数台ずつ別のシートに書く製品の、シート名の一覧（台の順。1枚のシートに書く製品は null）</param>
/// <param name="UnitsPerSheet">1枚のシートに書く台数（DataSheets と、両方ある、または両方ない）</param>
/// <param name="UnitCountCell">各シートに、そのシートで検査した台数を書くセル（例: C1。書かない製品は null）</param>
public sealed record TransferDefinition(
    string DataSheet,
    string? SettingSheet,
    int FirstRow,
    int LastRow,
    string? VerifyCell,
    string? VerifyValue,
    string FirstUnitColumn,
    int? UnitsPerBlock,
    int BlockRowOffset,
    string? ConfirmationColumn = null,
    IReadOnlyList<string>? DataSheets = null,
    int? UnitsPerSheet = null,
    string? UnitCountCell = null);
