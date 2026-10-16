using InspectionTools.Inspection;

namespace InspectionTools.Excel;

/// <summary>転記の設定（製品ごとに決まる）</summary>
/// <param name="DataSheet">結果を書き込むシート（EL5000AC では ACdata）</param>
/// <param name="UnitLayout">台ごとの結果を書き込む列・行の決め方</param>
/// <param name="Check">書き込む前に、正しい製品のテンプレートかを確かめる目印（確かめない製品は null）</param>
/// <param name="ConfirmationOkText">確認項目が合格のときに書き込む文字（テンプレートの判定式が見る文字）。確認の列を分ける製品では使わない</param>
public sealed record TransferSettings(string DataSheet, UnitLayout UnitLayout, TemplateCheck? Check, string ConfirmationOkText = "良") {

    /// <summary>製品の定義（JSON の transfer）から、転記の設定を作る</summary>
    public static TransferSettings From(TransferDefinition definition) => new(
        definition.DataSheet,
        new UnitLayout(definition.FirstUnitColumn, definition.UnitsPerBlock, definition.BlockRowOffset, definition.ConfirmationColumn),
        definition.VerifyCell is not null && definition.VerifyValue is not null
            ? new TemplateCheck(definition.DataSheet, definition.VerifyCell, definition.VerifyValue)
            : null);
}
