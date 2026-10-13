namespace InspectionTools.Excel;

/// <summary>定義（JSON）とテンプレートの合格範囲の食い違い1件</summary>
/// <param name="Row">Excel の行</param>
/// <param name="ItemName">項目名</param>
/// <param name="Reason">何が食い違っているか</param>
public sealed record RangeMismatch(int Row, string ItemName, string Reason);
