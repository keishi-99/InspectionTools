namespace InspectionTools.Excel;

/// <summary>テンプレートから読み取った1行分の検査項目の情報</summary>
/// <param name="Row">Excel の行</param>
/// <param name="Name">項目名（大項目と小項目を合わせたもの）</param>
/// <param name="LowerLimit">合格範囲の下限（確認項目は null）</param>
/// <param name="UpperLimit">合格範囲の上限（確認項目は null）</param>
/// <param name="IsConfirmation">合格範囲が数値ではなく「良」などの文字になっている行か（確認項目の目印）</param>
public sealed record TemplateRow(int Row, string Name, decimal? LowerLimit, decimal? UpperLimit, bool IsConfirmation);
