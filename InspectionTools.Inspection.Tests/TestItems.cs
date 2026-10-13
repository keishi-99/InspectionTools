namespace InspectionTools.Inspection.Tests;

/// <summary>テストで共通して使う検査項目の定義</summary>
internal static class TestItems {

    /// <summary>項目の位置（0: 測定、1: 入力、2: 確認、3: 作業）の並びで、4項目を作る</summary>
    public static List<InspectionItem> Create() => [
        new("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m),
        new("温度入力", ItemKind.Input, excelRow: 16, lowerLimit: -50.2m, upperLimit: -49.8m),
        new("電源投入", ItemKind.Confirmation, excelRow: 5),
        new("調整作業", ItemKind.Task),
    ];
}
