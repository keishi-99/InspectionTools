namespace InspectionTools.Inspection.Tests;

/// <summary>製品ごとの定義（JSON）が、検査項目の一覧になること、正しくない定義は理由つきで弾かれることを確認する</summary>
public class ProductDefinitionLoaderTests {

    // 4つの種類がそろった正しい定義
    private const string ValidJson = """
        {
          "product": "EL5000AC",
          "maxUnitCount": 30,
          "transfer": { "dataSheet": "ACdata", "settingSheet": "ACsetting", "firstRow": 5, "lastRow": 50, "verifyCell": "A3", "verifyValue": "EL5000", "firstUnitColumn": "F" },
          "items": [
            { "name": "電源投入", "kind": "Confirmation", "excelRow": 5 },
            { "name": "消費電流", "kind": "Measurement", "excelRow": 8, "lowerLimit": 0, "upperLimit": 230 },
            { "name": "温度入力 -50℃", "kind": "Input", "excelRow": 16, "lowerLimit": -50.2, "upperLimit": -49.8 },
            { "name": "調整: DCS 4mA", "kind": "Task" }
          ]
        }
        """;

    // JSON の1項目だけを差し替えた定義を作る（正しくない項目を1つだけ試すため）
    private static string JsonWithItems(string items) => $$"""{ "product": "EL5000AC", "maxUnitCount": 30, "transfer": { "dataSheet": "ACdata", "settingSheet": "ACsetting", "firstRow": 5, "lastRow": 50, "verifyCell": "A3", "verifyValue": "EL5000", "firstUnitColumn": "F" }, "items": [ {{items}} ] }""";

    private static ProductDefinitionException AssertRejected(string json) =>
        Assert.Throws<ProductDefinitionException>(() => ProductDefinitionLoader.Parse(json));

    [Fact]
    public void ParseReadsProductAndItemsInOrder() {
        var definition = ProductDefinitionLoader.Parse(ValidJson);

        Assert.Equal("EL5000AC", definition.Product);
        Assert.Equal(["電源投入", "消費電流", "温度入力 -50℃", "調整: DCS 4mA"], definition.Items.Select(item => item.Name));
        Assert.Equal([ItemKind.Confirmation, ItemKind.Measurement, ItemKind.Input, ItemKind.Task], definition.Items.Select(item => item.Kind));
    }

    [Fact]
    public void ParseReadsRangeAndExcelRow() {
        var definition = ProductDefinitionLoader.Parse(ValidJson);

        var input = definition.Items[2];
        Assert.Equal(16, input.ExcelRow);
        Assert.Equal(-50.2m, input.LowerLimit);
        Assert.Equal(-49.8m, input.UpperLimit);
        Assert.Null(definition.Items[3].ExcelRow);
    }

    [Fact]
    public void ParsedItemsCanBeUsedForSession() {
        var definition = ProductDefinitionLoader.Parse(ValidJson);

        var session = new InspectionSession(definition.Items, unitCount: 2);
        session.SetValue(1, 0, 146.745m, isManual: false);

        Assert.Equal(ResultState.Ok, session.GetResult(1, 0).State);
    }

    [Theory]
    [InlineData("{ これはJSONではない")]
    [InlineData("null")]
    [InlineData("")]
    public void BrokenJsonIsRejected(string json) {
        AssertRejected(json);
    }

    [Fact]
    public void MissingProductIsRejected() {
        var exception = AssertRejected("""{ "items": [ { "name": "電源投入", "kind": "Confirmation", "excelRow": 5 } ] }""");

        Assert.Contains("製品名", exception.Message);
    }

    [Fact]
    public void ParseReadsMaxUnitCount() {
        Assert.Equal(30, ProductDefinitionLoader.Parse(ValidJson).MaxUnitCount);
    }

    [Theory]
    [InlineData("""{ "product": "EL5000AC", "items": [ { "name": "作業", "kind": "Task" } ] }""")]
    [InlineData("""{ "product": "EL5000AC", "maxUnitCount": 0, "items": [ { "name": "作業", "kind": "Task" } ] }""")]
    [InlineData("""{ "product": "EL5000AC", "maxUnitCount": -1, "items": [ { "name": "作業", "kind": "Task" } ] }""")]
    public void MissingOrInvalidMaxUnitCountIsRejected(string json) {
        var exception = AssertRejected(json);

        Assert.Contains("maxUnitCount", exception.Message);
    }

    [Fact]
    public void ParseReadsTransferSettings() {
        var transfer = ProductDefinitionLoader.Parse(ValidJson).Transfer;

        Assert.Equal(new TransferDefinition("ACdata", "ACsetting", 5, 50, "A3", "EL5000", "F", UnitsPerBlock: null, BlockRowOffset: 0), transfer);
    }

    [Fact]
    public void ParseReadsWrappedLayout() {
        var json = ValidJson.Replace("\"firstUnitColumn\": \"F\"", "\"firstUnitColumn\": \"C\", \"unitsPerBlock\": 4, \"blockRowOffset\": 20");

        var transfer = ProductDefinitionLoader.Parse(json).Transfer;

        Assert.Equal(4, transfer.UnitsPerBlock);
        Assert.Equal(20, transfer.BlockRowOffset);
    }

    [Fact]
    public void MissingTransferIsRejected() {
        var exception = AssertRejected("""{ "product": "EL5000AC", "maxUnitCount": 30, "items": [ { "name": "作業", "kind": "Task" } ] }""");

        Assert.Contains("transfer", exception.Message);
    }

    [Theory]
    [InlineData("\"dataSheet\": \"ACdata\"", "\"dataSheet\": \"\"", "transfer.dataSheet")]
    [InlineData("\"settingSheet\": \"ACsetting\"", "\"settingSheet\": \" \"", "transfer.settingSheet")]
    [InlineData("\"verifyCell\": \"A3\"", "\"verifyCell\": null", "transfer.verifyCell")]
    [InlineData("\"verifyValue\": \"EL5000\"", "\"verifyValue\": \"\"", "transfer.verifyValue")]
    [InlineData("\"firstUnitColumn\": \"F\"", "\"firstUnitColumn\": \"5\"", "transfer.firstUnitColumn")]
    [InlineData("\"firstRow\": 5", "\"firstRow\": 0", "transfer.firstRow")]
    [InlineData("\"lastRow\": 50", "\"lastRow\": 4", "transfer.lastRow")]
    [InlineData("\"firstUnitColumn\": \"F\"", "\"firstUnitColumn\": \"F\", \"unitsPerBlock\": 4", "unitsPerBlock")]
    [InlineData("\"firstUnitColumn\": \"F\"", "\"firstUnitColumn\": \"F\", \"unitsPerBlock\": 0, \"blockRowOffset\": 20", "unitsPerBlock")]
    [InlineData("\"firstUnitColumn\": \"F\"", "\"firstUnitColumn\": \"F\", \"unitsPerBlock\": 4, \"blockRowOffset\": 0", "blockRowOffset")]
    public void InvalidTransferSettingsAreRejectedWithTheirName(string original, string replacement, string expectedName) {
        var json = ValidJson.Replace(original, replacement);
        Assert.NotEqual(ValidJson, json);

        var exception = AssertRejected(json);

        Assert.Contains(expectedName, exception.Message);
    }

    [Fact]
    public void EmptyItemsAreRejected() {
        AssertRejected("""{ "product": "EL5000AC", "maxUnitCount": 30, "items": [] }""");
        AssertRejected("""{ "product": "EL5000AC", "maxUnitCount": 30 }""");
    }

    [Fact]
    public void MissingKindIsRejectedWithItemPosition() {
        var exception = AssertRejected(JsonWithItems("""{ "name": "電源投入", "excelRow": 5 }"""));

        Assert.Contains("1番目の項目", exception.Message);
        Assert.Contains("kind", exception.Message);
    }

    [Fact]
    public void UnknownKindIsRejected() {
        AssertRejected(JsonWithItems("""{ "name": "電源投入", "kind": "Unknown", "excelRow": 5 }"""));
    }

    [Fact]
    public void RuleViolationOfItemIsReportedWithItemPositionAndName() {
        // 2番目の項目が、測定項目なのに合格範囲がない
        var exception = AssertRejected(JsonWithItems("""
            { "name": "電源投入", "kind": "Confirmation", "excelRow": 5 },
            { "name": "消費電流", "kind": "Measurement", "excelRow": 8 }
            """));

        Assert.Contains("2番目の項目", exception.Message);
        Assert.Contains("消費電流", exception.Message);
    }

    [Fact]
    public void TaskItemWithExcelRowIsRejected() {
        AssertRejected(JsonWithItems("""{ "name": "調整: DCS 4mA", "kind": "Task", "excelRow": 5 }"""));
    }

    [Fact]
    public void DuplicateExcelRowIsRejectedWithNames() {
        var exception = AssertRejected(JsonWithItems("""
            { "name": "電源投入", "kind": "Confirmation", "excelRow": 5 },
            { "name": "LCD表示確認", "kind": "Confirmation", "excelRow": 5 }
            """));

        Assert.Contains("5", exception.Message);
        Assert.Contains("電源投入", exception.Message);
        Assert.Contains("LCD表示確認", exception.Message);
    }

    [Fact]
    public void TaskItemsMayAppearManyTimesBecauseTheyHaveNoRow() {
        var definition = ProductDefinitionLoader.Parse(JsonWithItems("""
            { "name": "調整: DCS 4mA", "kind": "Task" },
            { "name": "調整: DCS 20mA", "kind": "Task" }
            """));

        Assert.Equal(2, definition.Items.Count);
    }
}
