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
    public void ParseReadsActionIdAndLeavesItNullWhenAbsent() {
        var definition = ProductDefinitionLoader.Parse(JsonWithItems(
            """{ "name": "調整", "kind": "Task", "actionId": "dcs-4mA" }, { "name": "確認", "kind": "Confirmation", "excelRow": 5 }"""));

        Assert.Equal("dcs-4mA", definition.Items[0].ActionId);
        Assert.Null(definition.Items[1].ActionId);
    }

    [Fact]
    public void ParseReadsDescriptionAndLeavesItNullWhenAbsent() {
        var definition = ProductDefinitionLoader.Parse(JsonWithItems(
            """{ "name": "調整", "kind": "Task", "description": "DCS を 4mA にして調整する" }, { "name": "確認", "kind": "Confirmation", "excelRow": 5 }"""));

        Assert.Equal("DCS を 4mA にして調整する", definition.Items[0].Description);
        Assert.Null(definition.Items[1].Description);
    }

    [Fact]
    public void BlankDescriptionIsRejectedWithItemPosition() {
        var exception = AssertRejected(JsonWithItems("""{ "name": "調整", "kind": "Task", "description": " " }"""));

        Assert.Contains("1番目の項目", exception.Message);
        Assert.Contains("description", exception.Message);
    }

    [Fact]
    public void BlankActionIdIsRejectedWithItemPosition() {
        var exception = AssertRejected(JsonWithItems("""{ "name": "調整", "kind": "Task", "actionId": " " }"""));

        Assert.Contains("1番目の項目", exception.Message);
        Assert.Contains("actionId", exception.Message);
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

    // ===== 台を5台ずつ別のシートに書く製品（dataSheets） =====

    // dataSheet の代わりに dataSheets と unitsPerSheet を書いた定義(2シート × 5台 = 最大10台)
    private static string MultiSheetJson(string transferExtra = "", string maxUnitCount = "10", string sheets = """["分離形　1～5", "分離形　6～10"]""") => $$"""
        {
          "product": "PAF5-分離形",
          "maxUnitCount": {{maxUnitCount}},
          "transfer": { "dataSheets": {{sheets}}, "unitsPerSheet": 5, "firstRow": 7, "lastRow": 48, "firstUnitColumn": "H" {{transferExtra}} },
          "items": [ { "name": "電源投入", "kind": "Confirmation", "excelRow": 8 } ]
        }
        """;

    [Fact]
    public void ParseReadsDataSheetsAndUnitsPerSheet() {
        var transfer = ProductDefinitionLoader.Parse(MultiSheetJson()).Transfer;

        Assert.Equal(["分離形　1～5", "分離形　6～10"], transfer.DataSheets);
        Assert.Equal(5, transfer.UnitsPerSheet);
        // 最初のシートが、これまでの dataSheet(確認用のセルがあるシート)になる
        Assert.Equal("分離形　1～5", transfer.DataSheet);
    }

    [Fact]
    public void SingleSheetDefinitionHasNoDataSheetsAndNoUnitsPerSheet() {
        var transfer = ProductDefinitionLoader.Parse(ValidJson).Transfer;

        Assert.Null(transfer.DataSheets);
        Assert.Null(transfer.UnitsPerSheet);
    }

    [Fact]
    public void DataSheetAndDataSheetsTogetherAreRejected() {
        var exception = AssertRejected(MultiSheetJson(transferExtra: ", \"dataSheet\": \"ACdata\""));

        Assert.Contains("transfer.dataSheets", exception.Message);
    }

    [Fact]
    public void DataSheetsWithoutUnitsPerSheetAreRejected() {
        var json = MultiSheetJson().Replace("\"unitsPerSheet\": 5, ", "");
        Assert.NotEqual(MultiSheetJson(), json);

        var exception = AssertRejected(json);

        Assert.Contains("transfer.unitsPerSheet", exception.Message);
    }

    [Fact]
    public void UnitsPerSheetWithoutDataSheetsIsRejected() {
        var json = ValidJson.Replace("\"firstUnitColumn\": \"F\"", "\"firstUnitColumn\": \"F\", \"unitsPerSheet\": 5");

        var exception = AssertRejected(json);

        Assert.Contains("transfer.unitsPerSheet", exception.Message);
    }

    [Theory]
    [InlineData("[]")]
    [InlineData("""["分離形　1～5", " "]""")]
    [InlineData("""["分離形　1～5", "分離形　1～5"]""")]
    public void EmptyBlankOrDuplicatedSheetNamesAreRejected(string sheets) {
        var exception = AssertRejected(MultiSheetJson(sheets: sheets));

        Assert.Contains("transfer.dataSheets", exception.Message);
    }

    [Theory]
    [InlineData("0")]
    [InlineData("-5")]
    public void NonPositiveUnitsPerSheetIsRejected(string unitsPerSheet) {
        var json = MultiSheetJson().Replace("\"unitsPerSheet\": 5", $"\"unitsPerSheet\": {unitsPerSheet}");

        var exception = AssertRejected(json);

        Assert.Contains("transfer.unitsPerSheet", exception.Message);
    }

    // シートが足りないと、最大台数の最後の台を書く先がなくなる
    [Fact]
    public void MaxUnitCountOverTheCapacityOfTheSheetsIsRejected() {
        var exception = AssertRejected(MultiSheetJson(maxUnitCount: "11"));

        Assert.Contains("maxUnitCount", exception.Message);
        Assert.Contains("10", exception.Message);
    }

    [Fact]
    public void MaxUnitCountWithinTheCapacityIsAccepted() {
        Assert.Equal(10, ProductDefinitionLoader.Parse(MultiSheetJson(maxUnitCount: "10")).MaxUnitCount);
        Assert.Equal(7, ProductDefinitionLoader.Parse(MultiSheetJson(maxUnitCount: "7")).MaxUnitCount);
    }

    [Fact]
    public void DataSheetsWithWrapAreRejected() {
        var exception = AssertRejected(MultiSheetJson(transferExtra: ", \"unitsPerBlock\": 4, \"blockRowOffset\": 20"));

        Assert.Contains("unitsPerBlock", exception.Message);
    }

    [Fact]
    public void ParseReadsUnitCountCellAndLeavesItNullWhenAbsent() {
        Assert.Equal("C1", ProductDefinitionLoader.Parse(MultiSheetJson(transferExtra: ", \"unitCountCell\": \"C1\"")).Transfer.UnitCountCell);
        Assert.Null(ProductDefinitionLoader.Parse(MultiSheetJson()).Transfer.UnitCountCell);
    }

    [Theory]
    [InlineData("1C")]
    [InlineData("C")]
    [InlineData("C0")]
    [InlineData("C1:D2")]
    public void InvalidUnitCountCellIsRejectedWithItsName(string cell) {
        var exception = AssertRejected(MultiSheetJson(transferExtra: $", \"unitCountCell\": \"{cell}\""));

        Assert.Contains("transfer.unitCountCell", exception.Message);
    }

    [Fact]
    public void ParseReadsTextKind() {
        var definition = ProductDefinitionLoader.Parse(JsonWithItems("""{ "name": "LCDシリアル", "kind": "Text", "excelRow": 7 }"""));

        Assert.Equal(ItemKind.Text, definition.Items[0].Kind);
        Assert.Equal(7, definition.Items[0].ExcelRow);
    }

    [Fact]
    public void TextItemWithRangeIsRejectedWithItemPosition() {
        var exception = AssertRejected(JsonWithItems("""{ "name": "LCDシリアル", "kind": "Text", "excelRow": 7, "lowerLimit": 0, "upperLimit": 1 }"""));

        Assert.Contains("1番目の項目", exception.Message);
    }

    [Fact]
    public void SettingSheetMayBeOmitted() {
        var json = ValidJson.Replace("\"settingSheet\": \"ACsetting\", ", "");
        Assert.NotEqual(ValidJson, json);

        var transfer = ProductDefinitionLoader.Parse(json).Transfer;

        Assert.Null(transfer.SettingSheet);
    }

    [Fact]
    public void VerifyCellAndValueMayBothBeOmitted() {
        var json = ValidJson.Replace("\"verifyCell\": \"A3\", \"verifyValue\": \"EL5000\", ", "");
        Assert.NotEqual(ValidJson, json);

        var transfer = ProductDefinitionLoader.Parse(json).Transfer;

        Assert.Null(transfer.VerifyCell);
        Assert.Null(transfer.VerifyValue);
    }

    // 片方だけでは、何を何と比べるのか決まらないので、書いていない方の名前を添えて弾く
    [Theory]
    [InlineData("\"verifyCell\": \"A3\", ", "transfer.verifyCell")]
    [InlineData("\"verifyValue\": \"EL5000\", ", "transfer.verifyValue")]
    public void OnlyOneOfVerifyCellAndValueIsRejectedWithTheMissingName(string removed, string expectedName) {
        var json = ValidJson.Replace(removed, "");
        Assert.NotEqual(ValidJson, json);

        var exception = AssertRejected(json);

        Assert.Contains(expectedName, exception.Message);
    }

    [Fact]
    public void ParseReadsConfirmationColumnAndLeavesItNullWhenAbsent() {
        var json = ValidJson.Replace("\"firstUnitColumn\": \"F\"", "\"firstUnitColumn\": \"H\", \"confirmationColumn\": \"M\"");

        Assert.Equal("M", ProductDefinitionLoader.Parse(json).Transfer.ConfirmationColumn);
        Assert.Null(ProductDefinitionLoader.Parse(ValidJson).Transfer.ConfirmationColumn);
    }

    [Theory]
    [InlineData("5")]
    [InlineData("M1")]
    [InlineData(" ")]
    public void InvalidConfirmationColumnIsRejectedWithItsName(string column) {
        var json = ValidJson.Replace("\"firstUnitColumn\": \"F\"", $"\"firstUnitColumn\": \"F\", \"confirmationColumn\": \"{column}\"");

        var exception = AssertRejected(json);

        Assert.Contains("transfer.confirmationColumn", exception.Message);
    }

    [Fact]
    public void ParseReadsDefaultValueAndLeavesItNullWhenAbsent() {
        var definition = ProductDefinitionLoader.Parse(JsonWithItems("""
            { "name": "ノッチ周波数", "kind": "Input", "excelRow": 23, "lowerLimit": 6250, "upperLimit": 6250, "defaultValue": 6250 },
            { "name": "消費電流", "kind": "Measurement", "excelRow": 8, "lowerLimit": 0, "upperLimit": 230 }
            """));

        Assert.Equal(6250m, definition.Items[0].DefaultValue);
        Assert.Null(definition.Items[1].DefaultValue);
    }

    // 初期値を入れる欄があるのは、値を入力する項目（測定、入力）だけ
    [Theory]
    [InlineData("""{ "name": "電源投入", "kind": "Confirmation", "excelRow": 5, "defaultValue": 1 }""")]
    [InlineData("""{ "name": "調整", "kind": "Task", "defaultValue": 1 }""")]
    public void DefaultValueOnConfirmationOrTaskIsRejectedWithItemPosition(string item) {
        var exception = AssertRejected(JsonWithItems(item));

        Assert.Contains("1番目の項目", exception.Message);
        Assert.Contains("defaultValue", exception.Message);
    }
}
