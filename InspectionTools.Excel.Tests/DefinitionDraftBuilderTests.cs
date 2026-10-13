using InspectionTools.Inspection;

namespace InspectionTools.Excel.Tests;

/// <summary>テンプレートから読み取った行が、製品の定義（JSON）のひな形になることを確認する</summary>
public class DefinitionDraftBuilderTests {

    private static readonly List<TemplateRow> Rows = [
        new(Row: 5, Name: "電源投入", LowerLimit: null, UpperLimit: null, IsConfirmation: true),
        new(Row: 8, Name: "消費電流", LowerLimit: 0m, UpperLimit: 230m, IsConfirmation: false),
        new(Row: 16, Name: "温度入力 -50℃", LowerLimit: -50.2m, UpperLimit: -49.8m, IsConfirmation: false),
    ];

    private static readonly TransferDefinition Transfer = new("ACdata", "ACsetting", 5, 50, "A3", "EL5000", "F", UnitsPerBlock: null, BlockRowOffset: 0);

    private static HashSet<int> MeasurementRows(params int[] rows) => [.. rows];

    [Fact]
    public void KindIsConfirmationMeasurementOrInput() {
        var json = DefinitionDraftBuilder.BuildJson("EL5000AC", Rows, MeasurementRows(8), maxUnitCount: 30, Transfer);

        var definition = ProductDefinitionLoader.Parse(json);

        Assert.Equal([ItemKind.Confirmation, ItemKind.Measurement, ItemKind.Input], definition.Items.Select(item => item.Kind));
    }

    [Fact]
    public void DraftKeepsNameRowAndRange() {
        var json = DefinitionDraftBuilder.BuildJson("EL5000AC", Rows, MeasurementRows(8), maxUnitCount: 30, Transfer);

        var item = ProductDefinitionLoader.Parse(json).Items[2];

        Assert.Equal("温度入力 -50℃", item.Name);
        Assert.Equal(16, item.ExcelRow);
        Assert.Equal(-50.2m, item.LowerLimit);
        Assert.Equal(-49.8m, item.UpperLimit);
    }

    [Fact]
    public void JapaneseTextIsNotEscaped() {
        var json = DefinitionDraftBuilder.BuildJson("EL5000AC", Rows, MeasurementRows(8), maxUnitCount: 30, Transfer);

        Assert.Contains("消費電流", json);
        Assert.DoesNotContain("\\u", json);
    }

    [Fact]
    public void MeasurementRowNotInTemplateThrows() {
        Assert.Throws<ArgumentException>(() => DefinitionDraftBuilder.BuildJson("EL5000AC", Rows, MeasurementRows(99), maxUnitCount: 30, Transfer));
    }

    [Fact]
    public void MeasurementRowOfConfirmationThrows() {
        Assert.Throws<ArgumentException>(() => DefinitionDraftBuilder.BuildJson("EL5000AC", Rows, MeasurementRows(5), maxUnitCount: 30, Transfer));
    }

    [Fact]
    public void ParseRowsAcceptsRangesAndSingleRows() {
        var rows = DefinitionDraftBuilder.ParseRows("8-10, 30,32-33");

        Assert.Equal([8, 9, 10, 30, 32, 33], rows.OrderBy(row => row));
    }

    [Theory]
    [InlineData("a-b")]
    [InlineData("10-8")]
    [InlineData("1-2-3")]
    [InlineData("0")]
    public void ParseRowsRejectsInvalidText(string text) {
        Assert.Throws<FormatException>(() => DefinitionDraftBuilder.ParseRows(text));
    }
}
