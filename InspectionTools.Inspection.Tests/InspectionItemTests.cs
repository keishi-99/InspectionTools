namespace InspectionTools.Inspection.Tests;

/// <summary>検査項目の定義が、種類ごとの決まり（Excel の行、合格範囲）を守っていることを確認する</summary>
public class InspectionItemTests {

    [Fact]
    public void MeasurementItemKeepsRangeAndRow() {
        var item = new InspectionItem("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m);

        Assert.Equal(8, item.ExcelRow);
        Assert.Equal(0m, item.LowerLimit);
        Assert.Equal(230m, item.UpperLimit);
    }

    [Theory]
    [InlineData(ItemKind.Measurement)]
    [InlineData(ItemKind.Input)]
    public void ValueItemWithoutRangeThrows(ItemKind kind) {
        Assert.Throws<ArgumentException>(() => new InspectionItem("電圧", kind, excelRow: 9));
    }

    // 文字入力の項目（シリアルなど）は、Excel の行を持ち、合格範囲は持たない
    [Fact]
    public void TextItemKeepsRowAndHasNoRange() {
        var item = new InspectionItem("LCDシリアル", ItemKind.Text, excelRow: 7);

        Assert.Equal(7, item.ExcelRow);
        Assert.Null(item.LowerLimit);
        Assert.Null(item.UpperLimit);
    }

    [Fact]
    public void TextItemWithoutExcelRowThrows() {
        Assert.Throws<ArgumentException>(() => new InspectionItem("LCDシリアル", ItemKind.Text));
    }

    [Theory]
    [InlineData(0, null)]
    [InlineData(null, 10)]
    [InlineData(0, 10)]
    public void TextItemWithRangeThrows(int? lower, int? upper) {
        Assert.Throws<ArgumentException>(() => new InspectionItem("LCDシリアル", ItemKind.Text, excelRow: 7, lowerLimit: lower, upperLimit: upper));
    }

    [Fact]
    public void LowerLimitGreaterThanUpperLimitThrows() {
        Assert.Throws<ArgumentException>(() => new InspectionItem("電圧", ItemKind.Input, excelRow: 9, lowerLimit: 5m, upperLimit: 1m));
    }

    [Fact]
    public void TaskItemWithExcelRowThrows() {
        Assert.Throws<ArgumentException>(() => new InspectionItem("調整作業", ItemKind.Task, excelRow: 5));
    }

    [Theory]
    [InlineData(ItemKind.Measurement)]
    [InlineData(ItemKind.Input)]
    [InlineData(ItemKind.Confirmation)]
    public void RecordedItemWithoutExcelRowThrows(ItemKind kind) {
        Assert.Throws<ArgumentException>(() => new InspectionItem("項目", kind, excelRow: null, lowerLimit: 0m, upperLimit: 1m));
    }

    [Fact]
    public void TaskItemAndConfirmationItemNeedNoRange() {
        var task = new InspectionItem("調整作業", ItemKind.Task);
        var confirmation = new InspectionItem("電源投入", ItemKind.Confirmation, excelRow: 5);

        Assert.Null(task.ExcelRow);
        Assert.Null(confirmation.LowerLimit);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void BlankNameThrows(string name) {
        Assert.Throws<ArgumentException>(() => new InspectionItem(name, ItemKind.Confirmation, excelRow: 5));
    }
}
