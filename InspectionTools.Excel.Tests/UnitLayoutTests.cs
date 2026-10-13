namespace InspectionTools.Excel.Tests;

/// <summary>台の位置から書き込むセルを求める計算（折り返しなし・あり）を確認する</summary>
public class UnitLayoutTests {

    // 列の番号（A が 1）
    private const int ColumnC = 3;
    private const int ColumnF = 6;
    private const int ColumnAI = 35;

    [Theory]
    [InlineData(0, ColumnF)]
    [InlineData(1, ColumnF + 1)]
    [InlineData(29, ColumnAI)]
    public void WithoutWrapUnitsAreWrittenToTheRightInTheSameRow(int unitIndex, int expectedColumn) {
        var layout = new UnitLayout("F");

        var cell = layout.GetCell(itemRow: 8, unitIndex);

        Assert.Equal(new CellPosition(Row: 8, Column: expectedColumn), cell);
    }

    // 4台ごとに折り返し、次のまとまりは 20 行下の同じ列から始まる製品
    [Theory]
    [InlineData(0, 8, ColumnC)]
    [InlineData(3, 8, ColumnC + 3)]
    [InlineData(4, 28, ColumnC)]
    [InlineData(9, 48, ColumnC + 1)]
    public void WithWrapUnitsContinueInTheSameColumnsBelow(int unitIndex, int expectedRow, int expectedColumn) {
        var layout = new UnitLayout("C", unitsPerBlock: 4, blockRowOffset: 20);

        var cell = layout.GetCell(itemRow: 8, unitIndex);

        Assert.Equal(new CellPosition(expectedRow, expectedColumn), cell);
    }

    [Fact]
    public void ItemRowIsKeptWhenOnlyOneBlockIsNeeded() {
        var layout = new UnitLayout("C", unitsPerBlock: 4, blockRowOffset: 20);

        Assert.Equal(16, layout.GetCell(itemRow: 16, unitIndex: 2).Row);
    }

    [Theory]
    [InlineData("A", 1)]
    [InlineData("Z", 26)]
    [InlineData("AA", 27)]
    [InlineData("ai", 35)]
    public void ColumnLettersAreConvertedToNumbers(string letters, int expectedNumber) {
        Assert.Equal(expectedNumber, new UnitLayout(letters).FirstUnitColumn);
    }

    [Theory]
    [InlineData("")]
    [InlineData("F1")]
    [InlineData("あ")]
    public void InvalidColumnIsRejected(string letters) {
        Assert.Throws<ArgumentException>(() => new UnitLayout(letters));
    }

    [Fact]
    public void WrapNeedsPositiveUnitsPerBlockAndRowOffset() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new UnitLayout("C", unitsPerBlock: 0, blockRowOffset: 20));
        Assert.Throws<ArgumentOutOfRangeException>(() => new UnitLayout("C", unitsPerBlock: 4, blockRowOffset: 0));
    }

    [Fact]
    public void NegativeUnitIndexIsRejected() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new UnitLayout("F").GetCell(8, -1));
    }
}
