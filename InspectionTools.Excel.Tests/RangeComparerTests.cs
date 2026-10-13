using InspectionTools.Inspection;

namespace InspectionTools.Excel.Tests;

/// <summary>定義（JSON）とテンプレートの合格範囲の食い違いを、見つけられることを確認する</summary>
public class RangeComparerTests {

    // 定義とテンプレートが一致している基準の組み合わせ
    private static List<InspectionItem> CreateItems() => [
        new("電源投入", ItemKind.Confirmation, excelRow: 5),
        new("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m),
        new("温度入力", ItemKind.Input, excelRow: 16, lowerLimit: -50.2m, upperLimit: -49.8m),
        new("調整作業", ItemKind.Task),
    ];

    private static List<TemplateRow> CreateRows() => [
        new(Row: 5, Name: "電源投入", LowerLimit: null, UpperLimit: null, IsConfirmation: true),
        new(Row: 8, Name: "消費電流", LowerLimit: 0m, UpperLimit: 230m, IsConfirmation: false),
        new(Row: 16, Name: "温度入力", LowerLimit: -50.2m, UpperLimit: -49.8m, IsConfirmation: false),
    ];

    [Fact]
    public void MatchingDefinitionAndTemplateHaveNoMismatch() {
        var mismatches = RangeComparer.Compare(CreateItems(), CreateRows());

        Assert.Empty(mismatches);
    }

    [Fact]
    public void TaskItemsAreNotCompared() {
        // 作業項目は Excel の行を持たないので、テンプレートに対応する行がなくても食い違いにしない
        var mismatches = RangeComparer.Compare(CreateItems(), CreateRows());

        Assert.DoesNotContain(mismatches, mismatch => mismatch.ItemName == "調整作業");
    }

    [Theory]
    [InlineData("0.0001", "230")]
    [InlineData("0", "229.9999")]
    public void SmallDifferenceInLimitIsDetected(string lower, string upper) {
        var rows = CreateRows();
        rows[1] = rows[1] with { LowerLimit = decimal.Parse(lower), UpperLimit = decimal.Parse(upper) };

        var mismatch = Assert.Single(RangeComparer.Compare(CreateItems(), rows));

        Assert.Equal(8, mismatch.Row);
        Assert.Equal("消費電流", mismatch.ItemName);
        Assert.Contains("合格範囲が違います", mismatch.Reason);
    }

    [Fact]
    public void ConfirmationItemAgainstNumericRangeIsDetected() {
        var rows = CreateRows();
        rows[0] = rows[0] with { IsConfirmation = false, LowerLimit = 0m, UpperLimit = 1m };

        var mismatch = Assert.Single(RangeComparer.Compare(CreateItems(), rows));

        Assert.Equal(5, mismatch.Row);
        Assert.Contains("確認項目", mismatch.Reason);
    }

    [Fact]
    public void ValueItemAgainstConfirmationRowIsDetected() {
        var rows = CreateRows();
        rows[1] = rows[1] with { IsConfirmation = true, LowerLimit = null, UpperLimit = null };

        var mismatch = Assert.Single(RangeComparer.Compare(CreateItems(), rows));

        Assert.Equal(8, mismatch.Row);
        Assert.Contains("確認項目", mismatch.Reason);
    }

    [Fact]
    public void ItemRowMissingInTemplateIsDetected() {
        var rows = CreateRows();
        rows.RemoveAt(2);

        var mismatch = Assert.Single(RangeComparer.Compare(CreateItems(), rows));

        Assert.Equal(16, mismatch.Row);
        Assert.Contains("テンプレートにこの行", mismatch.Reason);
    }

    [Fact]
    public void TemplateRowMissingInDefinitionIsDetected() {
        var items = CreateItems();
        items.RemoveAt(1);

        var mismatch = Assert.Single(RangeComparer.Compare(items, CreateRows()));

        Assert.Equal(8, mismatch.Row);
        Assert.Equal("消費電流", mismatch.ItemName);
        Assert.Contains("定義の項目にありません", mismatch.Reason);
    }

    [Fact]
    public void MismatchesAreOrderedByRow() {
        var items = CreateItems();
        items.RemoveAt(2);
        var rows = CreateRows();
        rows[0] = rows[0] with { IsConfirmation = false, LowerLimit = 0m, UpperLimit = 1m };

        var mismatches = RangeComparer.Compare(items, rows);

        Assert.Equal([5, 16], mismatches.Select(mismatch => mismatch.Row));
    }
}
