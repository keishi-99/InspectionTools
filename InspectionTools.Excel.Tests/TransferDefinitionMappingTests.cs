using InspectionTools.Inspection;

namespace InspectionTools.Excel.Tests;

/// <summary>製品の定義（JSON の transfer）から、転記と読み取りの設定が正しく作られることを確認する</summary>
public class TransferDefinitionMappingTests {

    private static readonly TransferDefinition NoWrap = new("ACdata", "ACsetting", 5, 50, "A3", "EL5000", "F", UnitsPerBlock: null, BlockRowOffset: 0);
    private static readonly TransferDefinition Wrapped = NoWrap with { FirstUnitColumn = "C", UnitsPerBlock = 4, BlockRowOffset = 20 };

    [Fact]
    public void SettingsKeepSheetAndProductCheck() {
        var settings = TransferSettings.From(NoWrap);

        Assert.Equal("ACdata", settings.DataSheet);
        Assert.Equal(new TemplateCheck("ACdata", "A3", "EL5000"), settings.Check);
    }

    [Fact]
    public void SettingsWithoutWrapPlaceUnitsInOneRow() {
        var layout = TransferSettings.From(NoWrap).UnitLayout;

        Assert.Equal(new CellPosition(8, 6), layout.GetCell(itemRow: 8, unitIndex: 0));
        Assert.Equal(new CellPosition(8, 35), layout.GetCell(itemRow: 8, unitIndex: 29));
    }

    [Fact]
    public void SettingsWithWrapPlaceUnitsInBlocks() {
        var layout = TransferSettings.From(Wrapped).UnitLayout;

        Assert.Equal(new CellPosition(5, 3), layout.GetCell(itemRow: 5, unitIndex: 0));
        Assert.Equal(new CellPosition(25, 3), layout.GetCell(itemRow: 5, unitIndex: 4));
    }

    [Fact]
    public void TemplateLayoutKeepsSheetsAndRowRange() {
        var layout = TemplateLayout.From(NoWrap);

        Assert.Equal(new TemplateLayout("ACdata", "ACsetting", 5, 50), layout);
    }

    [Fact]
    public void DraftWithWrapCanBeLoadedBack() {
        var json = DefinitionDraftBuilder.BuildJson(
            "EL3801",
            [new TemplateRow(Row: 5, Name: "出力電圧", LowerLimit: 1m, UpperLimit: 2m, IsConfirmation: false)],
            new HashSet<int> { 5 },
            maxUnitCount: 8,
            Wrapped);

        var transfer = ProductDefinitionLoader.Parse(json).Transfer;

        Assert.Equal(Wrapped, transfer);
    }
}
