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
    public void SettingsWithoutVerifyCellHaveNoProductCheck() {
        var settings = TransferSettings.From(NoWrap with { VerifyCell = null, VerifyValue = null });

        Assert.Null(settings.Check);
    }

    [Fact]
    public void SettingsKeepConfirmationColumn() {
        var layout = TransferSettings.From(NoWrap with { ConfirmationColumn = "K" }).UnitLayout;

        Assert.Equal(new CellPosition(8, 11), layout.GetConfirmationCell(itemRow: 8, unitIndex: 0));
    }

    [Fact]
    public void TemplateLayoutWithoutSettingSheetCannotBeMade() {
        var exception = Assert.Throws<InvalidOperationException>(() => TemplateLayout.From(NoWrap with { SettingSheet = null }));

        Assert.Contains("settingSheet", exception.Message);
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

    [Fact]
    public void DraftWithConfirmationColumnAndWithoutSettingSheetCanBeLoadedBack() {
        var definition = Wrapped with { SettingSheet = null, VerifyCell = null, VerifyValue = null, ConfirmationColumn = "H" };
        var json = DefinitionDraftBuilder.BuildJson(
            "PAF5",
            [new TemplateRow(Row: 5, Name: "実装", LowerLimit: null, UpperLimit: null, IsConfirmation: true)],
            new HashSet<int>(),
            maxUnitCount: 5,
            definition);

        var transfer = ProductDefinitionLoader.Parse(json).Transfer;

        Assert.Equal(definition, transfer);
    }
}
