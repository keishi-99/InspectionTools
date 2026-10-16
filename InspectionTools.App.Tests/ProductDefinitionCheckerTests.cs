using InspectionTools.Inspection;
using InspectionTools.Products;

namespace InspectionTools.App.Tests;

/// <summary>製品の定義と手順表の食い違いを、製品を開くときに見つけることを確認する</summary>
public sealed class ProductDefinitionCheckerTests {

    private static readonly TransferDefinition Transfer = new("ACdata", "ACsetting", 5, 50, "A3", "EL5000", "F", UnitsPerBlock: null, BlockRowOffset: 0);

    private static ProductDefinition DefinitionWith(params InspectionItem[] items) => new("EL5000AC", 30, Transfer, items);

    [Fact]
    public void ShippedEl5000AcDefinitionMatchesItsProcedure() {
        var path = RepositoryFiles.Find("products", "EL5000AC.json");
        var definition = ProductDefinitionLoader.LoadFile(path);

        ProductDefinitionChecker.EnsureActionsExist(definition, El5000AcProcedure.Create());
    }

    [Fact]
    public void ItemsWithoutActionIdAreAllowed() {
        var definition = DefinitionWith(new InspectionItem("確認", ItemKind.Confirmation, excelRow: 5));

        ProductDefinitionChecker.EnsureActionsExist(definition, El5000AcProcedure.Create());
    }

    [Fact]
    public void UnknownActionIdIsReportedWithTheItemName() {
        var definition = DefinitionWith(
            new InspectionItem("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m, actionId: "dmm01-aci"),
            new InspectionItem("調整", ItemKind.Task, actionId: "dcs-99mA"));

        var exception = Assert.Throws<ProductDefinitionException>(() => ProductDefinitionChecker.EnsureActionsExist(definition, El5000AcProcedure.Create()));

        Assert.Contains("調整(actionId: dcs-99mA)", exception.Message);
        Assert.DoesNotContain("消費電流", exception.Message);
    }

    [Fact]
    public void EveryUnknownActionIdIsListed() {
        var definition = DefinitionWith(
            new InspectionItem("調整1", ItemKind.Task, actionId: "xxx"),
            new InspectionItem("調整2", ItemKind.Task, actionId: "yyy"));

        var exception = Assert.Throws<ProductDefinitionException>(() => ProductDefinitionChecker.EnsureActionsExist(definition, El5000AcProcedure.Create()));

        Assert.Contains("xxx", exception.Message);
        Assert.Contains("yyy", exception.Message);
    }
}
