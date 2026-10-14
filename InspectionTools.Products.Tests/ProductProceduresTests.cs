namespace InspectionTools.Products.Tests;

/// <summary>製品名から手順表を探す動きを確認する</summary>
public sealed class ProductProceduresTests {

    [Fact]
    public void FindsTheProcedureOfEl5000Ac() {
        var procedure = ProductProcedures.TryCreate("EL5000AC");

        Assert.Equal("EL5000AC", procedure?.Product);
    }

    [Fact]
    public void ReturnsNullForAProductWithoutAProcedure() {
        Assert.Null(ProductProcedures.TryCreate("NoSuchProduct"));
    }
}
