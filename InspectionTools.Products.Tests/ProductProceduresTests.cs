namespace InspectionTools.Products.Tests;

/// <summary>製品名から手順表を探す動きを確認する</summary>
public sealed class ProductProceduresTests {

    [Fact]
    public void FindsTheProcedureOfEl5000Ac() {
        var procedure = ProductProcedures.TryCreate("EL5000AC");

        Assert.Equal("EL5000AC", procedure?.Product);
    }

    // PAF5 は、タイプごとに定義（製品名）が分かれるが、測定器の手順表は共通。製品名が「PAF5-」で始まる製品は、同じ手順表を使う
    [Theory]
    [InlineData("PAF5-分離形")]
    [InlineData("PAF5-一体形2線式")]
    [InlineData("PAF5-一体形3線式")]
    public void FindsTheSharedProcedureOfEveryPaf5Type(string product) {
        var procedure = ProductProcedures.TryCreate(product);

        Assert.Equal(product, procedure?.Product);
    }

    [Theory]
    [InlineData("PAF5")]
    [InlineData("PAF5amp")]
    public void ProductsWhoseNamesDoNotStartWithPaf5HyphenHaveNoPaf5Procedure(string product) {
        Assert.Null(ProductProcedures.TryCreate(product));
    }

    [Fact]
    public void ReturnsNullForAProductWithoutAProcedure() {
        Assert.Null(ProductProcedures.TryCreate("NoSuchProduct"));
    }
}
