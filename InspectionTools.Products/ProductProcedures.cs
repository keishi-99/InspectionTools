namespace InspectionTools.Products;

/// <summary>製品名から、その製品の測定器の手順表を探す。手順表のない製品は、測定器を使わずに手入力だけで検査する</summary>
public static class ProductProcedures {

    /// <summary>手順表を作る。その製品の手順表がなければ null</summary>
    public static ProductProcedure? TryCreate(string product) {
        if (product == El5000AcProcedure.Product) {
            return El5000AcProcedure.Create();
        }
        // PAF5 は、タイプ(分離形、一体形2線式など)ごとに製品名が分かれるが、測定器の手順表は共通
        if (product.StartsWith(Paf5Procedure.ProductPrefix, StringComparison.Ordinal)) {
            return Paf5Procedure.Create(product);
        }
        return null;
    }
}
