namespace InspectionTools.Products;

/// <summary>製品名から、その製品の測定器の手順表を探す。手順表のない製品は、測定器を使わずに手入力だけで検査する</summary>
public static class ProductProcedures {

    /// <summary>手順表を作る。その製品の手順表がなければ null</summary>
    public static ProductProcedure? TryCreate(string product) =>
        product == El5000AcProcedure.Product ? El5000AcProcedure.Create() : null;
}
