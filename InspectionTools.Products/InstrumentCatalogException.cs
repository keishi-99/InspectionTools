namespace InspectionTools.Products;

/// <summary>測定器の一覧（VisaAddress.xml）を読み込めないとき（ファイルがない、XML として正しくない）の例外</summary>
public sealed class InstrumentCatalogException(string message, Exception? innerException = null) : Exception(message, innerException);
