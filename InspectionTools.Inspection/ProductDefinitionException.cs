namespace InspectionTools.Inspection;

/// <summary>製品の定義（JSON）が正しくなく、読み込めないときの例外</summary>
public sealed class ProductDefinitionException(string message, Exception? innerException = null) : Exception(message, innerException);
