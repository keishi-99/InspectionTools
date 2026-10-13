namespace InspectionTools.Inspection;

/// <summary>途中保存の内容が壊れている、または項目の定義と合わないため、状態を復元できないときの例外</summary>
public sealed class SessionRestoreException(string message) : Exception(message);
