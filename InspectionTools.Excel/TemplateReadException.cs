namespace InspectionTools.Excel;

/// <summary>テンプレートのファイルを読み取れない、またはシートがないときの例外</summary>
public sealed class TemplateReadException(string message, Exception? innerException = null) : Exception(message, innerException);
