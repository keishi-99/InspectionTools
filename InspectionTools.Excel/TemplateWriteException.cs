namespace InspectionTools.Excel;

/// <summary>テンプレートへ書き込めないとき（他のプログラムで開かれている、書き込みが許されていない、など）の例外</summary>
public sealed class TemplateWriteException(string message, Exception? innerException = null) : Exception(message, innerException);
