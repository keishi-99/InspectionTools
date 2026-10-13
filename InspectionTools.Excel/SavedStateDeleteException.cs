namespace InspectionTools.Excel;

/// <summary>転記は終わったが、そのあとの途中保存の削除に失敗したときの例外（転記をやり直す必要はない）</summary>
public sealed class SavedStateDeleteException(string message, Exception? innerException = null) : Exception(message, innerException);
