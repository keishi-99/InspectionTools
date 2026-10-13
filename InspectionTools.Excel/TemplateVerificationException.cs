namespace InspectionTools.Excel;

/// <summary>選んだファイルが、正しい製品のテンプレートではないときの例外（書き込む前に気づくために使う）</summary>
public sealed class TemplateVerificationException(string message) : Exception(message);
