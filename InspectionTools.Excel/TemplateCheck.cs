namespace InspectionTools.Excel;

/// <summary>選んだファイルが正しい製品のテンプレートかを確かめるための目印（どのシートの、どのセルが、どの値か）</summary>
/// <param name="Sheet">確かめるシート名（末尾の空白も含めて完全に一致させる）</param>
/// <param name="Cell">確かめるセル（例: A3）</param>
/// <param name="ExpectedValue">そのセルに入っているはずの値（例: EL5000）</param>
public sealed record TemplateCheck(string Sheet, string Cell, string ExpectedValue);
