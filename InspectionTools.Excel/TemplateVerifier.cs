namespace InspectionTools.Excel;

/// <summary>選んだファイルが正しい製品のテンプレートかを、書き込む前に確かめる</summary>
public static class TemplateVerifier {

    /// <summary>
    /// 目印のシートがあり、目印のセルが期待する値になっているかを確かめる。
    /// 違うときは TemplateVerificationException、ファイルが読めないときは TemplateReadException
    /// </summary>
    public static void Verify(string path, TemplateCheck check) {
        using var workbook = TemplateFile.OpenForRead(path);
        var sheet = TemplateFile.FindSheet(workbook, check.Sheet)
            ?? throw new TemplateVerificationException($"このファイルには、シート '{check.Sheet}' がありません。別の製品のテンプレートではありませんか。");

        var actual = sheet.Cell(check.Cell).GetString().Trim();
        if (actual != check.ExpectedValue) {
            var shown = actual.Length == 0 ? "（空欄）" : $"'{actual}'";
            throw new TemplateVerificationException(
                $"このファイルは、対象の製品のテンプレートではありません。{check.Sheet}!{check.Cell} が '{check.ExpectedValue}' のはずですが、{shown} でした。");
        }
    }
}
