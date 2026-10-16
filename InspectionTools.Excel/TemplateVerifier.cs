using ClosedXML.Excel;

namespace InspectionTools.Excel;

/// <summary>選んだファイルが正しい製品のテンプレートかを、書き込む前に確かめる</summary>
public static class TemplateVerifier {

    /// <summary>
    /// 目印のシートがあり、目印のセルに期待する値が含まれているかを確かめる（製品名の前後に説明が付くテンプレートでも通すため、完全一致ではなく含むかで見る）。
    /// 違うときは TemplateVerificationException、ファイルが読めないときは TemplateReadException
    /// </summary>
    public static void Verify(string path, TemplateCheck check) {
        using var workbook = TemplateFile.OpenForRead(path);
        Verify(workbook, check);
    }

    /// <summary>開いてあるブックで確かめる(複数のシートを確かめるときに、ファイルを何度も開かないため)</summary>
    internal static void Verify(XLWorkbook workbook, TemplateCheck check) {
        var sheet = TemplateFile.FindSheet(workbook, check.Sheet)
            ?? throw new TemplateVerificationException($"このファイルには、シート '{check.Sheet}' がありません。別の製品のテンプレートではありませんか。");

        var actual = sheet.Cell(check.Cell).GetString().Trim();
        if (!actual.Contains(check.ExpectedValue, StringComparison.Ordinal)) {
            var shown = actual.Length == 0 ? "（空欄）" : $"'{actual}'";
            throw new TemplateVerificationException(
                $"このファイルは、対象の製品のテンプレートではありません。{check.Sheet}!{check.Cell} が '{check.ExpectedValue}' のはずですが、{shown} でした。");
        }
    }
}
