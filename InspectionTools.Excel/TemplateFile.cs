using ClosedXML.Excel;

namespace InspectionTools.Excel;

/// <summary>Excel のテンプレートのファイルを開く共通の処理</summary>
internal static class TemplateFile {

    /// <summary>読み取り用に開く。ファイルがない、または Excel として読めないときは TemplateReadException</summary>
    public static XLWorkbook OpenForRead(string path) {
        if (!File.Exists(path)) {
            throw new TemplateReadException($"テンプレートのファイルがありません: {path}");
        }

        try {
            // Excel で開いたままのファイルも読めるように、他のプログラムが書き込み中でも共有して開く
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return new XLWorkbook(stream);
        } catch (Exception ex) when (ex is IOException or InvalidDataException or FormatException) {
            throw new TemplateReadException($"テンプレートを読み取れません: {ex.Message}", ex);
        }
    }

    /// <summary>シート名が完全に一致するシートを探す（末尾の空白も区別する）。なければ null</summary>
    public static IXLWorksheet? FindSheet(XLWorkbook workbook, string name) {
        return workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == name);
    }
}
