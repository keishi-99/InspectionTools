using ClosedXML.Excel;

namespace InspectionTools.Excel;

/// <summary>Excel のテンプレートから、検査項目の名前と合格範囲を読み取る</summary>
public static class TemplateReader {

    // 小項目として名前に含めない目印（規格の表示。例: 24.0V±）
    private const string ToleranceMark = "±";

    /// <summary>テンプレートの指定した行の範囲を読み取る。合格範囲のない行（空行）は含めない</summary>
    public static IReadOnlyList<TemplateRow> ReadRows(string path, TemplateLayout layout) {
        if (!File.Exists(path)) {
            throw new TemplateReadException($"テンプレートのファイルがありません: {path}");
        }

        try {
            // Excel で開いたままのファイルも読めるように、他のプログラムが書き込み中でも共有して開く
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);
            var dataSheet = FindSheet(workbook, layout.DataSheet);
            var settingSheet = FindSheet(workbook, layout.SettingSheet);
            return ReadRows(dataSheet, settingSheet, layout);
        } catch (Exception ex) when (ex is IOException or InvalidDataException or FormatException) {
            throw new TemplateReadException($"テンプレートを読み取れません: {ex.Message}", ex);
        }
    }

    // シート名が完全に一致するシートを探す（末尾の空白も区別する）
    private static IXLWorksheet FindSheet(XLWorkbook workbook, string name) {
        return workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == name)
            ?? throw new TemplateReadException($"シートがありません: '{name}'");
    }

    private static List<TemplateRow> ReadRows(IXLWorksheet dataSheet, IXLWorksheet settingSheet, TemplateLayout layout) {
        var rows = new List<TemplateRow>();
        var group = string.Empty;
        for (var row = layout.FirstRow; row <= layout.LastRow; row++) {
            // 大項目（A列）は先頭の行にしか書かれないので、次に出てくるまで引き継ぐ
            var groupCell = Clean(dataSheet.Cell(row, 1).GetString());
            if (groupCell.Length > 0) {
                group = groupCell;
            }

            var lower = settingSheet.Cell(row, 6);
            var upper = settingSheet.Cell(row, 7);
            if (lower.IsEmpty()) {
                continue;
            }

            var isConfirmation = lower.Value.IsText;
            rows.Add(new TemplateRow(
                row,
                CreateName(dataSheet, row, group),
                isConfirmation ? null : ToDecimal(lower),
                isConfirmation ? null : ToDecimal(upper),
                isConfirmation));
        }
        return rows;
    }

    // 大項目と小項目（B〜D列のうち規格の表示を除いたもの）を空白でつないで、項目名にする
    private static string CreateName(IXLWorksheet dataSheet, int row, string group) {
        var parts = new List<string> { group };
        for (var column = 2; column <= 4; column++) {
            var text = Clean(dataSheet.Cell(row, column).GetString());
            if (text.Length > 0 && !text.Contains(ToleranceMark)) {
                parts.Add(text);
            }
        }
        return string.Join(' ', parts.Where(part => part.Length > 0));
    }

    // セルの文字から改行と前後の空白を取り除く（1つのセルの中で改行されている大項目名を1行にするため）
    private static string Clean(string text) => text.Replace("\r", string.Empty).Replace("\n", string.Empty).Trim();

    // 数式の結果の誤差（22.799999999999997 など）を丸めて、桁の少ない decimal にする
    private static decimal? ToDecimal(IXLCell cell) {
        return cell.Value.IsNumber ? Convert.ToDecimal(cell.Value.GetNumber()) : null;
    }
}
