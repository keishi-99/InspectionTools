using InspectionTools.Inspection;

namespace InspectionTools.Excel;

/// <summary>定義（JSON）の合格範囲と、テンプレートの合格範囲を突き合わせる（画面の判定と Excel の判定がずれるのを防ぐため）</summary>
public static class RangeComparer {

    /// <summary>食い違いを行の順に返す。1件もなければ、定義とテンプレートが一致している</summary>
    public static IReadOnlyList<RangeMismatch> Compare(IReadOnlyList<InspectionItem> items, IReadOnlyList<TemplateRow> rows) {
        var mismatches = new List<RangeMismatch>();
        var rowsByNumber = rows.ToDictionary(row => row.Row);

        // 作業項目は Excel の行を持たないので、突き合わせの対象にしない
        var recordedItems = items.Where(item => item.ExcelRow is not null).ToList();
        foreach (var item in recordedItems) {
            var excelRow = item.ExcelRow!.Value;
            if (!rowsByNumber.TryGetValue(excelRow, out var row)) {
                mismatches.Add(new RangeMismatch(excelRow, item.Name, "テンプレートにこの行の合格範囲がありません。"));
                continue;
            }

            var reason = FindDifference(item, row);
            if (reason is not null) {
                mismatches.Add(new RangeMismatch(excelRow, item.Name, reason));
            }
        }

        // テンプレートにあるのに定義にない行は、結果が書き込まれないまま判定だけが残るため知らせる
        var definedRows = recordedItems.Select(item => item.ExcelRow!.Value).ToHashSet();
        foreach (var row in rows.Where(row => !definedRows.Contains(row.Row))) {
            mismatches.Add(new RangeMismatch(row.Row, row.Name, "テンプレートにある行ですが、定義の項目にありません。"));
        }

        return mismatches.OrderBy(mismatch => mismatch.Row).ToList();
    }

    // 項目とテンプレートの行を比べて、違いがあれば理由を返す（なければ null）
    private static string? FindDifference(InspectionItem item, TemplateRow row) {
        var isConfirmation = item.Kind == ItemKind.Confirmation;
        if (isConfirmation != row.IsConfirmation) {
            return isConfirmation
                ? "定義では確認項目ですが、テンプレートの合格範囲は数値です。"
                : "定義では値の項目ですが、テンプレートの合格範囲は文字（確認項目）です。";
        }

        if (!isConfirmation && (item.LowerLimit != row.LowerLimit || item.UpperLimit != row.UpperLimit)) {
            return $"合格範囲が違います: 定義 {item.LowerLimit}〜{item.UpperLimit} / テンプレート {row.LowerLimit}〜{row.UpperLimit}";
        }

        return null;
    }
}
