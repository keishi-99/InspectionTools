using ClosedXML.Excel;
using InspectionTools.Inspection;

namespace InspectionTools.Excel;

/// <summary>検査結果を、Excel のテンプレートへ書き込んで保存する</summary>
public static class TemplateWriter {

    /// <summary>
    /// 検査結果をテンプレートへ書き込み、同じ名前で上書き保存する。
    /// 検査した台のセルだけを書き換え（値がなければ空欄にする）、それ以外の台の列や、判定の数式には触れない。
    /// 正しい製品のテンプレートでなければ TemplateVerificationException、他のプログラムで開かれているなどで書き込めなければ TemplateWriteException
    /// </summary>
    public static void Write(string path, TransferSettings settings, InspectionSession session) {
        if (settings.Check is { } check) {
            TemplateVerifier.Verify(path, check);
        }

        // 書き込み中に失敗しても、元のファイルが壊れないように、別名のファイルに最後まで保存してから置き換える
        var temporaryPath = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(path))!, $"{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
        try {
            // 他のプログラム（Excel など）がファイルを開いているときは、ここで失敗させる
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                using var workbook = new XLWorkbook(stream);
                var sheet = TemplateFile.FindSheet(workbook, settings.DataSheet)
                    ?? throw new TemplateReadException($"シートがありません: '{settings.DataSheet}'");

                WriteResults(sheet, settings, session);

                // ClosedXML は保存し直すと、シートの表示形式（改ページプレビュー）を通常表示に戻してしまうため、すべてのシートをそろえる
                foreach (var worksheet in workbook.Worksheets) {
                    worksheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
                }

                // 書き込んだ値で判定の数式の結果が変わるので、Excel で開いたときに1回だけ全部計算し直させる
                // （ForceFullCalculation は計算のたびに全部やり直す別の設定なので使わない）
                workbook.FullCalculationOnLoad = true;
                workbook.SaveAs(temporaryPath);
            }
            File.Move(temporaryPath, path, overwrite: true);
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            throw new TemplateWriteException(
                "テンプレートに書き込めません。Excel などで開いている場合は閉じてから、もう一度実行してください。" +
                $"（{ex.Message}）", ex);
        } finally {
            if (File.Exists(temporaryPath)) {
                File.Delete(temporaryPath);
            }
        }
    }

    // 検査した台ごとに、Excel の行を持つ項目の結果を書き込む（作業項目は行を持たないので書かない）
    private static void WriteResults(IXLWorksheet sheet, TransferSettings settings, InspectionSession session) {
        for (var unitIndex = 0; unitIndex < session.UnitCount; unitIndex++) {
            for (var itemIndex = 0; itemIndex < session.Items.Count; itemIndex++) {
                var item = session.Items[itemIndex];
                if (item.ExcelRow is not { } excelRow) {
                    continue;
                }

                var result = session.GetResult(itemIndex, unitIndex);
                if (item.Kind == ItemKind.Confirmation && settings.UnitLayout.ConfirmationColumn is not null) {
                    // チェックボックスのリンク先のセルには、合格だけ TRUE、不合格と未実施は FALSE を書く（テンプレートの集計が FALSE を数えるため）
                    var linked = settings.UnitLayout.GetConfirmationCell(excelRow, unitIndex);
                    sheet.Cell(linked.Row, linked.Column).Value = result.State == ResultState.Ok;
                    continue;
                }

                var position = settings.UnitLayout.GetCell(excelRow, unitIndex);
                var cell = sheet.Cell(position.Row, position.Column);
                ApplyResult(cell, item, result, settings.ConfirmationOkText);
            }
        }
    }

    // 1つのセルに結果を反映する（書式は変えず、中身だけを書き換える）
    internal static void ApplyResult(IXLCell cell, InspectionItem item, ItemResult result, string confirmationOkText) {
        switch (item.Kind) {
            case ItemKind.Measurement:
            case ItemKind.Input:
                if (result.Value is { } value) {
                    cell.Value = value;
                }
                else {
                    cell.Clear(XLClearOptions.Contents);
                }
                break;
            case ItemKind.Text:
                // 文字のまま書く（数値に変換されると、先頭の 0 や英字が消えるため）
                if (result.Text is { } text) {
                    cell.Value = text;
                }
                else {
                    cell.Clear(XLClearOptions.Contents);
                }
                break;
            case ItemKind.Confirmation:
                if (result.State == ResultState.Ok) {
                    cell.Value = confirmationOkText;
                }
                else {
                    cell.Clear(XLClearOptions.Contents);
                }
                break;
            default:
                break;
        }
    }
}
