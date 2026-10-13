using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using InspectionTools.Inspection;

namespace InspectionTools.Excel;

/// <summary>テンプレートから読み取った行を、製品の定義（JSON）のひな形にする</summary>
public static class DefinitionDraftBuilder {

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web) {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        // 日本語の項目名を \uXXXX の形にせず、そのまま出力する（手で直しやすくするため）
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    private sealed record ItemDraft(string Name, ItemKind Kind, int ExcelRow, decimal? LowerLimit, decimal? UpperLimit);

    private sealed record DefinitionDraft(string Product, List<ItemDraft> Items);

    /// <summary>
    /// 製品の定義（JSON）のひな形を作る。種類は、合格範囲が文字の行を確認項目、measurementRows に含まれる行を測定項目、
    /// それ以外を入力項目にする（測定器から取るか、製品の表示を見るかはテンプレートからは判断できないため）
    /// </summary>
    public static string BuildJson(string product, IReadOnlyList<TemplateRow> rows, IReadOnlySet<int> measurementRows) {
        foreach (var measurementRow in measurementRows) {
            var row = rows.FirstOrDefault(r => r.Row == measurementRow)
                ?? throw new ArgumentException($"測定項目に指定された行がテンプレートにありません: {measurementRow}", nameof(measurementRows));
            if (row.IsConfirmation) {
                throw new ArgumentException($"確認項目の行は、測定項目に指定できません: {measurementRow}", nameof(measurementRows));
            }
        }

        var items = rows.Select(row => new ItemDraft(
            row.Name,
            row.IsConfirmation ? ItemKind.Confirmation : measurementRows.Contains(row.Row) ? ItemKind.Measurement : ItemKind.Input,
            row.Row,
            row.LowerLimit,
            row.UpperLimit)).ToList();
        return JsonSerializer.Serialize(new DefinitionDraft(product, items), JsonOptions);
    }

    /// <summary>「8-15,30-38」のような行の指定を、行の集合にする（「30」のように1行だけの指定もできる）</summary>
    public static HashSet<int> ParseRows(string text) {
        var rows = new HashSet<int>();
        foreach (var part in text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) {
            var bounds = part.Split('-', StringSplitOptions.TrimEntries);
            if (!TryParseRange(bounds, out var first, out var last)) {
                throw new FormatException($"行の指定が正しくありません: '{part}'（例: 8-15,30）");
            }
            for (var row = first; row <= last; row++) {
                rows.Add(row);
            }
        }
        return rows;
    }

    // 「30」または「8-15」を、始めの行と終わりの行にする
    private static bool TryParseRange(string[] bounds, out int first, out int last) {
        first = 0;
        last = 0;
        switch (bounds.Length) {
            case 1:
                if (!int.TryParse(bounds[0], out first) || first < 1) {
                    return false;
                }
                last = first;
                return true;
            case 2:
                return int.TryParse(bounds[0], out first) && int.TryParse(bounds[1], out last) && first >= 1 && first <= last;
            default:
                return false;
        }
    }
}
