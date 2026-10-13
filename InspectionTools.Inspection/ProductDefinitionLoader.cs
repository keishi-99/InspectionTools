using System.Text.Json;
using System.Text.Json.Serialization;

namespace InspectionTools.Inspection;

/// <summary>製品ごとの定義（JSON）を読み込み、検査項目の一覧にする</summary>
public static class ProductDefinitionLoader {

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web) {
        Converters = { new JsonStringEnumConverter() },
    };

    // JSON の形（欠けている項目を検査で見つけられるように、すべて null を許す）
    private sealed record ItemFile(string? Name, ItemKind? Kind, int? ExcelRow, decimal? LowerLimit, decimal? UpperLimit);

    private sealed record DefinitionFile(string? Product, int? MaxUnitCount, List<ItemFile>? Items);

    /// <summary>ファイルから定義を読み込む</summary>
    public static ProductDefinition LoadFile(string path) => Parse(File.ReadAllText(path));

    /// <summary>JSON の文字列から定義を読み込む。正しくないときは、どの項目の何が悪いかを含む ProductDefinitionException</summary>
    public static ProductDefinition Parse(string json) {
        DefinitionFile file;
        try {
            file = JsonSerializer.Deserialize<DefinitionFile>(json, JsonOptions)
                ?? throw new ProductDefinitionException("定義の内容が空です。");
        } catch (JsonException ex) {
            throw new ProductDefinitionException($"定義の JSON が正しくありません: {ex.Message}", ex);
        }

        if (string.IsNullOrWhiteSpace(file.Product)) {
            throw new ProductDefinitionException("製品名（product）がありません。");
        }
        if (file.MaxUnitCount is null or < 1) {
            throw new ProductDefinitionException("最大台数（maxUnitCount）は1以上の整数で指定してください。");
        }
        if (file.Items is null || file.Items.Count == 0) {
            throw new ProductDefinitionException("検査項目（items）が1つもありません。");
        }

        var items = file.Items.Select((item, index) => CreateItem(item, index)).ToList();
        ValidateNoDuplicateRows(items);
        return new ProductDefinition(file.Product, file.MaxUnitCount.Value, items);
    }

    // JSON の1項目を InspectionItem にする（項目の決まりに反するときは、何番目の項目かを添えて知らせる）
    private static InspectionItem CreateItem(ItemFile item, int index) {
        var position = $"{index + 1}番目の項目（{item.Name}）";
        if (item.Kind is null) {
            throw new ProductDefinitionException($"{position}: 種類（kind）がありません。");
        }

        try {
            return new InspectionItem(item.Name ?? string.Empty, item.Kind.Value, item.ExcelRow, item.LowerLimit, item.UpperLimit);
        } catch (ArgumentException ex) {
            throw new ProductDefinitionException($"{position}: {ex.Message}", ex);
        }
    }

    // 同じ Excel の行に書き込む項目が2つあると、片方の結果で上書きされてしまうため認めない
    private static void ValidateNoDuplicateRows(List<InspectionItem> items) {
        var duplicate = items
            .Where(item => item.ExcelRow is not null)
            .GroupBy(item => item.ExcelRow)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null) {
            var names = string.Join("、", duplicate.Select(item => item.Name));
            throw new ProductDefinitionException($"Excel の行 {duplicate.Key} が重複しています: {names}");
        }
    }
}
