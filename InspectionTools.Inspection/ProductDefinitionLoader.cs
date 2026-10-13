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

    private sealed record TransferFile(
        string? DataSheet,
        string? SettingSheet,
        int? FirstRow,
        int? LastRow,
        string? VerifyCell,
        string? VerifyValue,
        string? FirstUnitColumn,
        int? UnitsPerBlock,
        int? BlockRowOffset);

    private sealed record DefinitionFile(string? Product, int? MaxUnitCount, TransferFile? Transfer, List<ItemFile>? Items);

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

        var transfer = CreateTransfer(file.Transfer);
        var items = file.Items.Select((item, index) => CreateItem(item, index)).ToList();
        ValidateNoDuplicateRows(items);
        return new ProductDefinition(file.Product, file.MaxUnitCount.Value, transfer, items);
    }

    // JSON の転記の設定を TransferDefinition にする（欠けている項目や、使えない値は、どの項目かを添えて知らせる）
    private static TransferDefinition CreateTransfer(TransferFile? file) {
        if (file is null) {
            throw new ProductDefinitionException("転記の設定（transfer）がありません。");
        }

        var dataSheet = RequireText(file.DataSheet, nameof(file.DataSheet));
        var settingSheet = RequireText(file.SettingSheet, nameof(file.SettingSheet));
        var verifyCell = RequireText(file.VerifyCell, nameof(file.VerifyCell));
        var verifyValue = RequireText(file.VerifyValue, nameof(file.VerifyValue));
        var firstUnitColumn = RequireText(file.FirstUnitColumn, nameof(file.FirstUnitColumn));
        if (!firstUnitColumn.All(char.IsAsciiLetter)) {
            throw new ProductDefinitionException($"transfer.firstUnitColumn は列の文字（例: F, AI）で指定してください: {firstUnitColumn}");
        }

        if (file.FirstRow is null or < 1) {
            throw new ProductDefinitionException("transfer.firstRow は1以上の整数で指定してください。");
        }
        if (file.LastRow is null || file.LastRow < file.FirstRow) {
            throw new ProductDefinitionException("transfer.lastRow は firstRow 以上の整数で指定してください。");
        }

        // 折り返しは、台数と行数の両方がそろったときだけ使える
        if (file.UnitsPerBlock is null != file.BlockRowOffset is null) {
            throw new ProductDefinitionException("transfer.unitsPerBlock と transfer.blockRowOffset は、両方を指定するか、両方とも指定しないでください。");
        }
        if (file.UnitsPerBlock is < 1 || file.BlockRowOffset is < 1) {
            throw new ProductDefinitionException("transfer.unitsPerBlock と transfer.blockRowOffset は1以上の整数で指定してください。");
        }

        return new TransferDefinition(
            dataSheet,
            settingSheet,
            file.FirstRow.Value,
            file.LastRow.Value,
            verifyCell,
            verifyValue,
            firstUnitColumn,
            file.UnitsPerBlock,
            file.BlockRowOffset ?? 0);
    }

    // 転記の設定の文字列が空でないことを確認する
    private static string RequireText(string? value, string name) {
        if (string.IsNullOrWhiteSpace(value)) {
            var jsonName = char.ToLowerInvariant(name[0]) + name[1..];
            throw new ProductDefinitionException($"transfer.{jsonName} がありません。");
        }
        return value;
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
