using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace InspectionTools.Inspection;

/// <summary>製品ごとの定義（JSON）を読み込み、検査項目の一覧にする</summary>
public static partial class ProductDefinitionLoader {

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web) {
        Converters = { new JsonStringEnumConverter() },
    };

    // JSON の形（欠けている項目を検査で見つけられるように、すべて null を許す）
    private sealed record ItemFile(string? Name, ItemKind? Kind, int? ExcelRow, decimal? LowerLimit, decimal? UpperLimit, string? ActionId, string? Description, decimal? DefaultValue);

    private sealed record TransferFile(
        string? DataSheet,
        string? SettingSheet,
        int? FirstRow,
        int? LastRow,
        string? VerifyCell,
        string? VerifyValue,
        string? FirstUnitColumn,
        int? UnitsPerBlock,
        int? BlockRowOffset,
        string? ConfirmationColumn,
        List<string>? DataSheets,
        int? UnitsPerSheet,
        string? UnitCountCell);

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

        var transfer = CreateTransfer(file.Transfer, file.MaxUnitCount.Value);
        var items = file.Items.Select((item, index) => CreateItem(item, index)).ToList();
        ValidateNoDuplicateRows(items);
        return new ProductDefinition(file.Product, file.MaxUnitCount.Value, transfer, items);
    }

    // JSON の転記の設定を TransferDefinition にする（欠けている項目や、使えない値は、どの項目かを添えて知らせる）
    private static TransferDefinition CreateTransfer(TransferFile? file, int maxUnitCount) {
        if (file is null) {
            throw new ProductDefinitionException("転記の設定（transfer）がありません。");
        }

        var (dataSheet, dataSheets, unitsPerSheet) = CreateDataSheets(file, maxUnitCount);
        var settingSheet = OptionalText(file.SettingSheet, nameof(file.SettingSheet));
        var firstUnitColumn = RequireText(file.FirstUnitColumn, nameof(file.FirstUnitColumn));
        if (!firstUnitColumn.All(char.IsAsciiLetter)) {
            throw new ProductDefinitionException($"transfer.firstUnitColumn は列の文字（例: F, AI）で指定してください: {firstUnitColumn}");
        }

        // 製品の確認は、セルと値の両方を書くか、両方とも書かない（確認しない）かのどちらか
        var verifyCell = OptionalText(file.VerifyCell, nameof(file.VerifyCell));
        var verifyValue = OptionalText(file.VerifyValue, nameof(file.VerifyValue));
        if (verifyCell is null && verifyValue is not null) {
            throw new ProductDefinitionException("transfer.verifyCell がありません（verifyValue を指定したときは必要です）。");
        }
        if (verifyCell is not null && verifyValue is null) {
            throw new ProductDefinitionException("transfer.verifyValue がありません（verifyCell を指定したときは必要です）。");
        }

        var confirmationColumn = OptionalText(file.ConfirmationColumn, nameof(file.ConfirmationColumn));
        if (confirmationColumn is not null && !confirmationColumn.All(char.IsAsciiLetter)) {
            throw new ProductDefinitionException($"transfer.confirmationColumn は列の文字（例: M, AI）で指定してください: {confirmationColumn}");
        }

        var unitCountCell = OptionalText(file.UnitCountCell, nameof(file.UnitCountCell));
        if (unitCountCell is not null && !CellAddressPattern().IsMatch(unitCountCell)) {
            throw new ProductDefinitionException($"transfer.unitCountCell はセルの位置（例: C1）で指定してください: {unitCountCell}");
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
            file.BlockRowOffset ?? 0,
            confirmationColumn,
            dataSheets,
            unitsPerSheet,
            unitCountCell);
    }

    [GeneratedRegex("^[A-Za-z]+[1-9][0-9]*$")]
    private static partial Regex CellAddressPattern();

    // 結果を書き込むシートを決める。1枚のシートに書く製品は dataSheet、台を数台ずつ別のシートに書く製品は dataSheets と unitsPerSheet を使う
    // （後者の dataSheet は、最初のシートの名前になる）
    private static (string DataSheet, IReadOnlyList<string>? DataSheets, int? UnitsPerSheet) CreateDataSheets(TransferFile file, int maxUnitCount) {
        if (file.DataSheets is null) {
            if (file.UnitsPerSheet is not null) {
                throw new ProductDefinitionException("transfer.unitsPerSheet は、transfer.dataSheets と一緒に指定してください。");
            }
            return (RequireText(file.DataSheet, nameof(file.DataSheet)), null, null);
        }

        if (file.DataSheet is not null) {
            throw new ProductDefinitionException("transfer.dataSheet と transfer.dataSheets は、どちらか一方だけを指定してください。");
        }
        if (file.DataSheets.Count == 0 || file.DataSheets.Any(string.IsNullOrWhiteSpace)) {
            throw new ProductDefinitionException("transfer.dataSheets には、空でないシート名を1つ以上書いてください。");
        }
        var duplicate = file.DataSheets.GroupBy(name => name).FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null) {
            throw new ProductDefinitionException($"transfer.dataSheets に、同じシート名が重複しています: {duplicate.Key}");
        }
        if (file.UnitsPerSheet is null or < 1) {
            throw new ProductDefinitionException("transfer.unitsPerSheet は1以上の整数で指定してください（transfer.dataSheets を指定したときは必要です）。");
        }
        if (file.UnitsPerBlock is not null) {
            throw new ProductDefinitionException("transfer.dataSheets と transfer.unitsPerBlock は、同時に指定できません。");
        }

        var capacity = file.DataSheets.Count * file.UnitsPerSheet.Value;
        if (maxUnitCount > capacity) {
            throw new ProductDefinitionException($"最大台数（maxUnitCount: {maxUnitCount}）が、transfer.dataSheets に書ける台数（{capacity}）を超えています。");
        }
        return (file.DataSheets[0], file.DataSheets, file.UnitsPerSheet);
    }

    // 転記の設定の文字列が空でないことを確認する
    private static string RequireText(string? value, string name) {
        if (string.IsNullOrWhiteSpace(value)) {
            var jsonName = char.ToLowerInvariant(name[0]) + name[1..];
            throw new ProductDefinitionException($"transfer.{jsonName} がありません。");
        }
        return value;
    }

    // 省略できる転記の設定の文字列。書かれていなければ null、書かれているのに空なら、どの項目かを添えて知らせる
    private static string? OptionalText(string? value, string name) {
        if (value is null) {
            return null;
        }
        return RequireText(value, name);
    }

    // JSON の1項目を InspectionItem にする（項目の決まりに反するときは、何番目の項目かを添えて知らせる）
    private static InspectionItem CreateItem(ItemFile item, int index) {
        var position = $"{index + 1}番目の項目（{item.Name}）";
        if (item.Kind is null) {
            throw new ProductDefinitionException($"{position}: 種類（kind）がありません。");
        }

        try {
            return new InspectionItem(item.Name ?? string.Empty, item.Kind.Value, item.ExcelRow, item.LowerLimit, item.UpperLimit, item.ActionId, item.Description, item.DefaultValue);
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
