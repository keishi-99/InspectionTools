using System.Text;
using InspectionTools.Excel;
using InspectionTools.Inspection;

// Excel のテンプレートから、製品の定義（JSON）のひな形を作る。作ったあとは JSON を直接編集して仕上げる（作業項目の追加など）
Console.OutputEncoding = Encoding.UTF8;

const string Usage = """
    使い方: DefinitionTool <テンプレートのパス> <製品名> <出力先のJSON> [オプション]
    オプション:
      --measurement-rows <行の指定>   測定項目にする行（例: 8-15,30-38）。指定しない行は入力項目になる
      --data-sheet <シート名>        項目名のシート（既定: ACdata）
      --setting-sheet <シート名>     合格範囲のシート（既定: ACsetting）
      --rows <行の指定>              読み取る行の範囲（既定: 5-50）
      --max-units <台数>             最大台数（既定: 30）
      --verify-cell <セル>           製品の確認に使うセル（既定: A3）
      --verify-value <値>            確認用のセルに入っているはずの値（必須。例: EL5000）
      --first-unit-column <列>       1台目を書き込む列（既定: F）
      --units-per-block <台数>       折り返すまでの台数（折り返す製品だけ。--block-row-offset と一緒に指定）
      --block-row-offset <行数>      折り返したときに下へ進める行数
    """;

if (args.Length < 3) {
    Console.WriteLine(Usage);
    return 1;
}

try {
    var templatePath = args[0];
    var product = args[1];
    var outputPath = args[2];
    var options = ParseOptions(args[3..]);

    var rowRange = DefinitionDraftBuilder.ParseRows(options.GetValueOrDefault("--rows", "5-50")).Order().ToList();
    var layout = new TemplateLayout(
        options.GetValueOrDefault("--data-sheet", "ACdata"),
        options.GetValueOrDefault("--setting-sheet", "ACsetting"),
        rowRange.First(),
        rowRange.Last());
    var measurementRows = DefinitionDraftBuilder.ParseRows(options.GetValueOrDefault("--measurement-rows", string.Empty));

    var maxUnitCount = int.Parse(options.GetValueOrDefault("--max-units", "30"));

    var transfer = new TransferDefinition(
        layout.DataSheet,
        layout.SettingSheet,
        layout.FirstRow,
        layout.LastRow,
        options.GetValueOrDefault("--verify-cell", "A3"),
        options.TryGetValue("--verify-value", out var verifyValue) ? verifyValue : throw new ArgumentException("--verify-value を指定してください。"),
        options.GetValueOrDefault("--first-unit-column", "F"),
        options.TryGetValue("--units-per-block", out var unitsPerBlock) ? int.Parse(unitsPerBlock) : null,
        options.TryGetValue("--block-row-offset", out var blockRowOffset) ? int.Parse(blockRowOffset) : 0);

    var rows = TemplateReader.ReadRows(templatePath, layout);
    var json = DefinitionDraftBuilder.BuildJson(product, rows, measurementRows, maxUnitCount, transfer);

    // 書き出す前に、アプリが読み込める定義になっていることを確認する
    var definition = ProductDefinitionLoader.Parse(json);

    Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);
    File.WriteAllText(outputPath, json, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

    Console.WriteLine($"{definition.Items.Count} 項目を書き出しました: {outputPath}");
    foreach (var kind in Enum.GetValues<ItemKind>()) {
        Console.WriteLine($"  {kind}: {definition.Items.Count(item => item.Kind == kind)}");
    }
    return 0;
} catch (Exception ex) when (ex is TemplateReadException or ProductDefinitionException or ArgumentException or FormatException) {
    Console.Error.WriteLine($"エラー: {ex.Message}");
    return 1;
}

// 「--名前 値」の形のオプションを、名前と値の組にする
static Dictionary<string, string> ParseOptions(string[] optionArgs) {
    var options = new Dictionary<string, string>();
    for (var i = 0; i < optionArgs.Length; i += 2) {
        if (!optionArgs[i].StartsWith("--") || i + 1 >= optionArgs.Length) {
            throw new ArgumentException($"オプションの指定が正しくありません: {optionArgs[i]}");
        }
        options[optionArgs[i]] = optionArgs[i + 1];
    }
    return options;
}
