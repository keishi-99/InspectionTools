using System.IO;
using InspectionTools.Inspection;

namespace InspectionTools.App;

/// <summary>メニューに並べる製品（製品名と、その定義ファイル）</summary>
public sealed record ProductMenuItem(string Name, string DefinitionPath);

/// <summary>製品の一覧。読めなかった定義ファイルは、Warnings に理由を残す</summary>
public sealed record ProductCatalogResult(IReadOnlyList<ProductMenuItem> Items, IReadOnlyList<string> Warnings);

/// <summary>products フォルダの定義ファイル（*.json）を調べて、メニューに並べる製品の一覧を作る</summary>
public static class ProductCatalog {

    public static ProductCatalogResult Scan(string directory) {
        var items = new List<ProductMenuItem>();
        var warnings = new List<string>();

        if (!Directory.Exists(directory)) {
            warnings.Add($"製品の定義フォルダがありません: {directory}");
            return new ProductCatalogResult(items, warnings);
        }

        foreach (var path in Directory.GetFiles(directory, "*.json").Order(StringComparer.OrdinalIgnoreCase)) {
            var fileName = Path.GetFileName(path);
            try {
                var name = ProductDefinitionLoader.LoadFile(path).Product;
                if (items.Any(item => item.Name == name)) {
                    warnings.Add($"{fileName}: 製品名「{name}」が、他のファイルと重なっています。");
                    continue;
                }
                items.Add(new ProductMenuItem(name, path));
            } catch (Exception ex) when (ex is ProductDefinitionException or IOException or UnauthorizedAccessException) {
                warnings.Add($"{fileName}: {ex.Message}");
            }
        }

        return new ProductCatalogResult(items.OrderBy(item => item.Name, StringComparer.Ordinal).ToList(), warnings);
    }
}
