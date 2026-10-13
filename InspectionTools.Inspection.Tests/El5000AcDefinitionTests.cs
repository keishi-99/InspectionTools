namespace InspectionTools.Inspection.Tests;

/// <summary>リポジトリにある EL5000AC の定義（products/EL5000AC.json）が、テンプレートと要件定義書の内容と合っていることを確認する</summary>
public class El5000AcDefinitionTests {

    private static readonly ProductDefinition Definition = ProductDefinitionLoader.LoadFile(FindDefinitionPath());

    // テストの実行場所から上の階層をたどって、リポジトリの products フォルダにある定義のパスを探す
    private static string FindDefinitionPath() {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent) {
            if (File.Exists(Path.Combine(directory.FullName, "InspectionTools.sln"))) {
                return Path.Combine(directory.FullName, "products", "EL5000AC.json");
            }
        }
        throw new FileNotFoundException("リポジトリのルート（InspectionTools.sln）が見つかりません。");
    }

    [Fact]
    public void DefinitionHasTheExpectedProductAndItemCount() {
        Assert.Equal("EL5000AC", Definition.Product);
        Assert.Equal(46 + 8, Definition.Items.Count);
    }

    [Fact]
    public void DefinitionHasTheExpectedKindCounts() {
        int Count(ItemKind kind) => Definition.Items.Count(item => item.Kind == kind);

        Assert.Equal(17, Count(ItemKind.Measurement));
        Assert.Equal(22, Count(ItemKind.Input));
        Assert.Equal(7, Count(ItemKind.Confirmation));
        Assert.Equal(8, Count(ItemKind.Task));
    }

    [Fact]
    public void RecordedItemsCoverExcelRows5To50InOrder() {
        var rows = Definition.Items.Where(item => item.ExcelRow is not null).Select(item => item.ExcelRow!.Value);

        Assert.Equal(Enumerable.Range(5, 46), rows);
    }

    [Fact]
    public void TaskItemsAreBetweenRow17And18() {
        var firstTask = Definition.Items.ToList().FindIndex(item => item.Kind == ItemKind.Task);
        var lastTask = Definition.Items.ToList().FindLastIndex(item => item.Kind == ItemKind.Task);

        Assert.Equal(17, Definition.Items[firstTask - 1].ExcelRow);
        Assert.Equal(18, Definition.Items[lastTask + 1].ExcelRow);
        Assert.Equal(8, lastTask - firstTask + 1);
    }

    [Fact]
    public void ItemNamesAreUnique() {
        var duplicate = Definition.Items.GroupBy(item => item.Name).FirstOrDefault(group => group.Count() > 1);

        Assert.Null(duplicate);
    }

    [Fact]
    public void RangesMatchTheTemplateForSampleRows() {
        var current = Definition.Items.Single(item => item.ExcelRow == 8);
        var pulseInput2 = Definition.Items.Single(item => item.ExcelRow == 47);

        Assert.Equal((0m, 230m), (current.LowerLimit, current.UpperLimit));
        Assert.Equal((6999m, 7001m), (pulseInput2.LowerLimit, pulseInput2.UpperLimit));
    }
}
