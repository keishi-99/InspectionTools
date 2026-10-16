using InspectionTools.App.ViewModels;
using InspectionTools.Inspection;

namespace InspectionTools.App.Tests;

/// <summary>製品の一覧の作成と、メニューと検査画面の切り替えを、画面なしで確認する</summary>
public sealed class ShellViewModelTests : IDisposable {

    private const string ProductJson = """
        {
          "product": "{NAME}",
          "maxUnitCount": 5,
          "transfer": { "dataSheet": "d", "settingSheet": "s", "firstRow": 5, "lastRow": 10, "verifyCell": "A3", "verifyValue": "X", "firstUnitColumn": "F" },
          "items": [ { "name": "外観確認", "kind": "Confirmation", "excelRow": 5 } ]
        }
        """;

    private sealed class FakeDialogService : IDialogService {
        public List<string> Errors { get; } = [];
        public List<string> Infos { get; } = [];

        public bool Confirm(string message) => false;

        public void ShowError(string message) => Errors.Add(message);

        public List<string> Warnings { get; } = [];

        public void ShowWarning(string message) => Warnings.Add(message);

        public void ShowInfo(string message) => Infos.Add(message);

        public void ShowInstrumentSelection(InstrumentSelectionViewModel viewModel) { }

        public string? SelectTemplateFilePath() => null;
    }

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));
    private readonly FakeDialogService _dialog = new();

    public ShellViewModelTests() {
        Directory.CreateDirectory(_directory);
    }

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private void WriteProduct(string fileName, string productName) =>
        File.WriteAllText(Path.Combine(_directory, fileName), ProductJson.Replace("{NAME}", productName));

    private ShellViewModel CreateShell(ProductCatalogResult catalog, Func<ProductMenuItem, MainViewModel>? open = null) =>
        new(catalog, open ?? (item => CreateMain(item)), _dialog);

    private MainViewModel CreateMain(ProductMenuItem item) =>
        new(ProductDefinitionLoader.LoadFile(item.DefinitionPath), new SessionStore(Path.Combine(_directory, "sessions")), _dialog);

    // ===== 製品の一覧 =====

    [Fact]
    public void ScanListsProductsByNameInOrder() {
        WriteProduct("b.json", "BBB");
        WriteProduct("a.json", "AAA");

        var result = ProductCatalog.Scan(_directory);

        Assert.Equal(["AAA", "BBB"], result.Items.Select(item => item.Name));
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void ScanSkipsBrokenFilesAndReportsThem() {
        WriteProduct("ok.json", "OK");
        File.WriteAllText(Path.Combine(_directory, "broken.json"), "{ これはJSONではない");

        var result = ProductCatalog.Scan(_directory);

        Assert.Equal("OK", Assert.Single(result.Items).Name);
        Assert.Contains("broken.json", Assert.Single(result.Warnings));
    }

    [Fact]
    public void ScanReportsDuplicateProductNames() {
        WriteProduct("a.json", "SAME");
        WriteProduct("b.json", "SAME");

        var result = ProductCatalog.Scan(_directory);

        Assert.Single(result.Items);
        Assert.Contains("b.json", Assert.Single(result.Warnings));
    }

    [Fact]
    public void ScanReportsMissingFolder() {
        var result = ProductCatalog.Scan(Path.Combine(_directory, "none"));

        Assert.Empty(result.Items);
        Assert.Single(result.Warnings);
    }

    [Fact]
    public void ShippedEl5000AcDefinitionIsListed() {
        var result = ProductCatalog.Scan(Path.GetDirectoryName(RepositoryFiles.Find("products", "EL5000AC.json"))!);

        Assert.Contains(result.Items, item => item.Name == "EL5000AC");
    }

    // ===== 切り替え =====

    [Fact]
    public void StartsOnTheMenu() {
        WriteProduct("a.json", "AAA");
        var shell = CreateShell(ProductCatalog.Scan(_directory));

        Assert.True(shell.IsMenuVisible);
        Assert.Null(shell.Current);
    }

    [Fact]
    public void SelectingAProductOpensItsInspectionScreen() {
        WriteProduct("a.json", "AAA");
        var catalog = ProductCatalog.Scan(_directory);
        var shell = CreateShell(catalog);

        shell.SelectProductCommand.Execute(catalog.Items[0]);

        Assert.False(shell.IsMenuVisible);
        Assert.Equal("AAA", shell.Current!.Product);
        // 手順表のない起動では、測定器の操作を出さない（手入力だけで検査する）
        Assert.False(shell.Current.HasInstrumentSelection);
    }

    [Fact]
    public void FailureToOpenKeepsTheMenuAndShowsTheError() {
        WriteProduct("a.json", "AAA");
        var catalog = ProductCatalog.Scan(_directory);
        var shell = CreateShell(catalog, _ => throw new ProductDefinitionException("壊れています"));

        shell.SelectProductCommand.Execute(catalog.Items[0]);

        Assert.True(shell.IsMenuVisible);
        Assert.Contains("壊れています", Assert.Single(_dialog.Errors));
    }

    [Fact]
    public void BackToMenuReturnsToTheMenuAndTellsThatTheSavedStateRemains() {
        WriteProduct("a.json", "AAA");
        var catalog = ProductCatalog.Scan(_directory);
        var shell = CreateShell(catalog);
        shell.SelectProductCommand.Execute(catalog.Items[0]);

        shell.BackToMenuCommand.Execute(null);

        Assert.True(shell.IsMenuVisible);
        Assert.Contains("続きから再開", Assert.Single(_dialog.Infos));
    }

    [Fact]
    public void BackToMenuOnTheMenuDoesNothing() {
        var shell = CreateShell(ProductCatalog.Scan(_directory));

        shell.BackToMenuCommand.Execute(null);

        Assert.Empty(_dialog.Infos);
    }

    [Fact]
    public void WarningTextIsShownOnlyWhenSomeDefinitionCouldNotBeRead() {
        WriteProduct("a.json", "AAA");
        Assert.False(CreateShell(ProductCatalog.Scan(_directory)).HasWarning);

        File.WriteAllText(Path.Combine(_directory, "broken.json"), "x");

        var shell = CreateShell(ProductCatalog.Scan(_directory));
        Assert.True(shell.HasWarning);
        Assert.Contains("broken.json", shell.WarningText);
    }
}
