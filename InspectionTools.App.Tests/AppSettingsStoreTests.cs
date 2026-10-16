using InspectionTools.App.Settings;

namespace InspectionTools.App.Tests;

/// <summary>設定ファイル(settings.json)の読み込みと保存を確認する</summary>
public sealed class AppSettingsStoreTests : IDisposable {

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));

    private string SettingsPath => Path.Combine(_directory, AppSettingsStore.FileName);

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    [Fact]
    public void MissingFileGivesTheDefaultSettingsWithoutWarning() {
        var result = AppSettingsStore.Load(SettingsPath);

        Assert.Equal(AppSettings.Default, result.Settings);
        Assert.Null(result.Warning);
    }

    [Fact]
    public void DefaultThemeIsLight() {
        Assert.Equal(ThemeSetting.Light, AppSettings.Default.Theme);
    }

    [Theory]
    [InlineData(ThemeSetting.Light)]
    [InlineData(ThemeSetting.Dark)]
    public void SavedSettingsAreReadBack(ThemeSetting theme) {
        AppSettingsStore.Save(SettingsPath, new AppSettings(theme));

        var result = AppSettingsStore.Load(SettingsPath);

        Assert.Equal(theme, result.Settings.Theme);
        Assert.Null(result.Warning);
    }

    [Fact]
    public void SaveWritesTheThemeAsLowerCaseText() {
        AppSettingsStore.Save(SettingsPath, new AppSettings(ThemeSetting.Dark));

        Assert.Contains("\"theme\": \"dark\"", File.ReadAllText(SettingsPath));
    }

    [Fact]
    public void SaveCreatesTheFolderAndLeavesNoTemporaryFile() {
        AppSettingsStore.Save(SettingsPath, new AppSettings(ThemeSetting.Dark));

        Assert.True(File.Exists(SettingsPath));
        Assert.False(File.Exists(SettingsPath + ".tmp"));
    }

    [Fact]
    public void SaveReplacesAnExistingFile() {
        AppSettingsStore.Save(SettingsPath, new AppSettings(ThemeSetting.Dark));
        AppSettingsStore.Save(SettingsPath, new AppSettings(ThemeSetting.Light));

        Assert.Equal(ThemeSetting.Light, AppSettingsStore.Load(SettingsPath).Settings.Theme);
    }

    [Theory]
    [InlineData("""{ "theme": "DARK" }""")]
    [InlineData("""{ "Theme": "Dark" }""")]
    public void ThemeNameIsCaseInsensitive(string json) {
        var result = AppSettingsStore.Parse(json);

        Assert.Equal(ThemeSetting.Dark, result.Settings.Theme);
        Assert.Null(result.Warning);
    }

    [Fact]
    public void MissingThemeGivesTheDefaultWithoutWarning() {
        var result = AppSettingsStore.Parse("{ }");

        Assert.Equal(AppSettings.Default, result.Settings);
        Assert.Null(result.Warning);
    }

    [Theory]
    [InlineData("{ これはJSONではない")]
    [InlineData("null")]
    [InlineData("")]
    public void BrokenJsonGivesTheDefaultWithAWarning(string json) {
        var result = AppSettingsStore.Parse(json);

        Assert.Equal(AppSettings.Default, result.Settings);
        Assert.Contains("settings.json", result.Warning);
    }

    [Theory]
    [InlineData("""{ "theme": "blue" }""")]
    [InlineData("""{ "theme": "" }""")]
    [InlineData("""{ "theme": "5" }""")]
    public void UnknownThemeGivesTheDefaultWithAWarning(string json) {
        var result = AppSettingsStore.Parse(json);

        Assert.Equal(AppSettings.Default, result.Settings);
        Assert.Contains("theme", result.Warning);
    }

    [Fact]
    public void UnreadableFileGivesTheDefaultWithAWarning() {
        // 他のところが、ほかの読み書きを許さずに開いている間は、読み込めない
        AppSettingsStore.Save(SettingsPath, new AppSettings(ThemeSetting.Dark));
        using var locked = new FileStream(SettingsPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

        var result = AppSettingsStore.Load(SettingsPath);

        Assert.Equal(AppSettings.Default, result.Settings);
        Assert.NotNull(result.Warning);
    }
}
