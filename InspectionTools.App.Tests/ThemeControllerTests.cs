using InspectionTools.App.Settings;
using InspectionTools.App.ViewModels;

namespace InspectionTools.App.Tests;

/// <summary>配色の切り替えと、設定ファイルへの保存を、画面なしで確認する</summary>
public sealed class ThemeControllerTests : IDisposable {

    // 適用された配色を記録するダミー
    private sealed class FakeThemeApplier : IThemeApplier {
        public List<ThemeSetting> Applied { get; } = [];

        public void Apply(ThemeSetting theme) => Applied.Add(theme);
    }

    private sealed class FakeDialogService : IDialogService {
        public List<string> Errors { get; } = [];

        public bool Confirm(string message) => false;

        public void ShowError(string message) => Errors.Add(message);

        public List<string> Warnings { get; } = [];

        public void ShowWarning(string message) => Warnings.Add(message);

        public void ShowInfo(string message) { }

        public void ShowInstrumentSelection(InstrumentSelectionViewModel viewModel) { }

        public string? SelectTemplateFilePath() => null;
    }

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));
    private readonly FakeThemeApplier _applier = new();
    private readonly FakeDialogService _dialog = new();

    private string SettingsPath => Path.Combine(_directory, AppSettingsStore.FileName);

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
        if (File.Exists(_directory)) {
            File.Delete(_directory);
        }
    }

    private ThemeController CreateController() => new(SettingsPath, _applier, _dialog);

    [Fact]
    public void InitializeWithoutSettingsAppliesLightAndDoesNotSave() {
        var controller = CreateController();

        controller.Initialize();

        Assert.Equal([ThemeSetting.Light], _applier.Applied);
        Assert.False(controller.IsDark);
        Assert.Empty(_dialog.Warnings);
        Assert.False(File.Exists(SettingsPath));
    }

    [Fact]
    public void InitializeAppliesTheSavedTheme() {
        AppSettingsStore.Save(SettingsPath, new AppSettings(ThemeSetting.Dark));
        var controller = CreateController();

        controller.Initialize();

        Assert.Equal([ThemeSetting.Dark], _applier.Applied);
        Assert.True(controller.IsDark);
    }

    [Fact]
    public void InitializeWithBrokenSettingsAppliesLightAndTellsTheUser() {
        Directory.CreateDirectory(_directory);
        File.WriteAllText(SettingsPath, "{ これはJSONではない");
        var controller = CreateController();

        controller.Initialize();

        Assert.Equal([ThemeSetting.Light], _applier.Applied);
        Assert.Contains("ライトで起動しました", Assert.Single(_dialog.Warnings));
    }

    [Fact]
    public void SetThemeAppliesAndSaves() {
        var controller = CreateController();
        controller.Initialize();

        controller.SetTheme(ThemeSetting.Dark);

        Assert.Equal([ThemeSetting.Light, ThemeSetting.Dark], _applier.Applied);
        Assert.True(controller.IsDark);
        Assert.Equal(ThemeSetting.Dark, AppSettingsStore.Load(SettingsPath).Settings.Theme);
    }

    [Fact]
    public void SavedThemeSurvivesRestart() {
        var first = CreateController();
        first.Initialize();
        first.SetTheme(ThemeSetting.Dark);

        var second = CreateController();
        second.Initialize();

        Assert.True(second.IsDark);
    }

    [Fact]
    public void SettingTheSameThemeDoesNothing() {
        var controller = CreateController();
        controller.Initialize();
        _applier.Applied.Clear();

        controller.SetTheme(ThemeSetting.Light);

        Assert.Empty(_applier.Applied);
        Assert.False(File.Exists(SettingsPath));
    }

    [Fact]
    public void SetThemeRaisesPropertyChangedForTheSwitch() {
        var controller = CreateController();
        controller.Initialize();
        var changed = new List<string?>();
        controller.PropertyChanged += (_, e) => changed.Add(e.PropertyName);

        controller.SetTheme(ThemeSetting.Dark);

        Assert.Contains(nameof(ThemeController.IsDark), changed);
    }

    [Fact]
    public void SaveFailureStillSwitchesTheThemeAndTellsTheUser() {
        // フォルダの代わりにファイルがあると、保存できない
        File.WriteAllText(_directory, "");
        var controller = CreateController();
        controller.Initialize();

        controller.SetTheme(ThemeSetting.Dark);

        Assert.True(controller.IsDark);
        Assert.Equal(ThemeSetting.Dark, _applier.Applied[^1]);
        Assert.Contains("保存できませんでした", Assert.Single(_dialog.Warnings));
    }
}
