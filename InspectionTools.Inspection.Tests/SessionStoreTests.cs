namespace InspectionTools.Inspection.Tests;

/// <summary>途中保存のファイルへの書き込み・読み込み・壊れたファイルの扱いを確認する（一時フォルダを使い、実際の保存先は触らない）</summary>
public sealed class SessionStoreTests : IDisposable {

    private const string Product = "EL5000AC";

    // 退避ファイル名に入る時刻を固定するための時計
    private sealed class FixedTimeProvider(DateTimeOffset now) : TimeProvider {
        public override DateTimeOffset GetUtcNow() => now;
        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));
    private readonly SessionStore _store;

    public SessionStoreTests() {
        _store = new SessionStore(_directory, new FixedTimeProvider(new DateTimeOffset(2026, 10, 14, 9, 30, 15, TimeSpan.Zero)));
    }

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private string SavedPath => Path.Combine(_directory, Product + ".json");
    private string BrokenPath => Path.Combine(_directory, Product + ".broken-20261014-093015.json");

    private static InspectionSession CreateSession() {
        var session = new InspectionSession(TestItems.Create(), unitCount: 2);
        session.SetValue(0, 0, 146.745m, isManual: false);
        session.CompleteTask(3, 1);
        return session;
    }

    [Fact]
    public void LoadWithoutFileReturnsNoSavedData() {
        var result = _store.Load(Product, TestItems.Create());

        Assert.Equal(LoadStatus.NoSavedData, result.Status);
        Assert.Null(result.Session);
    }

    [Fact]
    public void SaveCreatesDirectoryAndLoadRestoresSameState() {
        _store.Save(Product, CreateSession());

        var result = _store.Load(Product, TestItems.Create());

        Assert.Equal(LoadStatus.Restored, result.Status);
        Assert.Equal(146.745m, result.Session!.GetResult(0, 0).Value);
        Assert.Equal(ResultState.Done, result.Session.GetResult(3, 1).State);
    }

    [Fact]
    public void SaveOverwritesPreviousFileAndLeavesNoTemporaryFile() {
        var session = CreateSession();
        _store.Save(Product, session);
        session.SetValue(0, 0, 100m, isManual: false);

        _store.Save(Product, session);

        var result = _store.Load(Product, TestItems.Create());
        Assert.Equal(100m, result.Session!.GetResult(0, 0).Value);
        Assert.Equal([SavedPath], Directory.GetFiles(_directory));
    }

    [Theory]
    [InlineData("{ これはJSONではない")]
    [InlineData("null")]
    [InlineData("")]
    public void LoadBrokenFileReturnsFailedAndKeepsFileUnderAnotherName(string content) {
        Directory.CreateDirectory(_directory);
        File.WriteAllText(SavedPath, content);

        var result = _store.Load(Product, TestItems.Create());

        Assert.Equal(LoadStatus.Failed, result.Status);
        Assert.Null(result.Session);
        Assert.NotNull(result.Message);
        Assert.Equal(BrokenPath, result.BrokenFilePath);
        Assert.False(File.Exists(SavedPath));
        Assert.Equal(content, File.ReadAllText(BrokenPath));
    }

    [Fact]
    public void LoadFileThatDoesNotMatchItemsReturnsFailedAndKeepsFile() {
        _store.Save(Product, CreateSession());
        List<InspectionItem> fewerItems = [.. TestItems.Create().Take(3)];

        var result = _store.Load(Product, fewerItems);

        Assert.Equal(LoadStatus.Failed, result.Status);
        Assert.False(File.Exists(SavedPath));
        Assert.True(File.Exists(BrokenPath));
    }

    [Fact]
    public void SaveWorksAgainAfterBrokenFileWasSetAside() {
        Directory.CreateDirectory(_directory);
        File.WriteAllText(SavedPath, "壊れた内容");
        _store.Load(Product, TestItems.Create());

        _store.Save(Product, CreateSession());

        Assert.Equal(LoadStatus.Restored, _store.Load(Product, TestItems.Create()).Status);
    }

    [Fact]
    public void DeleteRemovesSavedFileSoThatNothingIsRestoredNextTime() {
        _store.Save(Product, CreateSession());

        _store.Delete(Product);

        Assert.False(File.Exists(SavedPath));
        Assert.Equal(LoadStatus.NoSavedData, _store.Load(Product, TestItems.Create()).Status);
    }

    [Fact]
    public void DeleteDoesNothingWhenThereIsNoSavedFileOrFolder() {
        // 保存先のフォルダ自体がまだない場合も、例外にしない
        _store.Delete(Product);

        Assert.False(Directory.Exists(_directory));
    }

    [Fact]
    public void DeleteLeavesOtherProductsAndBrokenFilesUntouched() {
        _store.Save(Product, CreateSession());
        _store.Save("EL5000DC", CreateSession());
        File.WriteAllText(BrokenPath, "壊れた内容");

        _store.Delete(Product);

        Assert.True(File.Exists(Path.Combine(_directory, "EL5000DC.json")));
        Assert.True(File.Exists(BrokenPath));
    }

    [Theory]
    [InlineData("")]
    [InlineData("EL/5000")]
    public void DeleteRejectsProductNameUnusableAsFileName(string product) {
        Assert.Throws<ArgumentException>(() => _store.Delete(product));
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    [InlineData("EL/5000")]
    [InlineData("EL:5000")]
    public void SaveRejectsProductNameUnusableAsFileName(string product) {
        Assert.Throws<ArgumentException>(() => _store.Save(product, CreateSession()));
    }
}
