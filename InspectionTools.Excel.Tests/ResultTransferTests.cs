using ClosedXML.Excel;
using InspectionTools.Inspection;

namespace InspectionTools.Excel.Tests;

/// <summary>転記が成功したときだけ途中保存が削除され、失敗したときは残ることを確認する</summary>
public sealed class ResultTransferTests : IDisposable {

    private const string Product = "EL5000AC";

    private static readonly TransferSettings Settings = new(
        DataSheet: "ACdata",
        UnitLayout: new UnitLayout("F"),
        Check: new TemplateCheck("ACdata", "A3", "EL5000"));

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));
    private readonly SessionStore _store;

    public ResultTransferTests() {
        _store = new SessionStore(Path.Combine(_directory, "sessions"));
    }

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private string TemplatePath => Path.Combine(_directory, "template.xlsx");
    private string SavedStatePath => Path.Combine(_directory, "sessions", Product + ".json");

    private void CreateTemplate(string productName = "EL5000") {
        Directory.CreateDirectory(_directory);
        using var workbook = new XLWorkbook();
        workbook.AddWorksheet("ACdata").Cell("A3").Value = productName;
        workbook.SaveAs(TemplatePath);
    }

    // 測定項目を1つ持つ、1台の検査結果を作り、途中保存まで済ませた状態にする
    private InspectionSession CreateSavedSession() {
        List<InspectionItem> items = [new("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m)];
        var session = new InspectionSession(items, unitCount: 1);
        session.SetValue(0, 0, 146.745m, isManual: false);
        _store.Save(Product, session);
        return session;
    }

    [Fact]
    public void SuccessfulTransferWritesTheTemplateAndDeletesTheSavedState() {
        CreateTemplate();
        var session = CreateSavedSession();

        ResultTransfer.Transfer(TemplatePath, Settings, session, _store, Product);

        using var workbook = new XLWorkbook(TemplatePath);
        Assert.Equal(146.745, workbook.Worksheet("ACdata").Cell("F8").Value.GetNumber());
        Assert.False(File.Exists(SavedStatePath));
    }

    [Fact]
    public void FailingToDeleteTheSavedStateAfterTransferIsReportedSeparately() {
        CreateTemplate();
        var session = CreateSavedSession();
        // 途中保存のファイルを別のプログラムが開いている状態にして、削除だけを失敗させる
        using var otherProgram = new FileStream(SavedStatePath, FileMode.Open, FileAccess.Read, FileShare.None);

        Assert.Throws<SavedStateDeleteException>(() => ResultTransfer.Transfer(TemplatePath, Settings, session, _store, Product));

        using var workbook = new XLWorkbook(TemplatePath);
        Assert.Equal(146.745, workbook.Worksheet("ACdata").Cell("F8").Value.GetNumber());
    }

    [Fact]
    public void TransferToDifferentProductKeepsTheSavedState() {
        CreateTemplate(productName: "EL9240");
        var session = CreateSavedSession();

        Assert.Throws<TemplateVerificationException>(() => ResultTransfer.Transfer(TemplatePath, Settings, session, _store, Product));

        Assert.True(File.Exists(SavedStatePath));
    }

    [Fact]
    public void TransferToFileOpenedByAnotherProgramKeepsTheSavedState() {
        CreateTemplate();
        var session = CreateSavedSession();
        using var otherProgram = new FileStream(TemplatePath, FileMode.Open, FileAccess.Read, FileShare.Read);

        Assert.Throws<TemplateWriteException>(() => ResultTransfer.Transfer(TemplatePath, Settings, session, _store, Product));

        Assert.True(File.Exists(SavedStatePath));
    }

    [Fact]
    public void TransferToMissingFileKeepsTheSavedState() {
        var session = CreateSavedSession();

        Assert.Throws<TemplateReadException>(() => ResultTransfer.Transfer(TemplatePath, Settings, session, _store, Product));

        Assert.True(File.Exists(SavedStatePath));
    }

    [Fact]
    public void SavedStateCanBeRestoredAndTransferredAgainAfterAFailure() {
        CreateTemplate();
        var session = CreateSavedSession();
        using (var otherProgram = new FileStream(TemplatePath, FileMode.Open, FileAccess.Read, FileShare.Read)) {
            Assert.Throws<TemplateWriteException>(() => ResultTransfer.Transfer(TemplatePath, Settings, session, _store, Product));
        }

        // Excel を閉じたあと、アプリを起動し直した想定: 途中保存から復元して、もう一度転記する
        var restored = _store.Load(Product, session.Items).Session!;
        ResultTransfer.Transfer(TemplatePath, Settings, restored, _store, Product);

        using var workbook = new XLWorkbook(TemplatePath);
        Assert.Equal(146.745, workbook.Worksheet("ACdata").Cell("F8").Value.GetNumber());
        Assert.False(File.Exists(SavedStatePath));
    }
}
