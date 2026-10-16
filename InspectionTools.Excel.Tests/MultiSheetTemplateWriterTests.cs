using ClosedXML.Excel;
using InspectionTools.Inspection;

namespace InspectionTools.Excel.Tests;

/// <summary>台を数台ずつ別のシートに書く製品（PAF5 など）の転記を確認する。5台目までは1枚目、6台目からは2枚目、のようにシートが切り替わる</summary>
public sealed class MultiSheetTemplateWriterTests : IDisposable {

    // 実物（PAF5 分離形）と同じく、1枚に2台ずつ（確認しやすいよう、実物の5台より少なくしている）の3枚。台の値は F 列から、台数のセルは C1
    private static readonly string[] SheetNames = ["S1", "S2", "S3"];

    private static readonly TransferSettings Settings = new(
        DataSheet: "S1",
        UnitLayout: new UnitLayout("F", confirmationColumn: "K"),
        Check: new TemplateCheck("S1", "B3", "分離形"),
        Split: new SheetSplit(SheetNames, UnitsPerSheet: 2),
        UnitCountCell: "C1");

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private string TemplatePath => Path.Combine(_directory, "template.xlsx");

    // 3枚のシートを持つテンプレートを作る。各シートの台数のセル（C1）は2、前回の値が全部の台の列に残っている
    private void CreateTemplate(string[]? sheetNames = null, string productCell = "PAF5 変換器ユニット(分離形)") {
        Directory.CreateDirectory(_directory);
        using var workbook = new XLWorkbook();
        foreach (var name in sheetNames ?? SheetNames) {
            var sheet = workbook.AddWorksheet(name);
            sheet.Cell("B3").Value = productCell;
            sheet.Cell("C1").Value = 2;
            foreach (var column in new[] { "F", "G" }) {
                sheet.Cell($"{column}8").Value = 999;
            }
        }
        workbook.SaveAs(TemplatePath);
    }

    // 測定項目(行8)と確認項目(行9)を持つ、unitCount 台の検査結果。台 n の測定値は n+1 * 10、確認はすべて合格
    private static InspectionSession CreateSession(int unitCount) {
        List<InspectionItem> items = [
            new("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 1000m),
            new("目視確認", ItemKind.Confirmation, excelRow: 9),
        ];
        var session = new InspectionSession(items, unitCount);
        for (var unit = 0; unit < unitCount; unit++) {
            session.SetValue(0, unit, (unit + 1) * 10m, isManual: true);
            session.SetConfirmation(1, unit, isOk: true);
        }
        return session;
    }

    private static IXLCell ReadCell(string path, string sheetName, string address) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = new XLWorkbook(stream);
        return workbook.Worksheet(sheetName).Cell(address);
    }

    [Fact]
    public void UnitsAreWrittenToTheSheetOfTheirGroupInOrder() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession(unitCount: 5));

        // 台1・2 → S1 の F・G 列、台3・4 → S2 の F・G 列、台5 → S3 の F 列
        Assert.Equal(10, ReadCell(TemplatePath, "S1", "F8").Value.GetNumber());
        Assert.Equal(20, ReadCell(TemplatePath, "S1", "G8").Value.GetNumber());
        Assert.Equal(30, ReadCell(TemplatePath, "S2", "F8").Value.GetNumber());
        Assert.Equal(40, ReadCell(TemplatePath, "S2", "G8").Value.GetNumber());
        Assert.Equal(50, ReadCell(TemplatePath, "S3", "F8").Value.GetNumber());
    }

    [Fact]
    public void ConfirmationsAreWrittenToTheConfirmationColumnOfTheSameSheet() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession(unitCount: 3));

        Assert.True(ReadCell(TemplatePath, "S1", "K9").Value.GetBoolean());
        Assert.True(ReadCell(TemplatePath, "S1", "L9").Value.GetBoolean());
        Assert.True(ReadCell(TemplatePath, "S2", "K9").Value.GetBoolean());
    }

    [Fact]
    public void UnitCountCellOfEachSheetGetsTheNumberOfUnitsOnThatSheet() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession(unitCount: 3));

        Assert.Equal(2, ReadCell(TemplatePath, "S1", "C1").Value.GetNumber());
        Assert.Equal(1, ReadCell(TemplatePath, "S2", "C1").Value.GetNumber());
    }

    [Fact]
    public void SheetsWithoutInspectedUnitsAreNotTouched() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession(unitCount: 3));

        // S3 には台がないので、前回の値も、台数のセルもそのまま
        Assert.Equal(999, ReadCell(TemplatePath, "S3", "F8").Value.GetNumber());
        Assert.Equal(2, ReadCell(TemplatePath, "S3", "C1").Value.GetNumber());
    }

    [Fact]
    public void UnitsOnTheLastUsedSheetThatWereNotInspectedKeepTheirOldValues() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession(unitCount: 3));

        // 台4（S2 の G 列）は検査していないので、前回の値のまま
        Assert.Equal(999, ReadCell(TemplatePath, "S2", "G8").Value.GetNumber());
    }

    [Fact]
    public void UnitCountCellIsWrittenEvenWhenOnlyOneSheetIsNeeded() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession(unitCount: 1));

        Assert.Equal(1, ReadCell(TemplatePath, "S1", "C1").Value.GetNumber());
        Assert.Equal(2, ReadCell(TemplatePath, "S2", "C1").Value.GetNumber());
    }

    [Fact]
    public void UnitCountCellIsNotWrittenWhenTheDefinitionHasNone() {
        CreateTemplate();
        var settings = Settings with { UnitCountCell = null };

        TemplateWriter.Write(TemplatePath, settings, CreateSession(unitCount: 1));

        Assert.Equal(2, ReadCell(TemplatePath, "S1", "C1").Value.GetNumber());
    }

    // 単一のシートの製品でも、台数のセルを書く設定があれば、検査した台数を書く
    [Fact]
    public void UnitCountCellIsWrittenForASingleSheetProductToo() {
        CreateTemplate();
        var settings = Settings with { Split = null };

        TemplateWriter.Write(TemplatePath, settings, CreateSession(unitCount: 2));

        Assert.Equal(2, ReadCell(TemplatePath, "S1", "C1").Value.GetNumber());
    }

    // 必要なシートがないときは、何も書き込まずに、シート名を知らせる（途中のシートまでしか書かれない状態を防ぐ）
    [Fact]
    public void MissingSheetFailsBeforeWritingAnythingAndNamesTheSheet() {
        CreateTemplate(sheetNames: ["S1", "S2"]);
        var before = File.ReadAllBytes(TemplatePath);

        var exception = Assert.Throws<TemplateReadException>(() => TemplateWriter.Write(TemplatePath, Settings, CreateSession(unitCount: 5)));

        Assert.Contains("'S3'", exception.Message);
        Assert.Equal(before, File.ReadAllBytes(TemplatePath));
        Assert.Equal([TemplatePath], Directory.GetFiles(_directory));
    }

    [Fact]
    public void SheetsThatAreNotNeededMayBeMissing() {
        CreateTemplate(sheetNames: ["S1", "S2"]);

        TemplateWriter.Write(TemplatePath, Settings, CreateSession(unitCount: 4));

        Assert.Equal(30, ReadCell(TemplatePath, "S2", "F8").Value.GetNumber());
    }

    // 製品の確認は、使うシートすべてで行う（途中のシートだけ別のタイプのテンプレートだった、を防ぐ）
    [Fact]
    public void ProductCheckIsDoneOnEveryUsedSheet() {
        CreateTemplate();
        using (var workbook = new XLWorkbook(TemplatePath)) {
            workbook.Worksheet("S2").Cell("B3").Value = "PAF5 変換器ユニット(一体形)";
            workbook.SaveAs(TemplatePath);
        }
        var before = File.ReadAllBytes(TemplatePath);

        var exception = Assert.Throws<TemplateVerificationException>(() => TemplateWriter.Write(TemplatePath, Settings, CreateSession(unitCount: 3)));

        Assert.Contains("S2!B3", exception.Message);
        Assert.Equal(before, File.ReadAllBytes(TemplatePath));
    }

    [Fact]
    public void ProductCheckIsNotDoneOnSheetsThatAreNotUsed() {
        CreateTemplate();
        using (var workbook = new XLWorkbook(TemplatePath)) {
            workbook.Worksheet("S3").Cell("B3").Value = "別のタイプ";
            workbook.SaveAs(TemplatePath);
        }

        TemplateWriter.Write(TemplatePath, Settings, CreateSession(unitCount: 3));

        Assert.Equal(30, ReadCell(TemplatePath, "S2", "F8").Value.GetNumber());
    }

    [Fact]
    public void SheetNamesOfTheSplitFollowTheUnitCount() {
        Assert.Equal(["S1"], Settings.GetSheetNames(unitCount: 1));
        Assert.Equal(["S1"], Settings.GetSheetNames(unitCount: 2));
        Assert.Equal(["S1", "S2"], Settings.GetSheetNames(unitCount: 3));
        Assert.Equal(["S1", "S2", "S3"], Settings.GetSheetNames(unitCount: 5));
        Assert.Throws<ArgumentOutOfRangeException>(() => Settings.GetSheetNames(unitCount: 7));
    }

    [Fact]
    public void SingleSheetSettingsUseOnlyTheDataSheet() {
        var settings = Settings with { Split = null };

        Assert.Equal(["S1"], settings.GetSheetNames(unitCount: 30));
    }
}
