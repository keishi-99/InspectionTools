using System.IO.Compression;
using ClosedXML.Excel;
using InspectionTools.Inspection;

namespace InspectionTools.Excel.Tests;

/// <summary>検査結果をテンプレートへ書き込む規則（値、確認項目、未実施、他の台、判定の数式、保存の安全性）を確認する</summary>
public sealed class TemplateWriterTests : IDisposable {

    // 項目の位置（0: 確認、1: 測定、2: 入力、3: 作業）。Excel の行は 5、8、9（作業項目は行を持たない）
    private const int ConfirmationIndex = 0;
    private const int MeasurementIndex = 1;
    private const int InputIndex = 2;

    private static readonly TransferSettings Settings = new(
        DataSheet: "ACdata",
        UnitLayout: new UnitLayout("F"),
        Check: new TemplateCheck("ACdata", "A3", "EL5000"));

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private string TemplatePath => Path.Combine(_directory, "template.xlsx");

    // 前回の検査の値が残っているテンプレートを作る（F・G列が検査する2台、H列が検査しない3台目）
    private void CreateTemplate(string productName = "EL5000", string sheetName = "ACdata") {
        Directory.CreateDirectory(_directory);
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet(sheetName);
        workbook.AddWorksheet("Other");
        sheet.Cell("A3").Value = productName;
        sheet.Cell("F4").FormulaA1 = "=1+1";

        foreach (var column in new[] { "F", "G", "H" }) {
            sheet.Cell($"{column}5").Value = "良";
            sheet.Cell($"{column}8").Value = 999;
            sheet.Cell($"{column}9").Value = 888;
        }
        // 書式（塗りつぶし）が書き込み後も残ることを確かめるための目印
        sheet.Cell("F8").Style.Fill.BackgroundColor = XLColor.Yellow;
        workbook.SaveAs(TemplatePath);
    }

    // 台1は全項目が実施済み、台2は一部だけ実施した状態
    private static InspectionSession CreateSession() {
        List<InspectionItem> items = [
            new("電源投入", ItemKind.Confirmation, excelRow: 5),
            new("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m),
            new("温度入力", ItemKind.Input, excelRow: 9, lowerLimit: 0m, upperLimit: 10m),
            new("調整作業", ItemKind.Task),
        ];
        var session = new InspectionSession(items, unitCount: 2);

        session.SetConfirmation(ConfirmationIndex, 0, isOk: true);
        session.SetValue(MeasurementIndex, 0, 146.745m, isManual: false);
        session.SetValue(InputIndex, 0, 5m, isManual: true);
        session.CompleteTask(3, 0);

        session.SetConfirmation(ConfirmationIndex, 1, isOk: false);
        session.SetValue(MeasurementIndex, 1, 300m, isManual: false);
        return session;
    }

    // 書き込み後のファイルを開いて、セルを読む
    private static IXLCell ReadCell(string path, string address) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = new XLWorkbook(stream);
        return workbook.Worksheet("ACdata").Cell(address);
    }

    [Fact]
    public void ValuesAreWrittenAsNumbersToTheCellOfEachUnit() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession());

        Assert.Equal(146.745, ReadCell(TemplatePath, "F8").Value.GetNumber());
        Assert.Equal(5, ReadCell(TemplatePath, "F9").Value.GetNumber());
        Assert.Equal(300, ReadCell(TemplatePath, "G8").Value.GetNumber());
    }

    [Fact]
    public void ValueOutOfRangeIsStillWrittenBecauseTheTemplateJudgesIt() {
        CreateTemplate();
        var session = CreateSession();
        Assert.Equal(ResultState.Ng, session.GetResult(MeasurementIndex, 1).State);

        TemplateWriter.Write(TemplatePath, Settings, session);

        Assert.Equal(300, ReadCell(TemplatePath, "G8").Value.GetNumber());
    }

    [Fact]
    public void ConfirmationOkWritesTheOkTextAndNgWritesNothing() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession());

        Assert.Equal("良", ReadCell(TemplatePath, "F5").GetString());
        Assert.True(ReadCell(TemplatePath, "G5").IsEmpty());
    }

    // チェックボックスのリンク先の列がある製品（PAF5）では、確認項目は値の列ではなくその列に TRUE/FALSE を書き、不合格と未実施は FALSE にする
    [Fact]
    public void ConfirmationIsWrittenAsTrueOrFalseToTheConfirmationColumn() {
        CreateTemplate();
        var settings = Settings with { UnitLayout = new UnitLayout("F", confirmationColumn: "K") };
        var session = CreateSession();

        TemplateWriter.Write(TemplatePath, settings, session);

        Assert.True(ReadCell(TemplatePath, "K5").Value.GetBoolean());
        Assert.False(ReadCell(TemplatePath, "L5").Value.GetBoolean());
        // 値の列の「良」には触れない
        Assert.Equal("良", ReadCell(TemplatePath, "F5").GetString());
        Assert.Equal("良", ReadCell(TemplatePath, "G5").GetString());
    }

    [Fact]
    public void NotDoneConfirmationIsWrittenAsFalseToo() {
        CreateTemplate();
        var settings = Settings with { UnitLayout = new UnitLayout("F", confirmationColumn: "K") };
        var session = new InspectionSession([new("電源投入", ItemKind.Confirmation, excelRow: 5)], unitCount: 1);

        TemplateWriter.Write(TemplatePath, settings, session);

        Assert.False(ReadCell(TemplatePath, "K5").Value.GetBoolean());
    }

    [Fact]
    public void WithoutProductCheckAnyTemplateIsWritten() {
        CreateTemplate(productName: "EL9240");
        var settings = Settings with { Check = null };

        TemplateWriter.Write(TemplatePath, settings, CreateSession());

        Assert.Equal(146.745, ReadCell(TemplatePath, "F8").Value.GetNumber());
    }

    [Fact]
    public void UnfinishedCellsOfInspectedUnitsAreClearedOfOldValues() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession());

        // 台2の入力項目は未実施。前回の値（888）を消して空欄にする
        Assert.True(ReadCell(TemplatePath, "G9").IsEmpty());
    }

    [Fact]
    public void ColumnsOfUnitsThatWereNotInspectedAreNotTouched() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession());

        Assert.Equal(999, ReadCell(TemplatePath, "H8").Value.GetNumber());
        Assert.Equal("良", ReadCell(TemplatePath, "H5").GetString());
    }

    [Fact]
    public void FormulasAndFormattingAreKept() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession());

        Assert.True(ReadCell(TemplatePath, "F4").HasFormula);
        Assert.Equal(XLColor.Yellow, ReadCell(TemplatePath, "F8").Style.Fill.BackgroundColor);
    }

    [Fact]
    public void WrappedLayoutWritesLaterUnitsToTheRowsBelow() {
        CreateTemplate();
        var settings = Settings with { UnitLayout = new UnitLayout("F", unitsPerBlock: 1, blockRowOffset: 20) };

        TemplateWriter.Write(TemplatePath, settings, CreateSession());

        Assert.Equal(146.745, ReadCell(TemplatePath, "F8").Value.GetNumber());
        Assert.Equal(300, ReadCell(TemplatePath, "F28").Value.GetNumber());
    }

    [Fact]
    public void ExcelIsAskedToRecalculateEverythingWhenTheFileIsOpened() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession());

        using var zip = ZipFile.OpenRead(TemplatePath);
        using var reader = new StreamReader(zip.GetEntry("xl/workbook.xml")!.Open());
        Assert.Contains("fullCalcOnLoad=\"1\"", reader.ReadToEnd());
    }

    // ClosedXML は保存し直すと、シートの表示形式（改ページプレビュー）を通常表示に戻してしまうため、すべてのシートを改ページプレビューにそろえる
    [Fact]
    public void AllSheetsAreSetToPageBreakPreview() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession());

        using var zip = ZipFile.OpenRead(TemplatePath);
        var sheetEntries = zip.Entries.Where(entry => entry.FullName.StartsWith("xl/worksheets/sheet") && entry.FullName.EndsWith(".xml")).ToList();
        Assert.Equal(2, sheetEntries.Count);
        foreach (var entry in sheetEntries) {
            using var reader = new StreamReader(entry.Open());
            Assert.Contains("view=\"pageBreakPreview\"", reader.ReadToEnd());
        }
    }

    [Fact]
    public void NoTemporaryFileIsLeftAfterSuccess() {
        CreateTemplate();

        TemplateWriter.Write(TemplatePath, Settings, CreateSession());

        Assert.Equal([TemplatePath], Directory.GetFiles(_directory));
    }

    [Fact]
    public void FileOpenedByAnotherProgramFailsWithoutChangingTheFile() {
        CreateTemplate();
        var before = File.ReadAllBytes(TemplatePath);
        using var otherProgram = new FileStream(TemplatePath, FileMode.Open, FileAccess.Read, FileShare.Read);

        var exception = Assert.Throws<TemplateWriteException>(() => TemplateWriter.Write(TemplatePath, Settings, CreateSession()));

        Assert.Contains("閉じて", exception.Message);
        Assert.Equal(before, File.ReadAllBytes(TemplatePath));
        Assert.Equal([TemplatePath], Directory.GetFiles(_directory));
    }

    [Fact]
    public void DifferentProductIsRejectedBeforeWritingAnything() {
        CreateTemplate(productName: "EL9240");
        var before = File.ReadAllBytes(TemplatePath);

        Assert.Throws<TemplateVerificationException>(() => TemplateWriter.Write(TemplatePath, Settings, CreateSession()));

        Assert.Equal(before, File.ReadAllBytes(TemplatePath));
    }

    [Fact]
    public void MissingDataSheetFailsWithoutChangingTheFile() {
        CreateTemplate();
        var before = File.ReadAllBytes(TemplatePath);
        var settings = Settings with { DataSheet = "DCdata" };

        Assert.Throws<TemplateReadException>(() => TemplateWriter.Write(TemplatePath, settings, CreateSession()));

        Assert.Equal(before, File.ReadAllBytes(TemplatePath));
        Assert.Equal([TemplatePath], Directory.GetFiles(_directory));
    }

    [Fact]
    public void MissingFileFailsWithReadException() {
        Assert.Throws<TemplateReadException>(() => TemplateWriter.Write(TemplatePath, Settings, CreateSession()));
    }
}
