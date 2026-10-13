using ClosedXML.Excel;

namespace InspectionTools.Excel.Tests;

/// <summary>仮のテンプレート（実物と同じ配置と数式で、テストの中で作る）から、項目名と合格範囲を読み取れることを確認する</summary>
public sealed class TemplateReaderTests : IDisposable {

    // 実物（EL5000AC）と同じ構成: 値の入力シートに項目名、設定シートに合格範囲
    private static readonly TemplateLayout Layout = new("ACdata", "ACsetting", FirstRow: 5, LastRow: 14);

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    // 実物のテンプレートの一部（確認、固定の範囲、数式の範囲、引き継がれる大項目、改行のある大項目）を再現したファイルを作る
    private string CreateTemplate() {
        Directory.CreateDirectory(_directory);
        var path = Path.Combine(_directory, "template.xlsx");

        using var workbook = new XLWorkbook();
        var data = workbook.AddWorksheet("ACdata");
        var setting = workbook.AddWorksheet("ACsetting");

        // 行5: 確認項目（合格範囲の代わりに「良」）
        data.Cell("A5").Value = "電源投入";
        data.Cell("C5").Value = "正常に立ち上がること";
        setting.Cell("I5").Value = "良";
        setting.Cell("F5").FormulaA1 = "=I5";

        // 行8: 固定の範囲
        data.Cell("A8").Value = "消費電流";
        data.Cell("C8").Value = "AC:230mA以下";
        setting.Cell("F8").Value = 0;
        setting.Cell("G8").Value = 230;

        // 行9: 中心値 ± 許容値の数式（規格の表示 24.0V± は名前に含めない）
        data.Cell("A9").Value = "圧力伝送器用電源の確認";
        data.Cell("C9").Value = "24.0V±";
        data.Cell("E9").Value = "1.2V";
        setting.Cell("I9").Value = 24;
        setting.Cell("J9").Value = 1.2;
        setting.Cell("F9").FormulaA1 = "=I9-J9";
        setting.Cell("G9").FormulaA1 = "=I9+J9";

        // 行10〜11: 大項目は先頭の行にだけ書かれ、次の行へ引き継がれる
        data.Cell("A10").Value = "発信器電源の確認";
        data.Cell("C10").Value = "SW16-0";
        data.Cell("D10").Value = "32.0V±";
        setting.Cell("F10").Value = 31.4;
        setting.Cell("G10").Value = 32.6;
        data.Cell("C11").Value = "SW16-1";
        setting.Cell("F11").Value = 22.8;
        setting.Cell("G11").Value = 25.2;

        // 行12: セルの中で改行された大項目名
        data.Cell("A12").Value = "パルス出力1～3の\nパルス幅の確認";
        data.Cell("C12").Value = "パルス出力1:";
        setting.Cell("F12").Value = 0.8;
        setting.Cell("G12").Value = 1.2;

        // 行13: 合格範囲のない行（読み取りの対象にしない）
        data.Cell("A13").Value = "備考";

        // 行20: 読む範囲の外
        data.Cell("A20").Value = "範囲外の項目";
        setting.Cell("F20").Value = 1;
        setting.Cell("G20").Value = 2;

        workbook.SaveAs(path);
        return path;
    }

    private IReadOnlyList<TemplateRow> ReadRows() => TemplateReader.ReadRows(CreateTemplate(), Layout);

    [Fact]
    public void TextInRangeCellMeansConfirmationRow() {
        var row = ReadRows().Single(r => r.Row == 5);

        Assert.True(row.IsConfirmation);
        Assert.Null(row.LowerLimit);
        Assert.Null(row.UpperLimit);
    }

    [Fact]
    public void NumericRangeIsRead() {
        var row = ReadRows().Single(r => r.Row == 8);

        Assert.False(row.IsConfirmation);
        Assert.Equal(0m, row.LowerLimit);
        Assert.Equal(230m, row.UpperLimit);
    }

    [Fact]
    public void RangeFromFormulaIsReadWithoutFloatingPointError() {
        // 24 - 1.2 は、そのまま計算すると 22.799999999999997 のような誤差が出る
        var row = ReadRows().Single(r => r.Row == 9);

        Assert.Equal(22.8m, row.LowerLimit);
        Assert.Equal(25.2m, row.UpperLimit);
    }

    [Theory]
    [InlineData(5, "電源投入 正常に立ち上がること")]
    [InlineData(8, "消費電流 AC:230mA以下")]
    [InlineData(9, "圧力伝送器用電源の確認")]
    [InlineData(10, "発信器電源の確認 SW16-0")]
    [InlineData(11, "発信器電源の確認 SW16-1")]
    [InlineData(12, "パルス出力1～3のパルス幅の確認 パルス出力1:")]
    public void NameJoinsGroupAndSubItemsWithoutToleranceText(int row, string expectedName) {
        Assert.Equal(expectedName, ReadRows().Single(r => r.Row == row).Name);
    }

    [Fact]
    public void RowsWithoutRangeAndRowsOutsideLayoutAreSkipped() {
        var rows = ReadRows();

        Assert.Equal([5, 8, 9, 10, 11, 12], rows.Select(r => r.Row));
    }

    [Fact]
    public void MissingFileThrows() {
        var path = Path.Combine(_directory, "not-found.xlsx");

        Assert.Throws<TemplateReadException>(() => TemplateReader.ReadRows(path, Layout));
    }

    [Fact]
    public void MissingSheetThrowsWithSheetName() {
        var path = CreateTemplate();

        var exception = Assert.Throws<TemplateReadException>(() =>
            TemplateReader.ReadRows(path, Layout with { SettingSheet = "ACsetting " }));

        Assert.Contains("'ACsetting '", exception.Message);
    }

    [Fact]
    public void FileThatIsNotWorkbookThrows() {
        Directory.CreateDirectory(_directory);
        var path = Path.Combine(_directory, "broken.xlsx");
        File.WriteAllText(path, "これはExcelのファイルではない");

        Assert.Throws<TemplateReadException>(() => TemplateReader.ReadRows(path, Layout));
    }

    [Fact]
    public void FileOpenedByAnotherProgramCanStillBeRead() {
        var path = CreateTemplate();
        using var otherProgram = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

        var rows = TemplateReader.ReadRows(path, Layout);

        Assert.NotEmpty(rows);
    }
}
