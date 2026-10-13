using ClosedXML.Excel;

namespace InspectionTools.Excel.Tests;

/// <summary>選んだファイルが正しい製品のテンプレートかを、シートの存在と目印のセルの値で確かめられることを確認する</summary>
public sealed class TemplateVerifierTests : IDisposable {

    // 実物（EL5000）と同じく、シート ACdata の A3 に製品名が入っている
    private static readonly TemplateCheck Check = new("ACdata", "A3", "EL5000");

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    // 指定したシート名と、A3 の値を持つファイルを作る
    private string CreateTemplate(string sheetName = "ACdata", string? productCell = "EL5000") {
        Directory.CreateDirectory(_directory);
        var path = Path.Combine(_directory, "template.xlsx");

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet(sheetName);
        if (productCell is not null) {
            sheet.Cell("A3").Value = productCell;
        }
        workbook.SaveAs(path);
        return path;
    }

    [Fact]
    public void MatchingTemplatePasses() {
        TemplateVerifier.Verify(CreateTemplate(), Check);
    }

    [Fact]
    public void SpacesAroundProductCellAreIgnored() {
        TemplateVerifier.Verify(CreateTemplate(productCell: " EL5000 "), Check);
    }

    [Fact]
    public void DifferentProductThrowsWithCellAddressAndBothValues() {
        var path = CreateTemplate(productCell: "EL9240");

        var exception = Assert.Throws<TemplateVerificationException>(() => TemplateVerifier.Verify(path, Check));

        Assert.Contains("ACdata!A3", exception.Message);
        Assert.Contains("'EL5000'", exception.Message);
        Assert.Contains("'EL9240'", exception.Message);
    }

    [Fact]
    public void EmptyProductCellThrows() {
        var path = CreateTemplate(productCell: null);

        var exception = Assert.Throws<TemplateVerificationException>(() => TemplateVerifier.Verify(path, Check));

        Assert.Contains("空欄", exception.Message);
    }

    [Fact]
    public void MissingSheetThrowsWithSheetName() {
        var path = CreateTemplate(sheetName: "DCdata");

        var exception = Assert.Throws<TemplateVerificationException>(() => TemplateVerifier.Verify(path, Check));

        Assert.Contains("'ACdata'", exception.Message);
    }

    [Fact]
    public void SheetNameMustMatchExactlyIncludingTrailingSpace() {
        var path = CreateTemplate(sheetName: "ACdata ");

        Assert.Throws<TemplateVerificationException>(() => TemplateVerifier.Verify(path, Check));
    }

    [Fact]
    public void MissingFileThrowsReadException() {
        var path = Path.Combine(_directory, "not-found.xlsx");

        Assert.Throws<TemplateReadException>(() => TemplateVerifier.Verify(path, Check));
    }
}
