namespace InspectionTools.Products.Tests;

/// <summary>測定器の一覧（VisaAddress.xml）の読み込みと、役割ごとの候補の絞り込みを確認する</summary>
public class InstrumentCatalogTests {

    // 旧アプリと同じ形式（名前空間つきの dsInstList / tblInstList）
    private const string SampleXml = """
        <?xml version="1.0" standalone="yes"?>
        <dsInstList xmlns="http://tempuri.org/dsInstList.xsd">
          <tblInstList>
            <Category>デジタルマルチメータ</Category>
            <Name>DMM-01</Name>
            <VisaAddress>1</VisaAddress>
            <SignalType>1</SignalType>
          </tblInstList>
          <tblInstList>
            <Category>デジタルマルチメータ</Category>
            <Name>DMM-02</Name>
            <VisaAddress>USB0::0x1334::0x0204::SN000001::INSTR</VisaAddress>
            <SignalType>2</SignalType>
          </tblInstList>
          <tblInstList>
            <Category>ファンクションジェネレータ</Category>
            <Name>FG-01</Name>
            <VisaAddress>USB0::0x1::0x2::SN1::INSTR</VisaAddress>
            <SignalType>2</SignalType>
          </tblInstList>
          <tblInstList>
            <Category>ユニバーサルカウンタ</Category>
            <Name>CNT-01</Name>
            <VisaAddress>GPIB0::22::INSTR</VisaAddress>
            <SignalType>3</SignalType>
          </tblInstList>
        </dsInstList>
        """;

    [Fact]
    public void ParseReadsEveryEntryInFileOrder() {
        var catalog = InstrumentCatalog.Parse(SampleXml);

        Assert.Equal(["DMM-01", "DMM-02", "FG-01", "CNT-01"], catalog.Entries.Select(entry => entry.Name));
        Assert.Empty(catalog.Warnings);
        Assert.Equal(new InstrumentEntry("ユニバーサルカウンタ", "CNT-01", "GPIB0::22::INSTR", 3), catalog.Entries[3]);
    }

    [Fact]
    public void FindCandidatesFiltersByCategoryAndSignalType() {
        var catalog = InstrumentCatalog.Parse(SampleXml);

        var multimeters = catalog.FindCandidates(new RoleRequirement("デジタルマルチメータ", [1, 2]));
        var usbOnly = catalog.FindCandidates(new RoleRequirement("デジタルマルチメータ", [2]));
        var counters = catalog.FindCandidates(new RoleRequirement("ユニバーサルカウンタ", [2]));

        Assert.Equal(["DMM-01", "DMM-02"], multimeters.Select(entry => entry.Name));
        Assert.Equal(["DMM-02"], usbOnly.Select(entry => entry.Name));
        Assert.Empty(counters);
    }

    [Fact]
    public void ParseWorksWithoutNamespace() {
        var xml = "<dsInstList><tblInstList><Category>A</Category><Name>X</Name><VisaAddress>1</VisaAddress><SignalType>1</SignalType></tblInstList></dsInstList>";

        Assert.Equal(["X"], InstrumentCatalog.Parse(xml).Entries.Select(entry => entry.Name));
    }

    [Fact]
    public void ParseTrimsSurroundingWhitespace() {
        var xml = "<dsInstList><tblInstList><Category> A </Category><Name> X </Name><VisaAddress> 1 </VisaAddress><SignalType> 1 </SignalType></tblInstList></dsInstList>";

        Assert.Equal(new InstrumentEntry("A", "X", "1", 1), InstrumentCatalog.Parse(xml).Entries.Single());
    }

    [Theory]
    [InlineData("<Category></Category><Name>X</Name><VisaAddress>1</VisaAddress><SignalType>1</SignalType>")]
    [InlineData("<Category>A</Category><Name></Name><VisaAddress>1</VisaAddress><SignalType>1</SignalType>")]
    [InlineData("<Category>A</Category><Name>X</Name><VisaAddress></VisaAddress><SignalType>1</SignalType>")]
    [InlineData("<Category>A</Category><Name>X</Name><SignalType>1</SignalType>")]
    [InlineData("<Category>A</Category><Name>X</Name><VisaAddress>1</VisaAddress>")]
    [InlineData("<Category>A</Category><Name>X</Name><VisaAddress>1</VisaAddress><SignalType>0</SignalType>")]
    [InlineData("<Category>A</Category><Name>X</Name><VisaAddress>1</VisaAddress><SignalType>5</SignalType>")]
    [InlineData("<Category>A</Category><Name>X</Name><VisaAddress>1</VisaAddress><SignalType>abc</SignalType>")]
    public void EntryWithMissingOrInvalidFieldIsSkippedWithWarning(string fields) {
        var xml = $"<dsInstList><tblInstList>{fields}</tblInstList></dsInstList>";

        var catalog = InstrumentCatalog.Parse(xml);

        Assert.Empty(catalog.Entries);
        Assert.Contains("1番目の測定器", Assert.Single(catalog.Warnings));
    }

    [Fact]
    public void InvalidEntryDoesNotHideTheOtherEntries() {
        var xml = "<dsInstList>"
            + "<tblInstList><Category>A</Category><Name>BAD</Name><VisaAddress>1</VisaAddress><SignalType>9</SignalType></tblInstList>"
            + "<tblInstList><Category>A</Category><Name>GOOD</Name><VisaAddress>2</VisaAddress><SignalType>2</SignalType></tblInstList>"
            + "</dsInstList>";

        var catalog = InstrumentCatalog.Parse(xml);

        Assert.Equal(["GOOD"], catalog.Entries.Select(entry => entry.Name));
        Assert.Contains("BAD", Assert.Single(catalog.Warnings));
    }

    [Fact]
    public void SameNameInTheSameCategoryIsSkippedButSameNameInAnotherCategoryIsKept() {
        var xml = "<dsInstList>"
            + "<tblInstList><Category>A</Category><Name>X</Name><VisaAddress>1</VisaAddress><SignalType>1</SignalType></tblInstList>"
            + "<tblInstList><Category>A</Category><Name>X</Name><VisaAddress>2</VisaAddress><SignalType>1</SignalType></tblInstList>"
            + "<tblInstList><Category>B</Category><Name>X</Name><VisaAddress>3</VisaAddress><SignalType>1</SignalType></tblInstList>"
            + "</dsInstList>";

        var catalog = InstrumentCatalog.Parse(xml);

        Assert.Equal(["1", "3"], catalog.Entries.Select(entry => entry.VisaAddress));
        Assert.Single(catalog.Warnings);
    }

    [Theory]
    [InlineData("")]
    [InlineData("これはXMLではない")]
    [InlineData("<dsInstList><tblInstList>")]
    public void BrokenXmlThrowsCatalogException(string xml) {
        Assert.Throws<InstrumentCatalogException>(() => InstrumentCatalog.Parse(xml));
    }

    [Fact]
    public void LoadFileWithMissingFileThrowsCatalogExceptionWithFileName() {
        var path = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"), "VisaAddress.xml");

        var exception = Assert.Throws<InstrumentCatalogException>(() => InstrumentCatalog.LoadFile(path));

        Assert.Contains("VisaAddress.xml", exception.Message);
    }

    [Fact]
    public void RepositoryTemplateCanBeLoadedWithoutWarnings() {
        var catalog = InstrumentCatalog.LoadFile(RepositoryFiles.Find("VisaAddress.template.xml"));

        Assert.Empty(catalog.Warnings);
        Assert.NotEmpty(catalog.Entries);
    }
}
