namespace InspectionTools.Inspection.Tests;

/// <summary>リポジトリにある PAF5 分離形の定義（products/PAF5-分離形.json）が、テンプレートの先頭ブロック（行8〜48）と合っていることを確認する</summary>
public class Paf5SeparateDefinitionTests {

    private static readonly ProductDefinition Definition = ProductDefinitionLoader.LoadFile(FindDefinitionPath());

    // テストの実行場所から上の階層をたどって、リポジトリの products フォルダにある定義のパスを探す
    private static string FindDefinitionPath() {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent) {
            if (File.Exists(Path.Combine(directory.FullName, "InspectionTools.sln"))) {
                return Path.Combine(directory.FullName, "products", "PAF5-分離形.json");
            }
        }
        throw new FileNotFoundException("リポジトリのルート（InspectionTools.sln）が見つかりません。");
    }

    private static InspectionItem ItemAt(int excelRow) => Definition.Items.Single(item => item.ExcelRow == excelRow);

    [Fact]
    public void DefinitionHasTheExpectedProductAndItemCount() {
        Assert.Equal("PAF5-分離形", Definition.Product);
        Assert.Equal(35, Definition.Items.Count);
    }

    // 最初の項目は、台ごとに手入力（またはバーコードの読み取り）する LCD ユニットのシリアル
    [Fact]
    public void FirstItemIsTheLcdSerialTextInput() {
        var first = Definition.Items[0];

        Assert.Equal(ItemKind.Text, first.Kind);
        Assert.Equal(7, first.ExcelRow);
        Assert.Null(first.LowerLimit);
        Assert.False(string.IsNullOrWhiteSpace(first.Description));
    }

    [Fact]
    public void DefinitionHasTheExpectedKindCounts() {
        int Count(ItemKind kind) => Definition.Items.Count(item => item.Kind == kind);

        Assert.Equal(7, Count(ItemKind.Measurement));
        Assert.Equal(4, Count(ItemKind.Input));
        Assert.Equal(23, Count(ItemKind.Confirmation));
        Assert.Equal(1, Count(ItemKind.Text));
        Assert.Equal(0, Count(ItemKind.Task));
    }

    // LCD のシリアル（行7）と、行8〜48のうち見出しや注意書きだけの行（16、18、22、26、31、34、39）を除いた34行。
    // ローカット機能（行40〜47）だけは、検査の流れの順（min/F、min/7、max/7、max/F）に並べるので、Excel の行の順とは違う
    [Fact]
    public void RecordedItemsCoverTheExpectedExcelRowsInOrder() {
        int[] expected = [7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 19, 20, 21, 23, 24, 25, 27, 28, 29, 30, 32, 33, 35, 36, 37, 38, 44, 45, 40, 41, 42, 43, 46, 47, 48];

        Assert.Equal(expected, Definition.Items.Select(item => item.ExcelRow!.Value));
    }

    // チェックボックスがある行（テンプレートでは M〜Q 列にリンクしている行）だけが、確認項目
    [Fact]
    public void ConfirmationItemsAreTheRowsWithCheckBoxes() {
        int[] expected = [8, 9, 10, 11, 27, 28, 29, 30, 32, 33, 35, 36, 37, 38, 44, 45, 40, 41, 42, 43, 46, 47, 48];

        var confirmationRows = Definition.Items.Where(item => item.Kind == ItemKind.Confirmation).Select(item => item.ExcelRow!.Value);

        Assert.Equal(expected, confirmationRows);
    }

    // ローカット機能は、FEQ と SEL の組み合わせごとに（min/F、min/7、max/7、max/F の順）、周波数の小さい方から並べる
    [Fact]
    public void LowCutItemsAreGroupedInTheOrderMinFMin7Max7MaxF() {
        var names = Definition.Items.Where(item => item.Name.StartsWith("ローカット機能")).Select(item => item.Name).ToList();

        Assert.Equal(
            [
                "ローカット機能 FEQ=min SEL=F 0.51Hz b1:0",
                "ローカット機能 FEQ=min SEL=F 1Hz b1:0以外",
                "ローカット機能 FEQ=min SEL=7 3.5Hz b1:0",
                "ローカット機能 FEQ=min SEL=7 5.7Hz b1:0以外",
                "ローカット機能 FEQ=max SEL=7 33Hz b1:0",
                "ローカット機能 FEQ=max SEL=7 46Hz b1:0以外",
                "ローカット機能 FEQ=max SEL=F 5.7Hz b1:0",
                "ローカット機能 FEQ=max SEL=F 8Hz b1:0以外",
            ],
            names);
    }

    [Fact]
    public void TransferPointsToTheSeparateTypeSheetWithTheCheckBoxColumn() {
        var transfer = Definition.Transfer;

        Assert.Equal("分離形　1～5", transfer.DataSheet);
        Assert.Equal(["分離形　1～5", "分離形　6～10", "分離形　11～15", "分離形　16～20"], transfer.DataSheets);
        Assert.Equal(5, transfer.UnitsPerSheet);
        Assert.Equal("C1", transfer.UnitCountCell);
        Assert.Null(transfer.SettingSheet);
        Assert.Equal((7, 48), (transfer.FirstRow, transfer.LastRow));
        Assert.Equal(("B3", "分離形"), (transfer.VerifyCell, transfer.VerifyValue));
        Assert.Equal("H", transfer.FirstUnitColumn);
        Assert.Equal("M", transfer.ConfirmationColumn);
        Assert.Null(transfer.UnitsPerBlock);
    }

    // 1枚のシートに5台ずつ、4枚（1～5、6～10、11～15、16～20）で最大20台
    [Fact]
    public void MaxUnitCountIsTwentyBecauseFourSheetsHoldFiveUnitsEach() {
        Assert.Equal(20, Definition.MaxUnitCount);
    }

    // 合格範囲は、画面に書かれた表記のとおり（ただしノッチの最小周波数は、表記に範囲がないので、テンプレートの判定式の範囲）
    [Theory]
    [InlineData(12, 1.45, 1.55)]
    [InlineData(13, -1.55, -1.45)]
    [InlineData(14, 0, 300)]
    [InlineData(15, 2100, 3500)]
    [InlineData(17, 0, 200)]
    [InlineData(19, 30, 34)]
    [InlineData(20, 40, 44)]
    [InlineData(21, 23, 27)]
    [InlineData(23, 5750, 6750)]
    [InlineData(24, 0, 50)]
    [InlineData(25, 100, 10000)]
    public void RangesMatchTheTemplate(int excelRow, double lower, double upper) {
        var item = ItemAt(excelRow);

        Assert.Equal(((decimal)lower, (decimal)upper), (item.LowerLimit, item.UpperLimit));
    }

    [Fact]
    public void OnlyTheNotchFrequencyHasADefaultValue() {
        var withDefault = Definition.Items.Where(item => item.DefaultValue is not null).ToList();

        var item = Assert.Single(withDefault);
        Assert.Equal(23, item.ExcelRow);
        Assert.Equal(6250m, item.DefaultValue);
    }

    // 本体の設定の指示は、各ブロックの先頭の項目の説明に入れる
    [Theory]
    [InlineData(17, "SEL=0, AMP=max, OSC=0Hz")]
    [InlineData(19, "VTX-0V間=400mVp-p")]
    [InlineData(23, "SEL=0, AMP=min, ATT=30dB")]
    [InlineData(27, "OSC=1000Hz, DIV=1/10")]
    [InlineData(32, "入力200Hz")]
    [InlineData(35, "ATT=50dB, OSC=400Hz")]
    [InlineData(44, "ATT=40dB")]
    public void FirstItemOfEachBlockHasTheSettingInstruction(int excelRow, string expectedText) {
        Assert.Contains(expectedText, ItemAt(excelRow).Description);
    }

    [Fact]
    public void ItemNamesAreUnique() {
        var duplicate = Definition.Items.GroupBy(item => item.Name).FirstOrDefault(group => group.Count() > 1);

        Assert.Null(duplicate);
    }
}
