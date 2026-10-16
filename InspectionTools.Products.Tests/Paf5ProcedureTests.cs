using InspectionTools.Common;
using InspectionTools.Inspection;
using InspectionTools.Instruments.Tests;

namespace InspectionTools.Products.Tests;

/// <summary>PAF5 の手順表が、分離形の定義（products/PAF5-分離形.json）と、旧アプリのコマンドと食い違っていないことを確認する</summary>
public class Paf5ProcedureTests {

    private static readonly ProductDefinition Definition =
        ProductDefinitionLoader.LoadFile(RepositoryFiles.Find("products", "PAF5-分離形.json"));

    private static readonly ProductProcedure Procedure = Paf5Procedure.Create(Definition.Product);

    // PAF5 が使う役割（DCS とカウンタは使わない）
    private static readonly InstrumentRole[] UsedRoles = [InstrumentRole.Dmm01, InstrumentRole.Dmm02, InstrumentRole.Fg, InstrumentRole.Osc];

    private static readonly InstrumentRole[] MeasuringRoles = [InstrumentRole.Dmm01, InstrumentRole.Dmm02, InstrumentRole.Osc];

    // 転記先の行から、その項目の動作 ID を引く
    private static string? ActionIdOfRow(int excelRow) => Definition.Items.Single(item => item.ExcelRow == excelRow).ActionId;

    private static ItemAction ActionOfRow(int excelRow) => Procedure.FindAction(ActionIdOfRow(excelRow));

    private static string? StateOf(int excelRow, InstrumentRole role) =>
        ActionOfRow(excelRow).Settings.SingleOrDefault(setting => setting.Role == role)?.Name;

    // ===== 定義と手順表の整合 =====

    [Fact]
    public void ProcedureIsForTheProductGivenToIt() {
        Assert.Equal("PAF5-分離形", Procedure.Product);
        Assert.Equal("PAF5-一体形2線式", Paf5Procedure.Create("PAF5-一体形2線式").Product);
    }

    [Fact]
    public void EveryActionIdInTheDefinitionExistsInTheProcedure() {
        var unknown = Definition.Items
            .Where(item => item.ActionId is not null && !Procedure.Actions.ContainsKey(item.ActionId))
            .Select(item => $"{item.Name}: {item.ActionId}");

        Assert.Empty(unknown);
    }

    [Fact]
    public void EveryActionInTheProcedureIsUsedByAnItem() {
        var usedIds = Definition.Items.Select(item => item.ActionId).ToHashSet();

        Assert.Empty(Procedure.Actions.Keys.Where(id => !usedIds.Contains(id)));
    }

    [Fact]
    public void EverySettingNamedByAnActionHasACommand() {
        var missing = Procedure.Actions
            .SelectMany(action => action.Value.Settings.Select(setting => (action.Key, setting)))
            .Where(pair => !Procedure.Profiles[pair.setting.Role].States.ContainsKey(pair.setting.Name))
            .Select(pair => $"{pair.Key}: {pair.setting.Role} {pair.setting.Name}");

        Assert.Empty(missing);
    }

    [Fact]
    public void NoActionSetsTheSameRoleTwice() {
        var duplicated = Procedure.Actions
            .Where(action => action.Value.Settings.GroupBy(setting => setting.Role).Any(group => group.Count() > 1))
            .Select(action => action.Key);

        Assert.Empty(duplicated);
    }

    [Fact]
    public void IdleStatesAndInitialStatesHaveCommands() {
        foreach (var (role, state) in Procedure.IdleStates) {
            Assert.True(Procedure.Profiles[role].States.ContainsKey(state), $"{role} の待機の状態 {state} のコマンドがありません。");
        }
        foreach (var role in UsedRoles) {
            Assert.True(Procedure.Profiles[role].States.ContainsKey(Procedure.Profiles[role].InitialState), $"{role} の初期状態のコマンドがありません。");
        }
    }

    [Fact]
    public void OnlyTheFunctionGeneratorIsReturnedToOffWhenAnItemDoesNotUseIt() {
        Assert.Equal([InstrumentRole.Fg], Procedure.IdleStates.Keys);
        Assert.Equal("OFF", Procedure.IdleStates[InstrumentRole.Fg]);
    }

    [Fact]
    public void ProfilesAndRequirementsExistForTheRolesPaf5Uses() {
        Assert.Equal(UsedRoles.Order(), Procedure.Profiles.Keys.Order());
        Assert.Equal(UsedRoles.Order(), Procedure.Requirements.Keys.Order());
        // 通信方式の一覧は配列なので、レコードの等価比較（参照の比較）ではなく、中身を別々に比べる
        void AssertRequirement(InstrumentRole role, string category, int[] signalTypes) {
            Assert.Equal(category, Procedure.Requirements[role].Category);
            Assert.Equal(signalTypes, Procedure.Requirements[role].SignalTypes);
        }

        AssertRequirement(InstrumentRole.Dmm01, "デジタルマルチメータ", [1, 2]);
        AssertRequirement(InstrumentRole.Dmm02, "デジタルマルチメータ", [1, 2]);
        AssertRequirement(InstrumentRole.Fg, "ファンクションジェネレータ", [2]);
        AssertRequirement(InstrumentRole.Osc, "オシロスコープ", [2]);
    }

    [Fact]
    public void OnlyMeasuringRolesAcquireValues() {
        var acquisitionRoles = Procedure.Actions.Values.Where(action => action.Acquisition is not null).Select(action => action.Acquisition!.Role);

        Assert.All(acquisitionRoles, role => Assert.Contains(role, MeasuringRoles));
    }

    [Fact]
    public void ItemsThatAcquireValuesAreExactlyTheMeasurementItems() {
        var acquiring = Definition.Items.Where(item => Procedure.FindAction(item.ActionId).Acquisition is not null).Select(item => item.Name);
        var measurement = Definition.Items.Where(item => item.Kind == ItemKind.Measurement).Select(item => item.Name);

        Assert.Equal(measurement, acquiring);
    }

    // ===== 項目ごとの動作 =====

    [Fact]
    public void CircuitVoltageRowsReadDmm01InVoltageModeInVolts() {
        foreach (var row in new[] { 12, 13 }) {
            Assert.Equal([new InstrumentSetting(InstrumentRole.Dmm01, "DCV")], ActionOfRow(row).Settings);
            Assert.Equal(new Acquisition(InstrumentRole.Dmm01, Scale: 1m, Decimals: 2), ActionOfRow(row).Acquisition);
        }
    }

    [Fact]
    public void CurrentConsumptionRowsReadDmm02InCurrentModeInMicroamperes() {
        foreach (var row in new[] { 14, 15 }) {
            Assert.Equal([new InstrumentSetting(InstrumentRole.Dmm02, "DCI")], ActionOfRow(row).Settings);
            Assert.Equal(new Acquisition(InstrumentRole.Dmm02, Scale: 1000000m, Decimals: 2), ActionOfRow(row).Acquisition);
        }
    }

    [Fact]
    public void OscilloscopeRowsReadMeasurementOneInMillivolts() {
        foreach (var row in new[] { 17, 24, 25 }) {
            Assert.Equal(new Acquisition(InstrumentRole.Osc, Scale: 1000m, Decimals: 3), ActionOfRow(row).Acquisition);
        }
    }

    // 行17〜25のオシロスコープの時間軸（行17=1ms、行19=10ms、行20=1ms、行21〜25=100μs）。行18、22は見出しの行
    [Theory]
    [InlineData(17, "1ms")]
    [InlineData(19, "10ms")]
    [InlineData(20, "1ms")]
    [InlineData(21, "100us")]
    [InlineData(23, "100us")]
    [InlineData(24, "100us")]
    [InlineData(25, "100us")]
    public void OscilloscopeTimeScaleOfEachRow(int excelRow, string expectedState) {
        Assert.Equal(expectedState, StateOf(excelRow, InstrumentRole.Osc));
    }

    [Fact]
    public void RowsOutside17To25DoNotTouchTheOscilloscope() {
        var others = Definition.Items.Where(item => item.ExcelRow is not (null or (>= 17 and <= 25)));

        Assert.All(others, item => Assert.Null(StateOf(item.ExcelRow!.Value, InstrumentRole.Osc)));
    }

    // 項目ごとの周波数発生器の状態。行32、33（スケーラ）は別の機器を使うので、周波数発生器を使わない（自動で OFF に戻る）
    [Theory]
    [InlineData(19, "20Hz")]
    [InlineData(20, "200Hz")]
    [InlineData(21, "2000Hz")]
    [InlineData(23, "6250Hz")]
    [InlineData(24, "6250Hz")]
    [InlineData(25, "5000Hz")]
    [InlineData(27, "1000Hz")]
    [InlineData(28, "1000Hz")]
    [InlineData(29, "1000Hz")]
    [InlineData(30, "1000Hz")]
    [InlineData(35, "400Hz")]
    [InlineData(36, "400Hz")]
    [InlineData(37, "400Hz")]
    [InlineData(38, "400Hz")]
    [InlineData(44, "0.51Hz")]
    [InlineData(45, "1Hz")]
    [InlineData(40, "3.5Hz")]
    [InlineData(41, "5.7Hz")]
    [InlineData(42, "33Hz")]
    [InlineData(43, "46Hz")]
    [InlineData(46, "5.7Hz")]
    [InlineData(47, "8Hz")]
    public void FunctionGeneratorFrequencyOfEachRow(int excelRow, string expectedState) {
        Assert.Equal(expectedState, StateOf(excelRow, InstrumentRole.Fg));
    }

    [Fact]
    public void LowCutRowsFollowTheGroupOrderOfTheDefinition() {
        // 定義の並び（min/F、min/7、max/7、max/F）のとおりに、周波数が小さい方から出る
        var states = new[] { 44, 45, 40, 41, 42, 43, 46, 47 }.Select(row => StateOf(row, InstrumentRole.Fg));

        Assert.Equal(["0.51Hz", "1Hz", "3.5Hz", "5.7Hz", "33Hz", "46Hz", "5.7Hz", "8Hz"], states);
    }

    [Fact]
    public void RowsThatDoNotUseInstrumentsHaveNoActionId() {
        // 行7〜11は目視と手入力、行32、33（スケーラ）は別の機器、行48は目視
        int[] rows = [7, 8, 9, 10, 11, 32, 33, 48];

        Assert.All(rows, row => Assert.Null(ActionIdOfRow(row)));
    }

    // ===== 旧アプリのコマンドとの照合 =====

    [Fact]
    public void EveryCommandStringIsWrittenInTheOldApp() {
        // 旧アプリには、改行を含む文字列（OSC の初期設定）があるので、空白と改行を取り除いて比べる
        static string Compact(string text) => string.Concat(text.Where(c => !char.IsWhiteSpace(c)));
        var oldSource = Compact(File.ReadAllText(RepositoryFiles.Find("InspectionTools", "Product", "PAF5UserControl.xaml.cs")));
        var allCommands = Procedure.Profiles.Values
            .SelectMany(profile => new[] { profile.Init, profile.AfterRead }.Concat(profile.States.Values))
            .OfType<SwitchCommand>();

        var missing = allCommands
            .SelectMany(command => new[] { command.Adc, command.Visa, command.Gpib })
            .Where(text => text.Length > 0 && !oldSource.Contains(Compact(text)))
            .Distinct();

        Assert.Empty(missing);
    }

    // ===== OSC の接続と値の取得 =====

    // 4つの測定器をすべて接続して、行17から順に動作を適用したときに送られるコマンド。
    // 状態が変わった役割にだけ送り、出力する FG は最後に出す(オシロスコープの時間軸を先に合わせる)
    [Fact]
    public async Task ApplyingRows17To27InOrderSendsOnlyTheChangesWithTheOutputLast() {
        var log = new LoggingConnection(new FakeInstrumentConnection());
        var connection = new InstrumentConnection(Procedure, log);
        await connection.ConnectAsync(new Dictionary<InstrumentRole, InstrumentEntry> {
            [InstrumentRole.Dmm01] = new("デジタルマルチメータ", "DMM-01", "ADDRESS-DMM-01", 2),
            [InstrumentRole.Dmm02] = new("デジタルマルチメータ", "DMM-02", "ADDRESS-DMM-02", 2),
            [InstrumentRole.Fg] = new("ファンクションジェネレータ", "FG-01", "ADDRESS-FG-01", 2),
            [InstrumentRole.Osc] = new("オシロスコープ", "OSC-01", "ADDRESS-OSC-01", 2),
        });
        var controller = connection.Controller!;
        log.Sent.Clear();

        var sentByRow = new Dictionary<int, List<string>>();
        foreach (var row in new[] { 17, 19, 20, 21, 23, 24, 25, 27 }) {
            await controller.ApplyAsync(ActionOfRow(row));
            sentByRow[row] = log.Sent.Select(sent => $"{sent.Instrument} {sent.Command}").ToList();
            log.Sent.Clear();
        }

        Assert.Empty(sentByRow[17]);
        Assert.Equal(["OSC-01 :HORIZONTAL:MAIN:SCALE 1.0E-2;*OPC?", "FG-01 :FREQ 20;:OUTPUT ON;*OPC?"], sentByRow[19]);
        Assert.Equal(["OSC-01 :HORIZONTAL:MAIN:SCALE 1.0E-3;*OPC?", "FG-01 :OUTPUT OFF;:FREQ 200;:OUTPUT ON;*OPC?"], sentByRow[20]);
        Assert.Equal(["OSC-01 :HORIZONTAL:MAIN:SCALE 1.0E-4;*OPC?", "FG-01 :OUTPUT OFF;:FREQ 2000;:OUTPUT ON;*OPC?"], sentByRow[21]);
        Assert.Equal(["FG-01 :OUTPUT OFF;:FREQ 6250;:OUTPUT ON;*OPC?"], sentByRow[23]);
        Assert.Empty(sentByRow[24]);
        Assert.Equal(["FG-01 :OUTPUT OFF;:FREQ 5000;:OUTPUT ON;*OPC?"], sentByRow[25]);
        // 行27以降はオシロスコープを使わないので、時間軸はそのまま
        Assert.Equal(["FG-01 :OUTPUT OFF;:FREQ 1000;:OUTPUT ON;*OPC?"], sentByRow[27]);
    }

    [Fact]
    public async Task OscilloscopeCanBeConnectedAndItsMeasurementCanBeRead() {
        var fake = new FakeInstrumentConnection();
        fake.QueueReading("MEASU:MEAS1:VAL?", "0.1234567");
        var log = new LoggingConnection(fake);
        var connection = new InstrumentConnection(Procedure, log);

        await connection.ConnectAsync(new Dictionary<InstrumentRole, InstrumentEntry> {
            [InstrumentRole.Osc] = new("オシロスコープ", "OSC-01", "ADDRESS-OSC-01", 2),
        });
        var controller = connection.Controller!;
        log.Sent.Clear();
        await controller.ApplyAsync(ActionOfRow(17));
        var value = await controller.AcquireAsync(ActionOfRow(17).Acquisition!);

        Assert.Equal([InstrumentRole.Osc], connection.ConnectedRoles);
        // 初期設定のあとの時間軸は 1ms で、行17の状態と同じなので、何も送らない。値の取得で meas1 を問い合わせる
        Assert.Equal([("OSC-01", "MEASU:MEAS1:VAL?")], log.Sent);
        Assert.Equal(123.457m, value);
    }
}
