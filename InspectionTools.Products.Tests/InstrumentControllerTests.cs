using InspectionTools.Common;
using InspectionTools.Instruments.Tests;

namespace InspectionTools.Products.Tests;

/// <summary>測定器への状態の適用（差分だけを送る、使わない出力は OFF に戻す）と、値の取得を、ダミーの接続で確認する</summary>
public class InstrumentControllerTests {

    private const string Dmm01Init = "*RST;:INIT:CONT 1;:CONF:CURR:AC;*OPC?";
    private const string Dmm02Dcv = "*RST;:INIT:CONT 1;:VOLT:DC:RANG:AUTO ON;*OPC?";
    private const string Dmm02Dci = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?";
    private const string FgInit = "*RST;:FUNC SQU;:FREQ 7E3;:VOLT 3.0VPP;:VOLT:OFFS 1.5;*OPC?";
    private const string DcsInit = "SIR3,SOI+0,SBY";
    private const string CntInit = "*RST;:FUNC PWID;:INPA:COUP DC;:INPA:LPF ON;:INPA:SLOP NEG;:FRUN ON;";

    private static readonly InstrumentRole[] AllRoles = [InstrumentRole.Dmm01, InstrumentRole.Dmm02, InstrumentRole.Fg, InstrumentRole.Dcs, InstrumentRole.Cnt];

    private readonly FakeInstrumentConnection _fake = new();
    private readonly LoggingConnection _log;
    private readonly ProductProcedure _procedure = El5000AcProcedure.Create();

    public InstrumentControllerTests() {
        _log = new LoggingConnection(_fake);
    }

    // 役割ごとの測定器を作る（CNT だけ GPIB 相当の通信方式 3、ほかは VISA の 2）
    private static Dictionary<InstrumentRole, InstClass> CreateInstruments(params InstrumentRole[] roles) {
        var all = new Dictionary<InstrumentRole, InstClass> {
            [InstrumentRole.Dmm01] = new DmmInstClass { Name = "DMM01", SignalType = 2 },
            [InstrumentRole.Dmm02] = new DmmInstClass { Name = "DMM02", SignalType = 2 },
            [InstrumentRole.Fg] = new FgInstClass { Name = "FG", SignalType = 2 },
            [InstrumentRole.Dcs] = new DcsInstClass { Name = "DCS", SignalType = 2 },
            [InstrumentRole.Cnt] = new CntInstClass { Name = "CNT", SignalType = 3 },
        };
        return roles.ToDictionary(role => role, role => all[role]);
    }

    private InstrumentController CreateController(params InstrumentRole[] roles) =>
        new(_procedure, CreateInstruments(roles.Length == 0 ? AllRoles : roles), _log);

    // 初期設定を済ませたコントローラーを作り、ここまでの送信の記録は消す
    private async Task<InstrumentController> CreateInitializedControllerAsync() {
        var controller = CreateController();
        await controller.InitializeAsync();
        _log.Sent.Clear();
        return controller;
    }

    private ItemAction Action(string actionId) => _procedure.FindAction(actionId);

    [Fact]
    public async Task InitializeSendsInitCommandsToEveryConnectedInstrumentInRoleOrder() {
        var controller = CreateController();

        await controller.InitializeAsync();

        Assert.Equal(
            [("DMM01", Dmm01Init), ("DMM02", Dmm02Dcv), ("FG", FgInit), ("DCS", DcsInit), ("CNT", CntInit)],
            _log.Sent);
        Assert.Equal("ACI", controller.States[InstrumentRole.Dmm01]);
        Assert.Equal("DCV", controller.States[InstrumentRole.Dmm02]);
        Assert.Equal("OFF", controller.States[InstrumentRole.Fg]);
        Assert.Equal("OFF", controller.States[InstrumentRole.Dcs]);
    }

    [Fact]
    public async Task InitializeSkipsRolesWithoutInstrument() {
        var controller = CreateController(InstrumentRole.Dcs);

        await controller.InitializeAsync();

        Assert.Equal([("DCS", DcsInit)], _log.Sent);
    }

    [Fact]
    public async Task ApplyingDcsActionSendsOnlyTheDcsCommand() {
        var controller = await CreateInitializedControllerAsync();

        await controller.ApplyAsync(Action("dcs-4mA"));

        Assert.Equal([("DCS", "SIR2,SOI+4MA,OPR")], _log.Sent);
        Assert.Equal("4mA", controller.States[InstrumentRole.Dcs]);
    }

    [Fact]
    public async Task ApplyingTheSameActionAgainSendsNothing() {
        var controller = await CreateInitializedControllerAsync();
        await controller.ApplyAsync(Action("dcs-4mA"));
        _log.Sent.Clear();

        await controller.ApplyAsync(Action("dcs-4mA"));

        Assert.Empty(_log.Sent);
    }

    [Fact]
    public async Task DcsStepsAreSentInTheOrderOfTheActions() {
        var controller = await CreateInitializedControllerAsync();

        foreach (var state in new[] { "4mA", "12mA", "20mA", "1V", "3V", "5V" }) {
            await controller.ApplyAsync(Action($"dcs-{state}"));
        }

        Assert.Equal(
            ["SIR2,SOI+4MA,OPR", "SIR2,SOI+12MA,OPR", "SIR2,SOI+20MA,OPR", "SVR5,SOV+1V,OPR", "SVR5,SOV+3V,OPR", "SVR5,SOV+5V,OPR"],
            _log.Sent.Select(sent => sent.Command));
    }

    [Fact]
    public async Task ActionWithoutDcsTurnsDcsOffAutomatically() {
        var controller = await CreateInitializedControllerAsync();
        await controller.ApplyAsync(Action("dcs-5V"));
        _log.Sent.Clear();

        await controller.ApplyAsync(ItemAction.None);

        Assert.Equal([("DCS", "SIR2,SOI+0MA,SBY")], _log.Sent);
        Assert.Equal("OFF", controller.States[InstrumentRole.Dcs]);
    }

    [Fact]
    public async Task ActionWithoutFgTurnsFgOffAutomatically() {
        var controller = await CreateInitializedControllerAsync();
        await controller.ApplyAsync(Action("fg-on"));
        Assert.Equal([("FG", ":OUTP ON;*OPC?")], _log.Sent);
        _log.Sent.Clear();

        await controller.ApplyAsync(Action("dmm02-dcv"));

        Assert.Equal([("FG", ":OUTP OFF;*OPC?")], _log.Sent);
    }

    [Fact]
    public async Task MeasuringInstrumentsAreNotTouchedByActionsThatDoNotMentionThem() {
        var controller = await CreateInitializedControllerAsync();
        await controller.ApplyAsync(Action("dmm02-dci"));
        Assert.Equal([("DMM02", Dmm02Dci)], _log.Sent);
        _log.Sent.Clear();

        await controller.ApplyAsync(Action("dmm01-aci"));
        await controller.ApplyAsync(ItemAction.None);

        Assert.Empty(_log.Sent);
        Assert.Equal("DCI", controller.States[InstrumentRole.Dmm02]);
    }

    [Fact]
    public async Task DmmModeIsSwitchedWhenTheActionAsksForAnotherMode() {
        var controller = await CreateInitializedControllerAsync();

        await controller.ApplyAsync(Action("dmm02-dci"));
        await controller.ApplyAsync(Action("dmm02-dcv"));

        Assert.Equal([("DMM02", Dmm02Dci), ("DMM02", Dmm02Dcv)], _log.Sent);
    }

    [Fact]
    public async Task ApplyingBeforeInitializeSendsEveryTargetBecauseTheStateIsUnknown() {
        var controller = CreateController();

        await controller.ApplyAsync(ItemAction.None);

        Assert.Equal([("FG", ":OUTP OFF;*OPC?"), ("DCS", "SIR2,SOI+0MA,SBY")], _log.Sent);
    }

    [Fact]
    public async Task RolesWithoutInstrumentAreSkippedAndReportedAsMissing() {
        var controller = CreateController(InstrumentRole.Dmm01);
        await controller.InitializeAsync();
        _log.Sent.Clear();

        await controller.ApplyAsync(Action("dcs-4mA"));

        Assert.Empty(_log.Sent);
        Assert.Equal([InstrumentRole.Dcs], controller.FindMissingRoles(Action("dcs-4mA")));
        Assert.Equal([InstrumentRole.Dmm02], controller.FindMissingRoles(Action("dmm02-dcv")));
        Assert.Equal([InstrumentRole.Cnt], controller.FindMissingRoles(Action("cnt-pulse-width")));
        Assert.Empty(controller.FindMissingRoles(Action("dmm01-aci")));
        Assert.Empty(controller.FindMissingRoles(ItemAction.None));
    }

    [Fact]
    public async Task FailingToSendKeepsThePreviousStateAndReportsTheRole() {
        var controller = await CreateInitializedControllerAsync();
        _log.ShouldFail = instrument => instrument.Name == "DCS";

        var exception = await Assert.ThrowsAsync<InstrumentCommunicationException>(() => controller.ApplyAsync(Action("dcs-4mA")));

        Assert.Equal(InstrumentRole.Dcs, exception.Role);
        Assert.Contains("DCS", exception.Message);
        Assert.Equal("OFF", controller.States[InstrumentRole.Dcs]);
    }

    [Fact]
    public async Task InstrumentWithoutCommandForItsCommunicationMethodIsReported() {
        // DCS の手順表には、VISA（2）と GPIB（3）のコマンドしかない。ADC（1）では送れない
        var instruments = new Dictionary<InstrumentRole, InstClass> { [InstrumentRole.Dcs] = new DcsInstClass { Name = "DCS", SignalType = 1 } };
        var controller = new InstrumentController(_procedure, instruments, _log);

        var exception = await Assert.ThrowsAsync<InstrumentCommunicationException>(controller.InitializeAsync);

        Assert.Equal(InstrumentRole.Dcs, exception.Role);
        Assert.Empty(_log.Sent);
    }

    // ===== 複数の役割が同時に変わるときの送る順序（出力を止める → 測定器の設定 → 出力を出す） =====

    private const string DcsOff = "SIR2,SOI+0MA,SBY";
    private const string FgOn = ":OUTP ON;*OPC?";
    private const string FgOff = ":OUTP OFF;*OPC?";
    private const string CounterGate = "GATE TEST";

    // 出力を出す機器（Fg、Dcs）より役割の順があとの測定器（Cnt）に状態を持たせた手順表で、コントローラーを作る。
    // EL5000AC の手順表では Cnt に状態がなく、「出力を出す送信は最後」と「役割の順」を区別できないため
    private async Task<InstrumentController> CreateControllerWithCounterStateAsync() {
        var profiles = new Dictionary<InstrumentRole, InstrumentProfile>(_procedure.Profiles);
        profiles[InstrumentRole.Cnt] = profiles[InstrumentRole.Cnt] with {
            States = new Dictionary<string, SwitchCommand> { ["GATE"] = new() { Gpib = CounterGate } },
        };
        var controller = new InstrumentController(_procedure with { Profiles = profiles }, CreateInstruments(AllRoles), _log);
        await controller.InitializeAsync();
        _log.Sent.Clear();
        return controller;
    }

    private static ItemAction CounterGateWith(params InstrumentSetting[] outputs) =>
        new([.. outputs, new InstrumentSetting(InstrumentRole.Cnt, "GATE")]);

    [Fact]
    public async Task OutputIsTurnedOffBeforeTheMeasuringInstrumentModeIsSwitched() {
        // アナログ入力の最後（行29、DCS=5V）から、アナログ出力の最初（行30、DMM02=DCI）へ移る場面
        var controller = await CreateInitializedControllerAsync();
        await controller.ApplyAsync(Action("dcs-5V"));
        _log.Sent.Clear();

        await controller.ApplyAsync(Action("dmm02-dci"));

        Assert.Equal([("DCS", DcsOff), ("DMM02", Dmm02Dci)], _log.Sent);
    }

    [Fact]
    public async Task OutputIsTurnedOnAfterTheMeasuringInstrumentIsSet() {
        var controller = await CreateControllerWithCounterStateAsync();

        await controller.ApplyAsync(CounterGateWith(new InstrumentSetting(InstrumentRole.Fg, "ON")));

        // 役割の順なら Fg が先になるが、出力を出す送信は最後にする
        Assert.Equal([("CNT", CounterGate), ("FG", FgOn)], _log.Sent);
    }

    [Fact]
    public async Task OffComesFirstAndOnComesLastWithMeasuringInstrumentInTheMiddle() {
        var controller = await CreateControllerWithCounterStateAsync();
        await controller.ApplyAsync(Action("dcs-5V"));
        _log.Sent.Clear();

        await controller.ApplyAsync(CounterGateWith(new InstrumentSetting(InstrumentRole.Fg, "ON")));

        Assert.Equal([("DCS", DcsOff), ("CNT", CounterGate), ("FG", FgOn)], _log.Sent);
    }

    [Fact]
    public async Task SeveralOffsKeepTheRoleOrder() {
        var controller = await CreateInitializedControllerAsync();
        await controller.ApplyAsync(new ItemAction([new(InstrumentRole.Fg, "ON"), new(InstrumentRole.Dcs, "5V")]));
        _log.Sent.Clear();

        await controller.ApplyAsync(ItemAction.None);

        Assert.Equal([("FG", FgOff), ("DCS", DcsOff)], _log.Sent);
    }

    [Fact]
    public async Task SeveralOnsKeepTheRoleOrder() {
        var controller = await CreateInitializedControllerAsync();

        await controller.ApplyAsync(new ItemAction([new(InstrumentRole.Dcs, "4mA"), new(InstrumentRole.Fg, "ON")]));

        Assert.Equal([("FG", FgOn), ("DCS", "SIR2,SOI+4MA,OPR")], _log.Sent);
    }

    [Fact]
    public async Task ChangingOneOutputToAnotherOutputStateCountsAsTurningOnNotOff() {
        // DCS が 4mA から 12mA に変わるのは、止める変更ではなく、出力を出す変更（測定器の設定のあとに送る）
        var controller = await CreateControllerWithCounterStateAsync();
        await controller.ApplyAsync(Action("dcs-4mA"));
        _log.Sent.Clear();

        await controller.ApplyAsync(CounterGateWith(new InstrumentSetting(InstrumentRole.Dcs, "12mA")));

        Assert.Equal([("CNT", CounterGate), ("DCS", "SIR2,SOI+12MA,OPR")], _log.Sent);
    }

    // ===== 値の取得 =====

    [Fact]
    public async Task AcquireFromDmm01ConvertsAmpereToMilliampereWithThreeDecimals() {
        var controller = CreateController();
        _fake.QueueReading("FETC?", "0.146745");

        var value = await controller.AcquireAsync(Action("dmm01-aci").Acquisition!);

        Assert.Equal(146.745m, value);
        Assert.Equal([("DMM01", "FETC?")], _log.Sent);
    }

    [Fact]
    public async Task AcquireFromDmm02InCurrentModeConvertsToMilliampere() {
        var controller = CreateController();
        _fake.QueueReading("FETC?", "0.0040012");

        var value = await controller.AcquireAsync(Action("dmm02-dci").Acquisition!);

        Assert.Equal(4.0012m, value);
    }

    [Fact]
    public async Task AcquireRoundsAwayFromZeroLikeTheOldAppFormatting() {
        var controller = CreateController();
        _fake.QueueReading("FETC?", "1.23455");

        var value = await controller.AcquireAsync(Action("dmm02-dcv").Acquisition!);

        Assert.Equal(1.2346m, value);
    }

    [Fact]
    public async Task AcquireFromCounterConvertsToMillisecondsAndRestartsMeasurement() {
        var controller = CreateController();
        _fake.QueueReading(":MEAS?XNOW", "+1.234E-03");

        var value = await controller.AcquireAsync(Action("cnt-pulse-width").Acquisition!);

        Assert.Equal(1.234m, value);
        Assert.Equal([("CNT", ":MEAS?XNOW"), ("CNT", ":FRUN ON")], _log.Sent);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("")]
    public async Task AcquireWithUnparsableReadingFailsInsteadOfReturningZero(string reading) {
        var controller = CreateController();
        _fake.QueueReading("FETC?", reading);

        var exception = await Assert.ThrowsAsync<InstrumentCommunicationException>(() => controller.AcquireAsync(Action("dmm01-aci").Acquisition!));

        Assert.Equal(InstrumentRole.Dmm01, exception.Role);
        Assert.IsType<FormatException>(exception.InnerException);
    }

    [Fact]
    public async Task AcquireFromRoleWithoutInstrumentFails() {
        var controller = CreateController(InstrumentRole.Dmm01);

        var exception = await Assert.ThrowsAsync<InstrumentCommunicationException>(() => controller.AcquireAsync(Action("cnt-pulse-width").Acquisition!));

        Assert.Equal(InstrumentRole.Cnt, exception.Role);
    }

    [Fact]
    public async Task AcquireWithCommunicationErrorReportsTheRole() {
        var controller = CreateController();
        _log.ShouldFail = instrument => instrument.Name == "DMM02";

        var exception = await Assert.ThrowsAsync<InstrumentCommunicationException>(() => controller.AcquireAsync(Action("dmm02-dcv").Acquisition!));

        Assert.Equal(InstrumentRole.Dmm02, exception.Role);
    }
}
