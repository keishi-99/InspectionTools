using InspectionTools.App.ViewModels;
using InspectionTools.Common;
using InspectionTools.Inspection;
using InspectionTools.Instruments.Tests;
using InspectionTools.Products;

namespace InspectionTools.App.Tests;

/// <summary>検査画面の測定器の操作（項目の移動での自動適用、再適用、値の取得、通信中の操作停止）を、ダミーの接続で確認する</summary>
public sealed class InstrumentOperationTests : IDisposable {

    private const string CatalogXml = """
        <dsInstList>
          <tblInstList><Category>デジタルマルチメータ</Category><Name>DMM-01</Name><VisaAddress>1</VisaAddress><SignalType>2</SignalType></tblInstList>
          <tblInstList><Category>デジタルマルチメータ</Category><Name>DMM-02</Name><VisaAddress>2</VisaAddress><SignalType>2</SignalType></tblInstList>
          <tblInstList><Category>ファンクションジェネレータ</Category><Name>FG-01</Name><VisaAddress>3</VisaAddress><SignalType>2</SignalType></tblInstList>
          <tblInstList><Category>電流電圧発生器</Category><Name>DCS-01</Name><VisaAddress>4</VisaAddress><SignalType>2</SignalType></tblInstList>
          <tblInstList><Category>ユニバーサルカウンタ</Category><Name>CNT-01</Name><VisaAddress>5</VisaAddress><SignalType>3</SignalType></tblInstList>
        </dsInstList>
        """;

    private const string Dcs4mA = "SIR2,SOI+4MA,OPR";
    private const string DcsOff = "SIR2,SOI+0MA,SBY";
    private const string FgOff = ":OUTP OFF;*OPC?";

    // 項目の位置
    private const int VoltageItem = 0;
    private const int DcsTaskItem = 1;
    private const int ManualItem = 2;
    private const int CurrentItem = 3;

    // 画面を開くときに、決めておいた操作（測定器の選択と接続）を行うダミー
    private sealed class StubDialogService : IDialogService {
        public Action<InstrumentSelectionViewModel>? OnShowInstrumentSelection { get; set; }
        public List<string> Errors { get; } = [];

        public bool Confirm(string message) => true;
        public void ShowError(string message) => Errors.Add(message);
        public void ShowInfo(string message) { }
        public void ShowInstrumentSelection(InstrumentSelectionViewModel viewModel) => OnShowInstrumentSelection?.Invoke(viewModel);
        public string? SelectTemplateFilePath() => null;
    }

    // 時刻を自分で進められる時計（DCS の経過時間のテスト用）
    private sealed class ManualTimeProvider : TimeProvider {
        private DateTimeOffset _now = new(2026, 10, 14, 9, 0, 0, TimeSpan.Zero);

        public override DateTimeOffset GetUtcNow() => _now;

        public void Advance(TimeSpan span) => _now += span;
    }

    // 送ったコマンドの記録、送信の失敗、通信の一時停止（通信中の状態を作るため）を再現できるダミー
    private sealed class ControllableConnection : IInstrumentConnection {
        public FakeInstrumentConnection Inner { get; } = new();
        public List<(string Instrument, string Command)> Sent { get; } = [];
        public Func<InstClass, bool> ShouldFail { get; set; } = _ => false;
        public TaskCompletionSource? Gate { get; set; }

        public async Task<string> ConnectAsync(InstClass instClass) {
            Sent.Add((instClass.Name, instClass.InstCommand));
            if (Gate is not null) {
                await Gate.Task;
            }
            if (ShouldFail(instClass)) {
                throw new IOException("通信できません。");
            }
            return await Inner.ConnectAsync(instClass);
        }
    }

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));
    private readonly StubDialogService _dialog = new();
    private readonly ControllableConnection _transport = new();
    private readonly SessionStore _store;
    private readonly InstrumentSelectionViewModel _selection;
    private readonly MainViewModel _viewModel;
    private readonly ProductDefinition _definition;
    private readonly ManualTimeProvider _time = new();

    public InstrumentOperationTests() {
        Directory.CreateDirectory(_directory);
        File.WriteAllText(Path.Combine(_directory, "VisaAddress.xml"), CatalogXml);

        var procedure = El5000AcProcedure.Create();
        _selection = new InstrumentSelectionViewModel(procedure, new InstrumentConnection(procedure, _transport), Path.Combine(_directory, "VisaAddress.xml"));
        _store = new SessionStore(Path.Combine(_directory, "sessions"));
        _definition = new ProductDefinition(
            "EL5000AC",
            30,
            new TransferDefinition("ACdata", "ACsetting", 5, 50, "A3", "EL5000", "F", UnitsPerBlock: null, BlockRowOffset: 0),
            [
                new InspectionItem("電圧の確認", ItemKind.Measurement, excelRow: 9, lowerLimit: 22.8m, upperLimit: 25.2m, actionId: "dmm02-dcv"),
                new InspectionItem("アナログ校正 DCS 4mA", ItemKind.Task, actionId: "dcs-4mA"),
                new InspectionItem("手入力の項目", ItemKind.Input, excelRow: 16, lowerLimit: -50.2m, upperLimit: -49.8m),
                new InspectionItem("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m, actionId: "dmm01-aci"),
            ]);
        _viewModel = new MainViewModel(_definition, _store, _dialog, _selection, _time);
    }

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    // 測定器の選択と接続の画面で、指定した役割に指定した測定器を選んで接続する
    private async Task ConnectAsync(params (InstrumentRole Role, string Name)[] selections) {
        _selection.Refresh();
        foreach (var (role, name) in selections) {
            _selection.Roles.Single(r => r.Role == role).SelectedName = name;
        }
        await _selection.ConnectAsync();
        Assert.True(_selection.IsConnected);
        _transport.Sent.Clear();
    }

    private Task ConnectAllAsync() => ConnectAsync(
        (InstrumentRole.Dmm01, "DMM-01"), (InstrumentRole.Dmm02, "DMM-02"), (InstrumentRole.Fg, "FG-01"), (InstrumentRole.Dcs, "DCS-01"));

    // ===== 項目の移動での自動適用 =====

    [Fact]
    public async Task MovingToAnItemAppliesItsInstrumentState() {
        await ConnectAllAsync();

        await _viewModel.MoveToAsync(DcsTaskItem, 0);

        Assert.Equal([("DCS-01", Dcs4mA)], _transport.Sent);
    }

    [Fact]
    public async Task MovingToAnItemWithoutActionTurnsOutputsOff() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(DcsTaskItem, 0);
        _transport.Sent.Clear();

        await _viewModel.MoveToAsync(ManualItem, 0);

        Assert.Equal([("DCS-01", DcsOff)], _transport.Sent);
    }

    [Fact]
    public async Task MovingToAnotherUnitOfTheSameItemSendsNothing() {
        _viewModel.IncreaseUnitCountCommand.Execute(null);
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(DcsTaskItem, 0);
        _transport.Sent.Clear();

        await _viewModel.MoveToAsync(DcsTaskItem, 1);

        Assert.Empty(_transport.Sent);
    }

    [Fact]
    public async Task ConfirmingAnItemMovesToTheNextAndAppliesIt() {
        await ConnectAllAsync();
        _viewModel.InputText = "24";

        await _viewModel.CommitValueCommand.ExecuteAsync(null);

        Assert.Equal(DcsTaskItem, _viewModel.CurrentItemIndex);
        Assert.Equal([("DCS-01", Dcs4mA)], _transport.Sent);
    }

    [Fact]
    public async Task NextAndPreviousApplyTheInstrumentState() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(VoltageItem, 0);
        _transport.Sent.Clear();

        await _viewModel.MoveNextCommand.ExecuteAsync(null);
        Assert.Equal([("DCS-01", Dcs4mA)], _transport.Sent);
        _transport.Sent.Clear();

        await _viewModel.MovePreviousCommand.ExecuteAsync(null);
        Assert.Equal([("DCS-01", DcsOff)], _transport.Sent);
    }

    [Fact]
    public async Task ConnectingFromTheScreenAppliesTheCurrentItemRightAway() {
        await _viewModel.MoveToAsync(DcsTaskItem, 0);
        Assert.Empty(_transport.Sent);
        _dialog.OnShowInstrumentSelection = selection => {
            selection.Refresh();
            selection.Roles.Single(role => role.Role == InstrumentRole.Dcs).SelectedName = "DCS-01";
            selection.ConnectAsync().GetAwaiter().GetResult();
        };

        await _viewModel.OpenInstrumentSelectionCommand.ExecuteAsync(null);

        Assert.Contains(("DCS-01", Dcs4mA), _transport.Sent);
    }

    [Fact]
    public async Task NothingIsSentWhileNotConnected() {
        await _viewModel.MoveToAsync(DcsTaskItem, 0);
        await _viewModel.MoveNextCommand.ExecuteAsync(null);

        Assert.Empty(_transport.Sent);
        Assert.False(_viewModel.CanReapplyInstrument);
    }

    // ===== 設定を再適用 =====

    [Fact]
    public async Task ReapplySendsTheStateAgainEvenWhenItMatchesTheRecord() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(DcsTaskItem, 0);
        _transport.Sent.Clear();

        await _viewModel.ReapplyInstrumentCommand.ExecuteAsync(null);

        Assert.Equal([("FG-01", FgOff), ("DCS-01", Dcs4mA)], _transport.Sent);
    }

    // ===== 値の取得 =====

    [Fact]
    public async Task AcquiringShowsTheValueInTheInputBoxWithoutConfirmingIt() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(CurrentItem, 0);
        _transport.Inner.QueueReading("FETC?", "0.146745");

        await _viewModel.AcquireValueCommand.ExecuteAsync(null);

        Assert.Equal("146.745", _viewModel.InputText);
        Assert.Equal(ResultState.NotDone, _viewModel.Rows[CurrentItem].Cells[0].State);
        Assert.False(_viewModel.IsInstrumentBusy);
    }

    [Fact]
    public async Task AcquiringAppliesTheItemStateFirstSoAFailedAutomaticApplyIsRecovered() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(DcsTaskItem, 0);
        // 次の項目への移動で、DCS を止める送信が失敗したとして、記録の DCS は 4mA のまま残る
        _transport.ShouldFail = instrument => instrument.Name == "DCS-01";
        await _viewModel.MoveToAsync(CurrentItem, 0);
        _transport.ShouldFail = _ => false;
        _transport.Sent.Clear();

        await _viewModel.AcquireValueCommand.ExecuteAsync(null);

        Assert.Equal([("DCS-01", DcsOff), ("DMM-01", "FETC?")], _transport.Sent);
    }

    [Fact]
    public async Task ConfirmingTheAcquiredValueAsIsRecordsItAsAcquired() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(CurrentItem, 0);
        _transport.Inner.QueueReading("FETC?", "0.146745");
        await _viewModel.AcquireValueCommand.ExecuteAsync(null);

        await _viewModel.CommitValueCommand.ExecuteAsync(null);

        Assert.False(LoadSavedResult(CurrentItem).IsManual);
        Assert.Equal(146.745m, LoadSavedResult(CurrentItem).Value);
    }

    [Fact]
    public async Task ConfirmingAnEditedValueRecordsItAsManual() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(CurrentItem, 0);
        _transport.Inner.QueueReading("FETC?", "0.146745");
        await _viewModel.AcquireValueCommand.ExecuteAsync(null);
        _viewModel.InputText = "146.7";

        await _viewModel.CommitValueCommand.ExecuteAsync(null);

        Assert.True(LoadSavedResult(CurrentItem).IsManual);
    }

    [Fact]
    public async Task TypingAValueWithoutAcquiringRecordsItAsManual() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(CurrentItem, 0);
        _viewModel.InputText = "146.745";

        await _viewModel.CommitValueCommand.ExecuteAsync(null);

        Assert.True(LoadSavedResult(CurrentItem).IsManual);
    }

    [Fact]
    public async Task ConfirmingAnAcquiredValueAgainAfterComingBackKeepsItAcquired() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(CurrentItem, 0);
        _transport.Inner.QueueReading("FETC?", "0.146745");
        await _viewModel.AcquireValueCommand.ExecuteAsync(null);
        await _viewModel.CommitValueCommand.ExecuteAsync(null);

        await _viewModel.MoveToAsync(CurrentItem, 0);
        await _viewModel.CommitValueCommand.ExecuteAsync(null);

        Assert.False(LoadSavedResult(CurrentItem).IsManual);
    }

    [Fact]
    public async Task AcquireButtonIsAvailableOnlyForAcquiringItemsWithTheirInstrumentConnected() {
        await ConnectAsync((InstrumentRole.Dcs, "DCS-01"));

        await _viewModel.MoveToAsync(VoltageItem, 0);
        Assert.True(_viewModel.HasAcquisition);
        Assert.False(_viewModel.CanAcquire);

        await _viewModel.MoveToAsync(ManualItem, 0);
        Assert.False(_viewModel.HasAcquisition);
        Assert.False(_viewModel.CanAcquire);
    }

    [Fact]
    public async Task AcquireButtonBecomesAvailableWhenTheInstrumentIsConnected() {
        await _viewModel.MoveToAsync(VoltageItem, 0);
        Assert.False(_viewModel.CanAcquire);
        var changed = new List<string?>();
        _viewModel.PropertyChanged += (_, e) => changed.Add(e.PropertyName);

        await ConnectAllAsync();

        Assert.True(_viewModel.CanAcquire);
        Assert.Contains(nameof(MainViewModel.CanAcquire), changed);
        Assert.Contains(nameof(MainViewModel.CanReapplyInstrument), changed);
    }

    [Fact]
    public async Task AcquireDoesNothingWhenItCannotBeUsed() {
        await ConnectAsync((InstrumentRole.Dcs, "DCS-01"));
        await _viewModel.MoveToAsync(VoltageItem, 0);
        _transport.Sent.Clear();

        await _viewModel.AcquireValueCommand.ExecuteAsync(null);

        Assert.Empty(_transport.Sent);
        Assert.Equal("", _viewModel.InputText);
    }

    // ===== 通信の失敗 =====

    [Fact]
    public async Task FailingToAcquireShowsAnErrorAndKeepsTheInputBox() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(CurrentItem, 0);
        _viewModel.InputText = "100";
        _transport.ShouldFail = instrument => instrument.Name == "DMM-01";

        await _viewModel.AcquireValueCommand.ExecuteAsync(null);

        Assert.Contains("値を取得できませんでした", Assert.Single(_dialog.Errors));
        Assert.Equal("100", _viewModel.InputText);
        Assert.False(_viewModel.IsInstrumentBusy);
    }

    [Fact]
    public async Task FailingToApplyShowsAnErrorButStillMovesAndCanBeRetried() {
        await ConnectAllAsync();
        _transport.ShouldFail = instrument => instrument.Name == "DCS-01";

        await _viewModel.MoveToAsync(DcsTaskItem, 0);

        Assert.Contains("測定器の設定を適用できませんでした", Assert.Single(_dialog.Errors));
        Assert.Equal(DcsTaskItem, _viewModel.CurrentItemIndex);
        Assert.False(_viewModel.IsInstrumentBusy);

        _transport.ShouldFail = _ => false;
        _transport.Sent.Clear();
        await _viewModel.ReapplyInstrumentCommand.ExecuteAsync(null);

        Assert.Contains(("DCS-01", Dcs4mA), _transport.Sent);
        Assert.Single(_dialog.Errors);
    }

    // ===== 通信中の操作停止 =====

    [Fact]
    public async Task OperationsAreIgnoredWhileTalkingToInstruments() {
        await ConnectAllAsync();
        _transport.Gate = new TaskCompletionSource();

        var moving = _viewModel.MoveToAsync(DcsTaskItem, 0);

        Assert.True(_viewModel.IsInstrumentBusy);
        Assert.False(_viewModel.IsInstrumentIdle);
        Assert.False(_viewModel.CanReapplyInstrument);
        // 通信中の確定や移動は、受け付けない（通信が重なって、コマンドが混ざるのを防ぐ）
        _viewModel.InputText = "1";
        await _viewModel.CommitValueCommand.ExecuteAsync(null);
        await _viewModel.MoveNextCommand.ExecuteAsync(null);
        await _viewModel.MoveToAsync(CurrentItem, 0);
        Assert.Equal(DcsTaskItem, _viewModel.CurrentItemIndex);

        _transport.Gate.SetResult();
        await moving;

        Assert.False(_viewModel.IsInstrumentBusy);
        Assert.Equal([("DCS-01", Dcs4mA)], _transport.Sent);
    }

    // ===== DCS の経過時間 =====

    [Fact]
    public async Task ElapsedTextIsEmptyWhenDcsIsNotConnected() {
        await ConnectAsync((InstrumentRole.Dmm01, "DMM-01"));

        _viewModel.UpdateDcsElapsed();

        Assert.Equal("", _viewModel.DcsElapsedText);
    }

    [Fact]
    public async Task ElapsedTextShowsZeroWhileDcsIsOff() {
        await ConnectAllAsync();

        _time.Advance(TimeSpan.FromMinutes(5));
        _viewModel.UpdateDcsElapsed();

        Assert.Equal("DCS 経過 00:00", _viewModel.DcsElapsedText);
    }

    [Fact]
    public async Task ElapsedTimeCountsFromTheMomentDcsStartsOutputting() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(DcsTaskItem, 0);

        _time.Advance(TimeSpan.FromSeconds(75));
        _viewModel.UpdateDcsElapsed();

        Assert.Equal("DCS 経過 01:15", _viewModel.DcsElapsedText);
    }

    [Fact]
    public async Task ElapsedTimeKeepsCountingPastAnHourInMinutes() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(DcsTaskItem, 0);

        _time.Advance(TimeSpan.FromSeconds(3725));
        _viewModel.UpdateDcsElapsed();

        Assert.Equal("DCS 経過 62:05", _viewModel.DcsElapsedText);
    }

    [Fact]
    public async Task ElapsedTimeResetsWhenDcsIsTurnedOff() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(DcsTaskItem, 0);
        _time.Advance(TimeSpan.FromSeconds(75));

        await _viewModel.MoveToAsync(ManualItem, 0);
        _time.Advance(TimeSpan.FromSeconds(10));
        _viewModel.UpdateDcsElapsed();

        Assert.Equal("DCS 経過 00:00", _viewModel.DcsElapsedText);
    }

    [Fact]
    public async Task ElapsedTimeKeepsCountingWhenTheSameDcsStateStays() {
        _viewModel.IncreaseUnitCountCommand.Execute(null);
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(DcsTaskItem, 0);
        _time.Advance(TimeSpan.FromSeconds(30));

        // 別の台の同じ項目へ移っても、DCS には何も送らないので、数え直さない
        await _viewModel.MoveToAsync(DcsTaskItem, 1);
        _time.Advance(TimeSpan.FromSeconds(30));
        _viewModel.UpdateDcsElapsed();

        Assert.Equal("DCS 経過 01:00", _viewModel.DcsElapsedText);
    }

    [Fact]
    public async Task ReapplyRestartsTheElapsedTimeBecauseTheOutputIsSentAgain() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(DcsTaskItem, 0);
        _time.Advance(TimeSpan.FromSeconds(30));

        await _viewModel.ReapplyInstrumentCommand.ExecuteAsync(null);
        _time.Advance(TimeSpan.FromSeconds(5));
        _viewModel.UpdateDcsElapsed();

        Assert.Equal("DCS 経過 00:05", _viewModel.DcsElapsedText);
    }

    [Fact]
    public async Task ElapsedTextIsClearedWhenTheInstrumentsAreDisconnected() {
        await ConnectAllAsync();
        await _viewModel.MoveToAsync(DcsTaskItem, 0);

        _selection.Disconnect();

        Assert.Equal("", _viewModel.DcsElapsedText);
    }

    // 保存された途中結果から、指定した項目の1台目の結果を読み直す
    private ItemResult LoadSavedResult(int itemIndex) {
        var loaded = _store.Load(_definition.Product, _definition.Items);
        Assert.Equal(LoadStatus.Restored, loaded.Status);
        return loaded.Session!.GetResult(itemIndex, 0);
    }
}
