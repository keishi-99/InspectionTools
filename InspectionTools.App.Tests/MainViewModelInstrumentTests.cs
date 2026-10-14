using InspectionTools.App.ViewModels;
using InspectionTools.Inspection;
using InspectionTools.Instruments.Tests;
using InspectionTools.Products;

namespace InspectionTools.App.Tests;

/// <summary>検査画面の ViewModel と、測定器の選択・接続との関わり（画面を開く、状態の表示、アプリを閉じるときの解除）を確認する</summary>
public sealed class MainViewModelInstrumentTests : IDisposable {

    // 画面を開いた回数と、通知された内容を記録するだけのダミー
    private sealed class StubDialogService : IDialogService {
        public int SelectionShownCount { get; private set; }
        public List<string> Errors { get; } = [];

        public bool Confirm(string message) => true;
        public void ShowError(string message) => Errors.Add(message);
        public void ShowInfo(string message) { }
        public void ShowInstrumentSelection(InstrumentSelectionViewModel viewModel) => SelectionShownCount++;
        public string? SelectTemplateFilePath() => null;
    }

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));
    private readonly StubDialogService _dialog = new();
    private readonly LoggingFake _transport = new();
    private readonly InstrumentSelectionViewModel _selection;
    private readonly MainViewModel _viewModel;

    // 送ったコマンドを記録するダミー
    private sealed class LoggingFake : InspectionTools.Common.IInstrumentConnection {
        private readonly FakeInstrumentConnection _inner = new();
        public List<string> Sent { get; } = [];

        public async Task<string> ConnectAsync(InspectionTools.Common.InstClass instClass) {
            Sent.Add(instClass.InstCommand);
            return await _inner.ConnectAsync(instClass);
        }
    }

    public MainViewModelInstrumentTests() {
        Directory.CreateDirectory(_directory);
        File.WriteAllText(
            Path.Combine(_directory, "VisaAddress.xml"),
            "<dsInstList><tblInstList><Category>電流電圧発生器</Category><Name>DCS-01</Name><VisaAddress>4</VisaAddress><SignalType>2</SignalType></tblInstList></dsInstList>");

        var procedure = El5000AcProcedure.Create();
        _selection = new InstrumentSelectionViewModel(procedure, new InstrumentConnection(procedure, _transport), Path.Combine(_directory, "VisaAddress.xml"));
        _viewModel = new MainViewModel(CreateDefinition(), new SessionStore(Path.Combine(_directory, "sessions")), _dialog, _selection);
    }

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private static ProductDefinition CreateDefinition() => new(
        "EL5000AC",
        30,
        new TransferDefinition("ACdata", "ACsetting", 5, 50, "A3", "EL5000", "F", UnitsPerBlock: null, BlockRowOffset: 0),
        [new InspectionItem("調整作業", ItemKind.Task)]);

    private async Task ConnectDcsAsync() {
        _selection.Refresh();
        _selection.Roles.Single(role => role.Role == InstrumentRole.Dcs).SelectedName = "DCS-01";
        await _selection.ConnectAsync();
    }

    [Fact]
    public void WithoutInstrumentSelectionTheInstrumentOperationsAreHidden() {
        var viewModel = new MainViewModel(CreateDefinition(), new SessionStore(Path.Combine(_directory, "sessions")), _dialog);

        Assert.False(viewModel.HasInstrumentSelection);
        Assert.False(viewModel.IsInstrumentConnected);
        Assert.Equal("", viewModel.InstrumentStatusText);

        viewModel.OpenInstrumentSelectionCommand.Execute(null);
        Assert.Equal(0, _dialog.SelectionShownCount);
    }

    [Fact]
    public void WithoutInstrumentSelectionShutdownDoesNothing() {
        var viewModel = new MainViewModel(CreateDefinition(), new SessionStore(Path.Combine(_directory, "sessions")), _dialog);

        viewModel.Shutdown();

        Assert.Empty(_dialog.Errors);
    }

    [Fact]
    public void OpeningTheSelectionLoadsTheCatalogThenShowsTheScreen() {
        _viewModel.OpenInstrumentSelectionCommand.Execute(null);

        Assert.Equal(1, _dialog.SelectionShownCount);
        Assert.Contains("DCS-01", _selection.Roles.Single(role => role.Role == InstrumentRole.Dcs).CandidateNames);
    }

    [Fact]
    public async Task ConnectionStatusIsShownOnTheInspectionScreen() {
        var changed = new List<string?>();
        _viewModel.PropertyChanged += (_, e) => changed.Add(e.PropertyName);
        Assert.Equal("測定器: 未接続", _viewModel.InstrumentStatusText);

        await ConnectDcsAsync();

        Assert.Equal("測定器: 接続済み（1台）", _viewModel.InstrumentStatusText);
        Assert.True(_viewModel.IsInstrumentConnected);
        Assert.Contains(nameof(MainViewModel.InstrumentStatusText), changed);
    }

    [Fact]
    public async Task ShutdownWhileConnectedDisconnectsWithoutSendingAnythingToTheInstruments() {
        await ConnectDcsAsync();
        await _selection.Controller!.ApplyAsync(El5000AcProcedure.Create().FindAction("dcs-5V"));
        _transport.Sent.Clear();

        _viewModel.Shutdown();

        Assert.Empty(_transport.Sent);
        Assert.False(_viewModel.IsInstrumentConnected);
        Assert.Empty(_dialog.Errors);
    }

    [Fact]
    public void ShutdownWhileNotConnectedDoesNothing() {
        _viewModel.Shutdown();

        Assert.Empty(_transport.Sent);
        Assert.Empty(_dialog.Errors);
    }
}
