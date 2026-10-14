using InspectionTools.App.ViewModels;
using InspectionTools.Common;
using InspectionTools.Instruments.Tests;
using InspectionTools.Products;

namespace InspectionTools.App.Tests;

/// <summary>測定器の選択と接続の画面の ViewModel（候補の読み込み、接続、失敗の表示、解除）を、ダミーの接続で確認する</summary>
public sealed class InstrumentSelectionViewModelTests : IDisposable {

    private const string CatalogXml = """
        <?xml version="1.0" standalone="yes"?>
        <dsInstList xmlns="http://tempuri.org/dsInstList.xsd">
          <tblInstList><Category>デジタルマルチメータ</Category><Name>DMM-01</Name><VisaAddress>1</VisaAddress><SignalType>2</SignalType></tblInstList>
          <tblInstList><Category>デジタルマルチメータ</Category><Name>DMM-02</Name><VisaAddress>2</VisaAddress><SignalType>2</SignalType></tblInstList>
          <tblInstList><Category>ファンクションジェネレータ</Category><Name>FG-01</Name><VisaAddress>3</VisaAddress><SignalType>2</SignalType></tblInstList>
          <tblInstList><Category>電流電圧発生器</Category><Name>DCS-01</Name><VisaAddress>4</VisaAddress><SignalType>2</SignalType></tblInstList>
          <tblInstList><Category>ユニバーサルカウンタ</Category><Name>CNT-01</Name><VisaAddress>5</VisaAddress><SignalType>3</SignalType></tblInstList>
        </dsInstList>
        """;

    // ダミーの接続に処理を渡しながら、送ったコマンドを記録し、指定した測定器への送信だけを失敗させる
    private sealed class RecordingConnection(IInstrumentConnection inner) : IInstrumentConnection {
        public List<(string Instrument, string Command)> Sent { get; } = [];
        public Func<InstClass, bool> ShouldFail { get; set; } = _ => false;

        public async Task<string> ConnectAsync(InstClass instClass) {
            Sent.Add((instClass.Name, instClass.InstCommand));
            if (ShouldFail(instClass)) {
                throw new IOException("通信できません。");
            }
            return await inner.ConnectAsync(instClass);
        }
    }

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));
    private readonly RecordingConnection _transport = new(new FakeInstrumentConnection());
    private readonly InstrumentSelectionViewModel _viewModel;

    public InstrumentSelectionViewModelTests() {
        Directory.CreateDirectory(_directory);
        File.WriteAllText(CatalogPath, CatalogXml);
        var procedure = El5000AcProcedure.Create();
        _viewModel = new InstrumentSelectionViewModel(procedure, new InstrumentConnection(procedure, _transport), CatalogPath);
    }

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private string CatalogPath => Path.Combine(_directory, "VisaAddress.xml");

    private RoleSelectionViewModel Role(InstrumentRole role) => _viewModel.Roles.Single(r => r.Role == role);

    // 役割に測定器を選ぶ
    private void Select(InstrumentRole role, string name) => Role(role).SelectedName = name;

    // 5つの役割すべてに別々の測定器を選ぶ
    private void SelectAll() {
        Select(InstrumentRole.Dmm01, "DMM-01");
        Select(InstrumentRole.Dmm02, "DMM-02");
        Select(InstrumentRole.Fg, "FG-01");
        Select(InstrumentRole.Dcs, "DCS-01");
        Select(InstrumentRole.Cnt, "CNT-01");
    }

    // ===== 候補の読み込み =====

    [Fact]
    public void RefreshFillsCandidatesForEachRoleFromTheCatalog() {
        _viewModel.Refresh();

        Assert.Equal([RoleSelectionViewModel.NoneName, "DMM-01", "DMM-02"], Role(InstrumentRole.Dmm01).CandidateNames);
        Assert.Equal([RoleSelectionViewModel.NoneName, "DMM-01", "DMM-02"], Role(InstrumentRole.Dmm02).CandidateNames);
        Assert.Equal([RoleSelectionViewModel.NoneName, "FG-01"], Role(InstrumentRole.Fg).CandidateNames);
        Assert.Equal([RoleSelectionViewModel.NoneName, "DCS-01"], Role(InstrumentRole.Dcs).CandidateNames);
        Assert.Equal([RoleSelectionViewModel.NoneName, "CNT-01"], Role(InstrumentRole.Cnt).CandidateNames);
        Assert.Equal("", _viewModel.CatalogMessage);
    }

    [Fact]
    public void NothingIsSelectedAtFirst() {
        _viewModel.Refresh();

        Assert.All(_viewModel.Roles, role => Assert.Null(role.SelectedEntry));
    }

    [Fact]
    public void RefreshWithMissingFileShowsTheReasonAndOnlyTheNoneChoice() {
        File.Delete(CatalogPath);

        _viewModel.Refresh();

        Assert.Contains("VisaAddress.xml", _viewModel.CatalogMessage);
        Assert.All(_viewModel.Roles, role => Assert.Equal([RoleSelectionViewModel.NoneName], role.CandidateNames));
    }

    [Fact]
    public void RefreshShowsWarningsForSkippedEntries() {
        File.WriteAllText(CatalogPath, CatalogXml.Replace("<SignalType>3</SignalType>", "<SignalType>9</SignalType>"));

        _viewModel.Refresh();

        Assert.Contains("CNT-01", _viewModel.CatalogMessage);
        Assert.Equal([RoleSelectionViewModel.NoneName], Role(InstrumentRole.Cnt).CandidateNames);
        Assert.Contains("DMM-01", Role(InstrumentRole.Dmm01).CandidateNames);
    }

    [Fact]
    public void RefreshKeepsTheSelectionWhenTheInstrumentIsStillInTheCatalog() {
        _viewModel.Refresh();
        Select(InstrumentRole.Dcs, "DCS-01");

        _viewModel.Refresh();

        Assert.Equal("DCS-01", Role(InstrumentRole.Dcs).SelectedEntry?.Name);
    }

    [Fact]
    public void RefreshResetsTheSelectionWhenTheInstrumentDisappearedFromTheCatalog() {
        _viewModel.Refresh();
        Select(InstrumentRole.Dcs, "DCS-01");
        File.WriteAllText(CatalogPath, CatalogXml.Replace("DCS-01", "DCS-99"));

        _viewModel.Refresh();

        Assert.Null(Role(InstrumentRole.Dcs).SelectedEntry);
    }

    // ===== 接続 =====

    [Fact]
    public async Task ConnectShowsResultsAndLocksTheSelection() {
        _viewModel.Refresh();
        Select(InstrumentRole.Dmm01, "DMM-01");
        Select(InstrumentRole.Dcs, "DCS-01");

        await _viewModel.ConnectAsync();

        Assert.True(_viewModel.IsConnected);
        Assert.NotNull(_viewModel.Controller);
        Assert.Equal("測定器: 接続済み（2台）", _viewModel.StatusText);
        Assert.Equal("接続しました", Role(InstrumentRole.Dmm01).ResultText);
        Assert.Equal("使いません", Role(InstrumentRole.Fg).ResultText);
        Assert.False(_viewModel.CanConnect);
        Assert.True(_viewModel.CanDisconnect);
        Assert.False(_viewModel.CanEditSelection);
        Assert.False(_viewModel.IsBusy);
    }

    [Fact]
    public async Task ConnectWithNothingSelectedAsksToChooseAndSendsNothing() {
        _viewModel.Refresh();

        await _viewModel.ConnectAsync();

        Assert.False(_viewModel.IsConnected);
        Assert.Contains("選んでください", _viewModel.Message);
        Assert.Empty(_transport.Sent);
    }

    [Fact]
    public async Task ConnectWithTheSameInstrumentForTwoRolesShowsAMessageAndSendsNothing() {
        _viewModel.Refresh();
        Select(InstrumentRole.Dmm01, "DMM-01");
        Select(InstrumentRole.Dmm02, "DMM-01");

        await _viewModel.ConnectAsync();

        Assert.False(_viewModel.IsConnected);
        Assert.Contains("DMM-01", _viewModel.Message);
        Assert.Empty(_transport.Sent);
    }

    [Fact]
    public async Task FailedConnectionReleasesEverythingAndShowsTheReasonForEachFailedRole() {
        _viewModel.Refresh();
        SelectAll();
        _transport.ShouldFail = instrument => instrument.Name is "DMM-02" or "CNT-01";

        await _viewModel.ConnectAsync();

        Assert.False(_viewModel.IsConnected);
        Assert.Equal("測定器: 未接続", _viewModel.StatusText);
        Assert.Contains("通信に失敗", Role(InstrumentRole.Dmm02).ResultText);
        Assert.Contains("通信に失敗", Role(InstrumentRole.Cnt).ResultText);
        Assert.Contains("ほかの測定器の失敗", Role(InstrumentRole.Dmm01).ResultText);
        Assert.Contains("すべて解除", _viewModel.Message);
        Assert.True(_viewModel.CanConnect);
        Assert.True(_viewModel.CanEditSelection);
    }

    [Fact]
    public async Task ConnectCanBeRetriedAfterAFailure() {
        _viewModel.Refresh();
        SelectAll();
        _transport.ShouldFail = instrument => instrument.Name == "FG-01";
        await _viewModel.ConnectAsync();
        _transport.ShouldFail = _ => false;

        await _viewModel.ConnectAsync();

        Assert.True(_viewModel.IsConnected);
        Assert.Equal("測定器: 接続済み（5台）", _viewModel.StatusText);
    }

    [Fact]
    public async Task ConnectRaisesPropertyChangedForTheStatusText() {
        _viewModel.Refresh();
        Select(InstrumentRole.Dcs, "DCS-01");
        var changed = new List<string?>();
        _viewModel.PropertyChanged += (_, e) => changed.Add(e.PropertyName);

        await _viewModel.ConnectAsync();

        Assert.Contains(nameof(InstrumentSelectionViewModel.StatusText), changed);
        Assert.Contains(nameof(InstrumentSelectionViewModel.IsConnected), changed);
    }

    [Fact]
    public async Task RefreshWhileConnectedDoesNotChangeAnything() {
        _viewModel.Refresh();
        Select(InstrumentRole.Dcs, "DCS-01");
        await _viewModel.ConnectAsync();
        File.Delete(CatalogPath);

        _viewModel.Refresh();

        Assert.Equal("", _viewModel.CatalogMessage);
        Assert.Equal("DCS-01", Role(InstrumentRole.Dcs).SelectedEntry?.Name);
    }

    // ===== 解除 =====

    [Fact]
    public async Task DisconnectSendsNothingClearsResultsAndUnlocksTheSelection() {
        _viewModel.Refresh();
        SelectAll();
        await _viewModel.ConnectAsync();
        await _viewModel.Controller!.ApplyAsync(FindAction("dcs-5V"));
        _transport.Sent.Clear();

        _viewModel.Disconnect();

        // 出力が ON のままでも、解除のときに測定器へは何も送らない
        Assert.Empty(_transport.Sent);
        Assert.False(_viewModel.IsConnected);
        Assert.Equal("測定器: 未接続", _viewModel.StatusText);
        Assert.Equal("", Role(InstrumentRole.Dcs).ResultText);
        Assert.True(_viewModel.CanEditSelection);
        Assert.True(_viewModel.CanConnect);
    }

    [Fact]
    public void DisconnectWhileNotConnectedDoesNothing() {
        _viewModel.Disconnect();

        Assert.Empty(_transport.Sent);
        Assert.Equal("", _viewModel.Message);
    }

    // 手順表の動作 ID から、動作を引く
    private static ItemAction FindAction(string actionId) => El5000AcProcedure.Create().FindAction(actionId);
}
