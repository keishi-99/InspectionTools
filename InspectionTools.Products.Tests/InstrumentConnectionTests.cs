using InspectionTools.Instruments.Tests;

namespace InspectionTools.Products.Tests;

/// <summary>測定器への接続の管理（接続、重複の拒否、失敗時の全解除、解除時の OFF 戻し）を、ダミーの接続で確認する</summary>
public class InstrumentConnectionTests {

    private readonly LoggingConnection _log = new(new FakeInstrumentConnection());
    private readonly InstrumentConnection _connection;

    public InstrumentConnectionTests() {
        _connection = new InstrumentConnection(El5000AcProcedure.Create(), _log);
    }

    private static InstrumentEntry Entry(string category, string name, int signalType) =>
        new(category, name, $"ADDRESS-{name}", signalType);

    // 5つの役割すべてに、別々の測定器を選んだ状態
    private static Dictionary<InstrumentRole, InstrumentEntry> AllSelections() => new() {
        [InstrumentRole.Dmm01] = Entry("デジタルマルチメータ", "DMM-01", 2),
        [InstrumentRole.Dmm02] = Entry("デジタルマルチメータ", "DMM-02", 2),
        [InstrumentRole.Fg] = Entry("ファンクションジェネレータ", "FG-01", 2),
        [InstrumentRole.Dcs] = Entry("電流電圧発生器", "DCS-01", 2),
        [InstrumentRole.Cnt] = Entry("ユニバーサルカウンタ", "CNT-01", 3),
    };

    [Fact]
    public async Task ConnectSendsInitCommandsAndMakesTheControllerAvailable() {
        await _connection.ConnectAsync(AllSelections());

        Assert.True(_connection.IsConnected);
        Assert.NotNull(_connection.Controller);
        Assert.Equal(
            [InstrumentRole.Dmm01, InstrumentRole.Dmm02, InstrumentRole.Fg, InstrumentRole.Dcs, InstrumentRole.Cnt],
            _connection.ConnectedRoles);
        Assert.Equal(["DMM-01", "DMM-02", "FG-01", "DCS-01", "CNT-01"], _log.Sent.Select(sent => sent.Instrument));
    }

    [Fact]
    public async Task ConnectWithOnlySomeRolesSelectedConnectsOnlyThose() {
        var selections = new Dictionary<InstrumentRole, InstrumentEntry> {
            [InstrumentRole.Dcs] = Entry("電流電圧発生器", "DCS-01", 2),
        };

        await _connection.ConnectAsync(selections);

        Assert.Equal([InstrumentRole.Dcs], _connection.ConnectedRoles);
        Assert.False(_connection.Controller!.IsConnected(InstrumentRole.Dmm01));
        Assert.Equal(["DCS-01"], _log.Sent.Select(sent => sent.Instrument));
    }

    [Fact]
    public async Task ConnectWithNothingSelectedIsRejected() {
        await Assert.ThrowsAsync<ArgumentException>(() => _connection.ConnectAsync(new Dictionary<InstrumentRole, InstrumentEntry>()));

        Assert.False(_connection.IsConnected);
    }

    [Fact]
    public async Task ConnectWithTheSameInstrumentForTwoRolesIsRejectedBeforeSendingAnything() {
        var selections = AllSelections();
        selections[InstrumentRole.Dmm02] = selections[InstrumentRole.Dmm01];

        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _connection.ConnectAsync(selections));

        Assert.Contains("DMM-01", exception.Message);
        Assert.Empty(_log.Sent);
        Assert.False(_connection.IsConnected);
    }

    [Fact]
    public async Task ConnectWhileConnectedIsRejected() {
        await _connection.ConnectAsync(AllSelections());

        await Assert.ThrowsAsync<InvalidOperationException>(() => _connection.ConnectAsync(AllSelections()));
    }

    [Fact]
    public async Task FailingToConnectSomeInstrumentsReleasesEverythingAndReportsEveryFailure() {
        _log.ShouldFail = instrument => instrument.Name is "DMM-02" or "FG-01";

        var exception = await Assert.ThrowsAsync<InstrumentConnectionException>(() => _connection.ConnectAsync(AllSelections()));

        Assert.Equal([InstrumentRole.Dmm02, InstrumentRole.Fg], exception.Failures.Select(failure => failure.Role));
        Assert.Contains("DMM-02", exception.Message);
        Assert.Contains("FG-01", exception.Message);
        Assert.False(_connection.IsConnected);
        Assert.Null(_connection.Controller);
        Assert.Empty(_connection.ConnectedRoles);
    }

    [Fact]
    public async Task ConnectCanBeRetriedAfterAFailure() {
        _log.ShouldFail = instrument => instrument.Name == "CNT-01";
        await Assert.ThrowsAsync<InstrumentConnectionException>(() => _connection.ConnectAsync(AllSelections()));
        _log.ShouldFail = _ => false;

        await _connection.ConnectAsync(AllSelections());

        Assert.True(_connection.IsConnected);
    }

    [Fact]
    public async Task DisconnectReleasesWithoutSendingAnythingEvenWhenOutputsAreOn() {
        await _connection.ConnectAsync(AllSelections());
        await _connection.Controller!.ApplyAsync(new ItemAction([new(InstrumentRole.Fg, "ON"), new(InstrumentRole.Dcs, "5V")]));
        _log.Sent.Clear();

        _connection.Disconnect();

        Assert.Empty(_log.Sent);
        Assert.False(_connection.IsConnected);
        Assert.Null(_connection.Controller);
        Assert.Empty(_connection.ConnectedRoles);
    }

    [Fact]
    public void DisconnectWhenNotConnectedDoesNothing() {
        _connection.Disconnect();

        Assert.Empty(_log.Sent);
        Assert.False(_connection.IsConnected);
    }

    [Fact]
    public async Task ConnectAgainAfterDisconnectWorks() {
        await _connection.ConnectAsync(AllSelections());
        _connection.Disconnect();

        await _connection.ConnectAsync(AllSelections());

        Assert.True(_connection.IsConnected);
    }
}
