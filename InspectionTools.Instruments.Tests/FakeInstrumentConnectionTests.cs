using InspectionTools.Common;

namespace InspectionTools.Instruments.Tests;

/// <summary>ダミーの接続が、要件定義書のテスト方針に書いた規則どおりに動くことを確認する</summary>
public class FakeInstrumentConnectionTests {

    // VISA 接続でコマンドを送る測定器を作る（Index は 0 以上で「使用中」を表す）
    private static DmmInstClass CreateDmm(string command, bool isQuery) =>
        new() { SignalType = 2, Index = 0, InstCommand = command, Query = isQuery };

    [Fact]
    public async Task UnusedInstrumentReturnsEmptyAndSendsNothing() {
        var fake = new FakeInstrumentConnection();
        var dmm = CreateDmm("*RST", isQuery: false);
        dmm.Index = -1;

        var reply = await fake.ConnectAsync(dmm);

        Assert.Equal(string.Empty, reply);
        Assert.Empty(fake.Sent);
    }

    [Fact]
    public async Task AdcInstrumentThrowsNotSupported() {
        var fake = new FakeInstrumentConnection();
        var dmm = CreateDmm("F1", isQuery: true);
        dmm.SignalType = 1;

        await Assert.ThrowsAsync<NotSupportedException>(() => fake.ConnectAsync(dmm));
    }

    [Fact]
    public async Task NonQueryCommandIsRecordedAndReturnsEmpty() {
        var fake = new FakeInstrumentConnection();

        var reply = await fake.ConnectAsync(CreateDmm("SIR2,SOI+4MA,OPR", isQuery: false));

        Assert.Equal(string.Empty, reply);
        Assert.Equal(["SIR2,SOI+4MA,OPR"], fake.Sent);
    }

    [Fact]
    public async Task OpcQueryReturnsOne() {
        var fake = new FakeInstrumentConnection();

        var reply = await fake.ConnectAsync(CreateDmm("*RST;:INIT:CONT 1;*OPC?", isQuery: true));

        Assert.Equal("1", reply);
    }

    [Fact]
    public async Task ReadCommandReturnsDefaultWhenNothingIsQueued() {
        var fake = new FakeInstrumentConnection();

        var reply = await fake.ConnectAsync(CreateDmm("FETC?", isQuery: true));

        Assert.Equal("1.234", reply);
    }

    [Fact]
    public async Task ReadCommandReturnsQueuedValuesInOrderThenDefault() {
        var fake = new FakeInstrumentConnection();
        fake.QueueReading("FETC?", "4.00", "4.50");
        var dmm = CreateDmm("FETC?", isQuery: true);

        var first = await fake.ConnectAsync(dmm);
        var second = await fake.ConnectAsync(dmm);
        var third = await fake.ConnectAsync(dmm);

        Assert.Equal(["4.00", "4.50", "1.234"], [first, second, third]);
    }

    [Fact]
    public async Task UnknownQueryThrows() {
        var fake = new FakeInstrumentConnection();

        await Assert.ThrowsAsync<InvalidOperationException>(() => fake.ConnectAsync(CreateDmm("*IDN?", isQuery: true)));
    }
}
