using InspectionTools.Common;

namespace InspectionTools.Instruments.Tests;

/// <summary>InstrumentService の接続を指定する版が、ダミー経由で正しいコマンドを送ることを確認する</summary>
public class InstrumentServiceTests {

    [Fact]
    public async Task ReadDmmSendsFetcAndReturnsDefaultReading() {
        var fake = new FakeInstrumentConnection();
        var dmm = new DmmInstClass { SignalType = 2, Index = 0, Query = true };

        var value = await InstrumentService.ReadDmmAsync(dmm, fake);

        Assert.Equal(["FETC?"], fake.Sent);
        Assert.Equal(1.234m, value);
    }

    [Fact]
    public async Task ReadDcsSendsMeasCurrAndReturnsQueuedReading() {
        var fake = new FakeInstrumentConnection();
        fake.QueueReading("MEAS:CURR?", "4.002");
        var dcs = new DcsInstClass { SignalType = 2, Index = 0, Query = true };

        var value = await InstrumentService.ReadDcsAsync(dcs, fake);

        Assert.Equal(["MEAS:CURR?"], fake.Sent);
        Assert.Equal(4.002m, value);
    }

    [Fact]
    public async Task ReadCntSendsMeasXnowAndParsesExponentNotation() {
        var fake = new FakeInstrumentConnection();
        fake.QueueReading(":MEAS?XNOW", "+1.234E-03");
        var cnt = new CntInstClass { SignalType = 3, Index = 0 };

        var value = await InstrumentService.ReadCntAsync(cnt, fake);

        Assert.Equal([":MEAS?XNOW"], fake.Sent);
        Assert.Equal(0.001234m, value);
    }

    [Fact]
    public async Task ReadOscSendsMeasuredIndexInCommand() {
        var fake = new FakeInstrumentConnection();
        var osc = new OscInstClass { SignalType = 2, Index = 0, Query = true };

        var value = await InstrumentService.ReadOscAsync(osc, 2, fake);

        Assert.Equal(["MEASU:MEAS2:VAL?"], fake.Sent);
        Assert.Equal(1.234m, value);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("")]
    public async Task UnparsableReadingBecomesZeroByDefaultToKeepTheOldAppBehavior(string reading) {
        var fake = new FakeInstrumentConnection();
        fake.QueueReading("FETC?", reading);
        var dmm = new DmmInstClass { SignalType = 2, Index = 0, Query = true };

        var value = await InstrumentService.ReadDmmAsync(dmm, fake);

        Assert.Equal(0m, value);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("")]
    public async Task UnparsableReadingThrowsWhenStrict(string reading) {
        var fake = new FakeInstrumentConnection();
        fake.QueueReading("FETC?", reading);
        var dmm = new DmmInstClass { SignalType = 2, Index = 0, Query = true };

        await Assert.ThrowsAsync<FormatException>(() => InstrumentService.ReadDmmAsync(dmm, fake, strict: true));
    }

    [Fact]
    public async Task StrictReadingStillParsesExponentNotation() {
        var fake = new FakeInstrumentConnection();
        fake.QueueReading(":MEAS?XNOW", "+1.234E-03");
        var cnt = new CntInstClass { SignalType = 3, Index = 0 };

        var value = await InstrumentService.ReadCntAsync(cnt, fake, strict: true);

        Assert.Equal(0.001234m, value);
    }
}
