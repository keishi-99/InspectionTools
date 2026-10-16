namespace InspectionTools.Products.Tests;

/// <summary>接続の失敗の文面を確認する</summary>
public sealed class InstrumentConnectionExceptionTests {

    [Fact]
    public void MessageListsEveryFailureAndTellsThatTheConnectionWasReleased() {
        var exception = new InstrumentConnectionException([
            new InstrumentCommunicationException(InstrumentRole.Dmm01, "DMM01: 通信に失敗しました。"),
            new InstrumentCommunicationException(InstrumentRole.Fg, "FG: 通信に失敗しました。"),
        ]);

        Assert.Contains("・DMM01: 通信に失敗しました。", exception.Message);
        Assert.Contains("・FG: 通信に失敗しました。", exception.Message);
        Assert.EndsWith("接続を解除しました。選び直して、もう一度接続してください。", exception.Message);
    }
}
