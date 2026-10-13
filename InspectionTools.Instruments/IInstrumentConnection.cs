namespace InspectionTools.Common;

/// <summary>測定器との1回の通信（コマンド送信と、問い合わせのときの返事の受信）を行う</summary>
public interface IInstrumentConnection {

    /// <summary>InstClass に設定されたコマンドを送り、問い合わせ（Query）のときは返事を返す（それ以外は空文字）</summary>
    Task<string> ConnectAsync(InstClass instClass);
}
