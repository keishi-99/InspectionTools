using InspectionTools.Common;

namespace InspectionTools.Products.Tests;

/// <summary>別の接続（ダミー）に処理を渡しながら、どの測定器にどのコマンドを送ったかを記録する（ダミーはコマンドだけを記録し、測定器の名前は残さないため）</summary>
public sealed class LoggingConnection(IInstrumentConnection inner) : IInstrumentConnection {

    /// <summary>送られた（測定器の名前、コマンド）の一覧（送られた順）</summary>
    public List<(string Instrument, string Command)> Sent { get; } = [];

    /// <summary>true を返した測定器への送信を失敗させる（通信エラーの再現用）</summary>
    public Func<InstClass, bool> ShouldFail { get; set; } = _ => false;

    public async Task<string> ConnectAsync(InstClass instClass) {
        Sent.Add((instClass.Name, instClass.InstCommand));
        if (ShouldFail(instClass)) {
            throw new IOException("通信できません。");
        }
        return await inner.ConnectAsync(instClass);
    }
}
