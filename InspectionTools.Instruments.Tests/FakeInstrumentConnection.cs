using InspectionTools.Common;
using System.Text.RegularExpressions;

namespace InspectionTools.Instruments.Tests;

/// <summary>測定器の代わりに使うダミーの接続（送られたコマンドを記録し、規則に従って返事を返す）</summary>
public sealed class FakeInstrumentConnection : IInstrumentConnection {

    // 測定値の返事を指定しなかったときに返す値
    private const string DefaultReading = "1.234";

    private readonly List<string> _sent = [];
    private readonly Dictionary<string, Queue<string>> _readings = [];

    /// <summary>送られたコマンドの一覧（送られた順）</summary>
    public IReadOnlyList<string> Sent => _sent;

    /// <summary>測定値の読み取りコマンドに返す値を、指定した順に返すよう登録する（登録した分を返し終えたら標準の値を返す）</summary>
    public void QueueReading(string command, params string[] values) {
        if (!_readings.TryGetValue(command, out var queue)) {
            queue = new Queue<string>();
            _readings[command] = queue;
        }
        foreach (var value in values) {
            queue.Enqueue(value);
        }
    }

    public Task<string> ConnectAsync(InstClass instClass) {
        // 未使用の測定器には何も送らない（本物の ConnectAsync と同じ）
        if (instClass.Index == -1) {
            return Task.FromResult(string.Empty);
        }

        // VISA 以外（ADC など）はテストの対象外なので、気づけるように例外にする
        if (instClass.SignalType is not (2 or 3 or 4)) {
            throw new NotSupportedException($"ダミーは VISA のみ対応です（SignalType: {instClass.SignalType}）");
        }

        _sent.Add(instClass.InstCommand);

        // 返事を求めないコマンドは、送った記録だけを残す
        if (!instClass.Query) {
            return Task.FromResult(string.Empty);
        }

        return Task.FromResult(ResolveReply(instClass.InstCommand));
    }

    // 問い合わせに対する返事を決める（想定外の問い合わせは、コマンドの打ち間違いに気づくため例外にする）
    private string ResolveReply(string command) {
        if (command.EndsWith("*OPC?")) {
            return "1";
        }

        if (IsReadCommand(command)) {
            return _readings.TryGetValue(command, out var queue) && queue.Count > 0
                ? queue.Dequeue()
                : DefaultReading;
        }

        throw new InvalidOperationException($"想定外の問い合わせです: {command}");
    }

    /// <summary>測定値を読み取るコマンドかどうかを判定する（FETC?、MEAS:CURR?、:MEAS?XNOW、MEASU:MEASn:VAL? が対象）</summary>
    internal static bool IsReadCommand(string command) {
        var pattern = "^FETC\\?$|^MEAS:CURR\\?$|^:MEAS\\?XNOW$|^MEASU:MEAS\\d+:VAL\\?$";
        return Regex.IsMatch(command, pattern);
    }
}
