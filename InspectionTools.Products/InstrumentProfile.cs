using InspectionTools.Common;

namespace InspectionTools.Products;

/// <summary>1つの役割の測定器を、製品の検査で使うためのコマンド表</summary>
/// <param name="Init">接続した直後に送る初期設定のコマンド</param>
/// <param name="InitialState">初期設定を送った直後の状態の名前</param>
/// <param name="States">状態の名前ごとの、その状態にするコマンド</param>
/// <param name="AfterRead">値を読み取った直後に送るコマンド（読み取りのたびに次の測定を準備する測定器だけ）</param>
public sealed record InstrumentProfile(
    SwitchCommand Init,
    string InitialState,
    IReadOnlyDictionary<string, SwitchCommand> States,
    SwitchCommand? AfterRead = null);
