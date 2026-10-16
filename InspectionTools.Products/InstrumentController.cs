using InspectionTools.Common;

namespace InspectionTools.Products;

/// <summary>
/// 接続済みの測定器に、項目ごとの動作（状態の適用と値の取得）を行わせる。
/// 状態は覚えておき、変わった役割にだけコマンドを送る（測定器が設定の変更に時間がかかるため）
/// </summary>
/// <param name="procedure">製品の手順表</param>
/// <param name="instruments">使う役割ごとの測定器（使わない役割は含めない）</param>
/// <param name="connection">測定器との通信（テストではダミーに差し替える）</param>
public sealed class InstrumentController(
    ProductProcedure procedure,
    IReadOnlyDictionary<InstrumentRole, InstClass> instruments,
    IInstrumentConnection connection) {

    private readonly Dictionary<InstrumentRole, string> _states = [];

    /// <summary>役割ごとの現在の状態の名前（初期設定のあと、または状態を送ったあとのもの）</summary>
    public IReadOnlyDictionary<InstrumentRole, string> States => _states;

    /// <summary>測定器に状態を送れたときに起きる（送った役割と状態。送るのをやめた役割や、失敗した役割では起きない）</summary>
    public event Action<InstrumentSetting>? StateSent;

    /// <summary>その役割の測定器が接続されているか</summary>
    public bool IsConnected(InstrumentRole role) => instruments.ContainsKey(role);

    /// <summary>動作に必要なのに、接続されていない役割を返す（画面で「未接続」と知らせたり、ボタンを押せなくしたりするため）</summary>
    public IReadOnlyList<InstrumentRole> FindMissingRoles(ItemAction action) {
        var neededRoles = action.Settings.Select(setting => setting.Role);
        if (action.Acquisition is not null) {
            neededRoles = neededRoles.Append(action.Acquisition.Role);
        }
        return neededRoles.Distinct().Where(role => !IsConnected(role)).ToList();
    }

    /// <summary>
    /// 接続したすべての測定器に初期設定を送り、状態を初期状態にする。失敗した測定器があっても残りには送り続け、
    /// 最後に失敗した役割をすべて持つ InstrumentConnectionException にする（その場合、状態はすべて不明に戻る）
    /// </summary>
    public async Task InitializeAsync() {
        _states.Clear();
        var failures = new List<InstrumentCommunicationException>();
        foreach (var role in instruments.Keys.Order()) {
            var profile = procedure.Profiles[role];
            try {
                await SendAsync(role, profile.Init);
                _states[role] = profile.InitialState;
            } catch (InstrumentCommunicationException ex) {
                failures.Add(ex);
            }
        }

        if (failures.Count > 0) {
            _states.Clear();
            throw new InstrumentConnectionException(failures);
        }
    }

    /// <summary>
    /// 動作の状態を測定器に適用する。動作に書かれた役割はその状態へ、書かれていない役割は IdleStates にあればその状態へ、
    /// なければ触らない。接続されていない役割は無視する。現在の状態と同じ役割には何も送らない。
    /// ただし force が true のときは、記録の状態を無視して、対象の役割すべてに必ず送る（本体で操作して、記録とずれたときの対処）
    /// </summary>
    public async Task ApplyAsync(ItemAction action, bool force = false) {
        foreach (var change in OrderChanges(FindChanges(action, force))) {
            await SendAsync(change.Role, FindStateCommand(change));
            _states[change.Role] = change.Name;
            StateSent?.Invoke(change);
        }
    }

    /// <summary>動作に指定された測定器から値を取得し、倍率を掛けて、指定の桁に丸めて返す。取得できなかったときは InstrumentCommunicationException</summary>
    public async Task<decimal> AcquireAsync(Acquisition acquisition) {
        var role = acquisition.Role;
        var instrument = FindInstrument(role);

        decimal reading;
        try {
            reading = instrument switch {
                // 数値に変換できない返事は、0 にせずエラーにする（旧アプリは黙って 0 にしていた）。
                // DMM の読み取りは返事を受け取る問い合わせだが、InstrumentService はその指定を呼び出し側に任せているため、ここで指定する
                DmmInstClass dmm => await ReadDmmAsync(dmm),
                CntInstClass cnt => await InstrumentService.ReadCntAsync(cnt, connection, strict: true),
                OscInstClass osc => await ReadOscAsync(osc),
                _ => throw new NotSupportedException($"値の取得に対応していない測定器です: {instrument.GetType().Name}"),
            };
        } catch (FormatException ex) {
            // 通信はできたが、返事が数値として読めなかった
            throw new InstrumentCommunicationException(role, $"{DescribeInstrument(role)}の返事を数値として読めませんでした。値を取得できませんでした。（詳細: {ex.Message}）", ex);
        } catch (Exception ex) {
            throw new InstrumentCommunicationException(role, $"{CommunicationFailed(role)}値を取得できませんでした。ケーブルと電源を確認してください。（詳細: {ex.Message}）", ex);
        }

        // 読み取りのたびに次の測定を準備する測定器（カウンタ）は、取得の一部として準備のコマンドを送る
        if (procedure.Profiles[role].AfterRead is { } afterRead) {
            await SendAsync(role, afterRead);
        }

        return Math.Round(reading * acquisition.Scale, acquisition.Decimals, MidpointRounding.AwayFromZero);
    }

    // 動作が目指す状態のうち、送るものを、役割の順に並べて返す（force でなければ、現在の状態から変わるものだけ）
    private List<InstrumentSetting> FindChanges(ItemAction action, bool force) {
        var targets = action.Settings.ToDictionary(setting => setting.Role, setting => setting.Name);
        foreach (var (role, idleState) in procedure.IdleStates) {
            targets.TryAdd(role, idleState);
        }

        return targets
            .Where(target => IsConnected(target.Key) && (force || IsDifferentFromCurrent(target.Key, target.Value)))
            .OrderBy(target => target.Key)
            .Select(target => new InstrumentSetting(target.Key, target.Value))
            .ToList();
    }

    // 現在の状態と違うか（現在の状態が分からない役割は、違うものとして扱う）
    private bool IsDifferentFromCurrent(InstrumentRole role, string state) {
        return !_states.TryGetValue(role, out var current) || current != state;
    }

    // 状態を送る順序を決める
    private IEnumerable<InstrumentSetting> OrderChanges(List<InstrumentSetting> changes) {
        return changes.OrderBy(RankOf);
    }

    // 0: 出力を止める、1: それ以外、2: 出力を出す
    private int RankOf(InstrumentSetting change) {
        // 出力する機器でなければ（測定する機器なら）、止める・出すの区別はない
        if (!procedure.IdleStates.TryGetValue(change.Role, out var idleState)) {
            return 1;
        }
        return idleState == change.Name ? 0 : 2;
    }

    // DMM の測定値を読み取る
    private Task<decimal> ReadDmmAsync(DmmInstClass dmm) {
        dmm.Query = true;
        return InstrumentService.ReadDmmAsync(dmm, connection, strict: true);
    }

    // OSC の測定値を読み取る(測定項目は meas1 だけを使う。DMM と同じく、返事を受け取る問い合わせにする)
    private Task<decimal> ReadOscAsync(OscInstClass osc) {
        osc.Query = true;
        return InstrumentService.ReadOscAsync(osc, 1, connection, strict: true);
    }

    // 役割の状態の名前から、その状態にするコマンドを探す
    private SwitchCommand FindStateCommand(InstrumentSetting setting) {
        return procedure.Profiles[setting.Role].States.TryGetValue(setting.Name, out var command)
            ? command
            : throw new InvalidOperationException($"{procedure.Product} の手順表に、{setting.Role} の状態 '{setting.Name}' のコマンドがありません。");
    }

    // 役割の測定器を探す
    private InstClass FindInstrument(InstrumentRole role) {
        return instruments.TryGetValue(role, out var instrument)
            ? instrument
            : throw new InstrumentCommunicationException(role, $"{Describe(role)} の測定器が接続されていません。");
    }

    // 測定器の通信方式に合ったコマンドを、測定器に1回送る
    private async Task SendAsync(InstrumentRole role, SwitchCommand command) {
        var instrument = FindInstrument(role);
        var (text, isQuery) = InstrumentHelper.ResolveCommand(command, instrument.SignalType);
        if (string.IsNullOrEmpty(text)) {
            throw new InstrumentCommunicationException(role, $"{Describe(role)}: この通信方式（SignalType {instrument.SignalType}）のコマンドが手順表にありません。");
        }

        instrument.InstCommand = text;
        instrument.Query = isQuery;
        try {
            await connection.ConnectAsync(instrument);
        } catch (Exception ex) {
            throw new InstrumentCommunicationException(role, $"{CommunicationFailed(role)}ケーブルと電源を確認してください。（詳細: {ex.Message}）", ex);
        }
    }

    // 「測定器 Dmm01（DMM-01）」の形(機器の名前が分かるときは、名前も添える)。あとに助詞をつなげるために、名前があるときは、そのまま、ないときは、半角の空白を足す
    private string DescribeInstrument(InstrumentRole role) {
        var target = Describe(role);
        return target.EndsWith('）') ? $"測定器 {target}" : $"測定器 {target} ";
    }

    // 通信に失敗した、という文(例: 「測定器 Dmm01（DMM-01）との通信に失敗しました。」)
    private string CommunicationFailed(InstrumentRole role) => $"{DescribeInstrument(role)}との通信に失敗しました。";

    // エラーの文面に出す、役割と測定器の名前
    private string Describe(InstrumentRole role) {
        return instruments.TryGetValue(role, out var instrument) && !string.IsNullOrEmpty(instrument.Name)
            ? $"{role}（{instrument.Name}）"
            : role.ToString();
    }
}
