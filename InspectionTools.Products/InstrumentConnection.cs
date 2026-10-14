using InspectionTools.Common;

namespace InspectionTools.Products;

/// <summary>
/// 選んだ測定器への接続を管理する（接続すると InstrumentController ができ、解除すると破棄される）。
/// 一部の測定器に接続できなかったときは、接続できた分も含めてすべて解除する
/// </summary>
/// <param name="procedure">製品の手順表</param>
/// <param name="transport">測定器との通信（テストではダミーに差し替える）</param>
public sealed class InstrumentConnection(ProductProcedure procedure, IInstrumentConnection transport) {

    private List<InstClass> _instruments = [];
    private List<InstrumentRole> _roles = [];

    /// <summary>接続中の測定器を操作するコントローラー（接続していないときは null）</summary>
    public InstrumentController? Controller { get; private set; }

    /// <summary>接続しているか</summary>
    public bool IsConnected => Controller is not null;

    /// <summary>接続している役割（役割の順）</summary>
    public IReadOnlyList<InstrumentRole> ConnectedRoles => _roles;

    /// <summary>
    /// 選んだ測定器（役割ごと）に接続し、初期設定を送る。1台でも失敗したら、すべて解除して InstrumentConnectionException にする。
    /// 同じ測定器が複数の役割に選ばれているときは、接続せずに ArgumentException にする
    /// </summary>
    public async Task ConnectAsync(IReadOnlyDictionary<InstrumentRole, InstrumentEntry> selections) {
        if (IsConnected) {
            throw new InvalidOperationException("すでに接続しています。解除してから接続してください。");
        }
        if (selections.Count == 0) {
            throw new ArgumentException("接続する測定器が選ばれていません。", nameof(selections));
        }

        var duplicate = selections.GroupBy(selection => selection.Value.Name).FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null) {
            var roles = string.Join("、", duplicate.Select(selection => selection.Key));
            throw new ArgumentException($"同じ測定器（{duplicate.Key}）が複数の役割（{roles}）に選ばれています。", nameof(selections));
        }

        var instruments = selections.ToDictionary(selection => selection.Key, selection => CreateInstrument(selection.Key, selection.Value));
        var controller = new InstrumentController(procedure, instruments, transport);
        try {
            await controller.InitializeAsync();
        } catch {
            foreach (var instrument in instruments.Values) {
                InstrumentHelper.SafeDispose(instrument);
            }
            throw;
        }

        Controller = controller;
        _instruments = [.. instruments.Values];
        _roles = [.. instruments.Keys.Order()];
    }

    /// <summary>
    /// 測定器の接続を解除する。測定器には何も送らない（出力が ON のままでも、そのまま残す。必要なら本体で止める）。
    /// 接続していないときは何もしない
    /// </summary>
    public void Disconnect() {
        foreach (var instrument in _instruments) {
            InstrumentHelper.SafeDispose(instrument);
        }
        Controller = null;
        _instruments = [];
        _roles = [];
    }

    // 役割に合った種類の InstClass を作り、選んだ測定器の情報を入れる
    private static InstClass CreateInstrument(InstrumentRole role, InstrumentEntry entry) {
        InstClass instrument = role switch {
            InstrumentRole.Dmm01 or InstrumentRole.Dmm02 => new DmmInstClass(),
            InstrumentRole.Fg => new FgInstClass(),
            InstrumentRole.Dcs => new DcsInstClass(),
            InstrumentRole.Cnt => new CntInstClass(),
            _ => throw new ArgumentOutOfRangeException(nameof(role), role, "未対応の役割です。"),
        };
        instrument.Category = entry.Category;
        instrument.Name = entry.Name;
        instrument.VisaAddress = entry.VisaAddress;
        instrument.SignalType = entry.SignalType;
        return instrument;
    }
}
