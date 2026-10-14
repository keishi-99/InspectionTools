namespace InspectionTools.Products;

/// <summary>製品ごとの測定器の手順表（役割ごとのコマンド表と、動作 ID ごとの動作）</summary>
/// <param name="Product">製品名</param>
/// <param name="Profiles">役割ごとのコマンド表</param>
/// <param name="Actions">動作 ID ごとの動作</param>
/// <param name="Requirements">役割ごとに選べる測定器の条件（測定器の一覧から候補を絞り込むため）</param>
/// <param name="IdleStates">
/// 項目が指定しなかったときに、自動で戻す役割と状態（出力する機器を出しっぱなしにしないため）。
/// ここにない役割（測定する機器）は、項目が指定しない限り触らない
/// </param>
public sealed record ProductProcedure(
    string Product,
    IReadOnlyDictionary<InstrumentRole, InstrumentProfile> Profiles,
    IReadOnlyDictionary<string, ItemAction> Actions,
    IReadOnlyDictionary<InstrumentRole, string> IdleStates,
    IReadOnlyDictionary<InstrumentRole, RoleRequirement> Requirements) {

    /// <summary>動作 ID から動作を引く。ID のない項目は測定器を使わない動作になり、手順表にない ID は例外にする（ID の書き間違いに気づくため）</summary>
    public ItemAction FindAction(string? actionId) {
        if (actionId is null) {
            return ItemAction.None;
        }
        return Actions.TryGetValue(actionId, out var action)
            ? action
            : throw new KeyNotFoundException($"{Product} の手順表にない動作 ID です: {actionId}");
    }
}
