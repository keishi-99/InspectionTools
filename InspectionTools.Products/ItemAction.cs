namespace InspectionTools.Products;

/// <summary>検査項目ごとに測定器へ行わせる動作（目標の状態と、値の取得）</summary>
/// <param name="Settings">この項目で測定器に取らせる状態。ここに書かない役割の扱いは、製品の手順表の IdleStates による</param>
/// <param name="Acquisition">測定器から値を取得する項目だけ指定する（手入力や、確認項目、作業項目は null）</param>
public sealed record ItemAction(IReadOnlyList<InstrumentSetting> Settings, Acquisition? Acquisition = null) {

    /// <summary>測定器を使わない項目の動作（actionId のない項目はこれになる）</summary>
    public static ItemAction None { get; } = new([]);
}
