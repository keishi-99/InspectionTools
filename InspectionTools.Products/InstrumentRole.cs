namespace InspectionTools.Products;

/// <summary>検査で測定器が受け持つ役割（どの製品でも、役割の名前で測定器を指す。実際の機種は検査の前に選ぶ）</summary>
public enum InstrumentRole {
    /// <summary>マルチメーター1台目</summary>
    Dmm01,
    /// <summary>マルチメーター2台目</summary>
    Dmm02,
    /// <summary>周波数発生器</summary>
    Fg,
    /// <summary>電圧電流発生器</summary>
    Dcs,
    /// <summary>ユニバーサルカウンタ</summary>
    Cnt,
}
