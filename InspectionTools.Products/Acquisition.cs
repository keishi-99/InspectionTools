namespace InspectionTools.Products;

/// <summary>測定器から値を取得するときの決まり</summary>
/// <param name="Role">値を読み取る測定器の役割</param>
/// <param name="Scale">読み取った値に掛ける倍率（A を mA にするなら 1000）</param>
/// <param name="Decimals">結果に残す小数の桁数（四捨五入。Excel に入れる値を、旧アプリが入力していた桁に合わせるため）</param>
public sealed record Acquisition(InstrumentRole Role, decimal Scale, int Decimals);
