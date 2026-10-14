namespace InspectionTools.Products;

/// <summary>測定器の役割と、その役割に取らせる状態の名前（例: Dcs の "4mA"、Dmm02 の "DCI"）</summary>
public sealed record InstrumentSetting(InstrumentRole Role, string Name);
