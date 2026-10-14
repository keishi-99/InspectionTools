namespace InspectionTools.Products;

/// <summary>測定器の一覧（VisaAddress.xml）の1台分</summary>
/// <param name="Category">カテゴリ（例: デジタルマルチメータ）</param>
/// <param name="Name">測定器の名前（選ぶときに表示する。カテゴリの中で重複しない）</param>
/// <param name="VisaAddress">接続先（VISA のアドレス、または ADC のチャンネル番号）</param>
/// <param name="SignalType">通信方式（1=ADC、2=VISA USB、3=GPIB、4=VISA）</param>
public sealed record InstrumentEntry(string Category, string Name, string VisaAddress, int SignalType);
