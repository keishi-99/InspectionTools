namespace InspectionTools.Products;

/// <summary>ある役割に選べる測定器の条件（測定器の一覧のうち、どれが候補になるか）</summary>
/// <param name="Category">測定器の一覧（VisaAddress.xml）のカテゴリ名</param>
/// <param name="SignalTypes">使える通信方式（1=ADC、2=VISA USB、3=GPIB、4=VISA）</param>
public sealed record RoleRequirement(string Category, IReadOnlyList<int> SignalTypes);
