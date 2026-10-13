namespace InspectionTools.Inspection;

/// <summary>製品ごとの検査項目の定義（製品名と、検査の流れの順に並んだ検査項目）</summary>
public sealed record ProductDefinition(string Product, IReadOnlyList<InspectionItem> Items);
