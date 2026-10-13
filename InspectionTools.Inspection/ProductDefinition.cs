namespace InspectionTools.Inspection;

/// <summary>製品ごとの定義（製品名、最大台数、転記の設定と、検査の流れの順に並んだ検査項目）</summary>
public sealed record ProductDefinition(string Product, int MaxUnitCount, TransferDefinition Transfer, IReadOnlyList<InspectionItem> Items);
