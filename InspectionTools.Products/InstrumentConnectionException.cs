namespace InspectionTools.Products;

/// <summary>測定器への接続（初期設定の送信）に、1台以上失敗したときの例外。失敗した役割をすべて持つ</summary>
public sealed class InstrumentConnectionException(IReadOnlyList<InstrumentCommunicationException> failures)
    : Exception(CreateMessage(failures)) {

    /// <summary>失敗した役割ごとの例外</summary>
    public IReadOnlyList<InstrumentCommunicationException> Failures { get; } = failures;

    private static string CreateMessage(IReadOnlyList<InstrumentCommunicationException> failures) {
        return "測定器に接続できませんでした。\n" + string.Join("\n", failures.Select(failure => $"・{failure.Message}")) + "\n接続を解除しました。選び直して、もう一度接続してください。";
    }
}
