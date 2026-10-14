namespace InspectionTools.Products;

/// <summary>測定器との通信や、測定値の取得に失敗したときの例外（どの役割の測定器かを持つ）</summary>
public sealed class InstrumentCommunicationException(InstrumentRole role, string message, Exception? innerException = null)
    : Exception(message, innerException) {

    /// <summary>失敗した測定器の役割</summary>
    public InstrumentRole Role { get; } = role;
}
