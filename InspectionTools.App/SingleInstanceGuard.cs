namespace InspectionTools.App;

/// <summary>アプリを同時に1つしか起動できないようにする（名前付きミューテックス）。途中保存、測定器との通信、全体のホットキーが、2つのアプリで取り合いにならないようにするため</summary>
public sealed class SingleInstanceGuard : IDisposable {

    private readonly Mutex _mutex;

    /// <summary>自分が最初の1つか（false なら、すでに別のアプリが起動している）</summary>
    public bool IsFirstInstance { get; }

    /// <param name="name">ミューテックスの名前（同じ名前の起動どうしで、1つだけが通る）</param>
    public SingleInstanceGuard(string name) {
        _mutex = new Mutex(initiallyOwned: true, name, out var createdNew);
        IsFirstInstance = createdNew;
    }

    public void Dispose() {
        if (IsFirstInstance) {
            _mutex.ReleaseMutex();
        }
        _mutex.Dispose();
    }
}
