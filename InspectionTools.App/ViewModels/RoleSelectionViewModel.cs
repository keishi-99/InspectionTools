using CommunityToolkit.Mvvm.ComponentModel;
using InspectionTools.Products;

namespace InspectionTools.App.ViewModels;

/// <summary>測定器の選択画面の1行（1つの役割と、その役割に選べる測定器）</summary>
public sealed partial class RoleSelectionViewModel(InstrumentRole role, string label) : ObservableObject {

    /// <summary>「使わない」を表す選択肢の表示</summary>
    public const string NoneName = "（使わない）";

    private IReadOnlyList<InstrumentEntry> _entries = [];

    /// <summary>役割</summary>
    public InstrumentRole Role { get; } = role;

    /// <summary>画面に出す役割の名前</summary>
    public string Label { get; } = label;

    /// <summary>選択肢の表示（先頭は「使わない」。そのあとに候補の測定器の名前が続く）</summary>
    [ObservableProperty]
    private IReadOnlyList<string> _candidateNames = [NoneName];

    /// <summary>選んだ選択肢の表示</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(SelectedEntry))]
    private string _selectedName = NoneName;

    /// <summary>接続の結果や、使わないことの表示</summary>
    [ObservableProperty]
    private string _resultText = "";

    /// <summary>選んだ測定器（「使わない」のときは null）</summary>
    public InstrumentEntry? SelectedEntry => _entries.FirstOrDefault(entry => entry.Name == SelectedName);

    /// <summary>候補を入れ替える。選んでいた測定器が候補に残っていれば選択を保ち、なければ「使わない」に戻す</summary>
    public void ReplaceCandidates(IReadOnlyList<InstrumentEntry> entries) {
        var previousName = SelectedName;
        _entries = entries;
        CandidateNames = [NoneName, .. entries.Select(entry => entry.Name)];
        SelectedName = CandidateNames.Contains(previousName) ? previousName : NoneName;
    }
}
