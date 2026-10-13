using CommunityToolkit.Mvvm.ComponentModel;
using InspectionTools.Inspection;

namespace InspectionTools.App.ViewModels;

/// <summary>表の1マス（項目 × 台）の表示用の状態。結果そのものは InspectionSession が持ち、ここは画面に出す形に直すだけ</summary>
public sealed partial class CellViewModel(InspectionSession session, int itemIndex, int unitIndex) : ObservableObject {

    /// <summary>マスに表示する文字列</summary>
    public string Display => GetDisplay();

    /// <summary>色分けに使う状態</summary>
    public ResultState State => session.GetResult(itemIndex, unitIndex).State;

    /// <summary>セッションの内容が変わったあとに呼び、表示を更新する</summary>
    public void Refresh() {
        OnPropertyChanged(nameof(Display));
        OnPropertyChanged(nameof(State));
    }

    private string GetDisplay() {
        var result = session.GetResult(itemIndex, unitIndex);
        return session.Items[itemIndex].Kind switch {
            ItemKind.Measurement or ItemKind.Input => result.Value?.ToString() ?? "",
            ItemKind.Confirmation => result.State switch {
                ResultState.Ok => "合格",
                ResultState.Ng => "不合格",
                _ => "",
            },
            ItemKind.Task => result.State == ResultState.Done ? "完了" : "",
            _ => "",
        };
    }
}
