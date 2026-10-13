using System.Windows;
using Microsoft.Win32;
using InspectionTools.App.ViewModels;

namespace InspectionTools.App;

/// <summary>メッセージボックスで確認・通知を行う</summary>
public sealed class WpfDialogService : IDialogService {

    /// <summary>メッセージボックスの親にするウィンドウ（表示前は null のまま）</summary>
    public Window? Owner { get; set; }

    public bool Confirm(string message) =>
        Show(message, "確認", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;

    public void ShowError(string message) =>
        Show(message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);

    public void ShowInfo(string message) =>
        Show(message, "完了", MessageBoxButton.OK, MessageBoxImage.Information);

    public string? SelectTemplateFilePath() {
        var dialog = new OpenFileDialog {
            Title = "転記先のテンプレートを選択",
            Filter = "Excel ブック (*.xlsx)|*.xlsx",
        };
        return dialog.ShowDialog(Owner) == true ? dialog.FileName : null;
    }

    private MessageBoxResult Show(string message, string caption, MessageBoxButton button, MessageBoxImage image) =>
        Owner is null
            ? MessageBox.Show(message, caption, button, image)
            : MessageBox.Show(Owner, message, caption, button, image);
}
