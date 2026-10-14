using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using InspectionTools.Inspection;

namespace InspectionTools.App.ViewModels;

/// <summary>メニュー（製品の選択）と検査画面の切り替え。ウィンドウは1つで、選んだ製品の検査画面（MainViewModel）を Current に持つ</summary>
public sealed partial class ShellViewModel : ObservableObject {

    private readonly Func<ProductMenuItem, MainViewModel> _openProduct;
    private readonly IDialogService _dialog;

    /// <param name="openProduct">選ばれた製品の検査画面を作る（定義の読み込みに失敗したときは、ProductDefinitionException か IOException）</param>
    public ShellViewModel(ProductCatalogResult catalog, Func<ProductMenuItem, MainViewModel> openProduct, IDialogService dialog) {
        Items = catalog.Items;
        WarningText = catalog.Warnings.Count == 0
            ? ""
            : "読めなかった製品の定義があります。\n" + string.Join("\n", catalog.Warnings.Select(warning => $"・{warning}"));
        _openProduct = openProduct;
        _dialog = dialog;
    }

    /// <summary>メニューに並べる製品</summary>
    public IReadOnlyList<ProductMenuItem> Items { get; }

    /// <summary>読めなかった定義ファイルの案内（なければ空）</summary>
    public string WarningText { get; }

    public bool HasWarning => WarningText.Length > 0;

    /// <summary>検査中の製品（メニューを表示しているときは null）</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsMenuVisible))]
    private MainViewModel? _current;

    public bool IsMenuVisible => Current is null;

    [RelayCommand]
    private void SelectProduct(ProductMenuItem item) {
        try {
            var viewModel = _openProduct(item);
            viewModel.Start();
            Current = viewModel;
        } catch (Exception ex) when (ex is ProductDefinitionException or IOException) {
            _dialog.ShowError($"製品の定義を読み込めませんでした。\n{ex.Message}");
        }
    }

    /// <summary>検査画面からメニューに戻る。測定器は解除し（測定器には何も送らない）、途中保存はそのまま残す</summary>
    [RelayCommand]
    private void BackToMenu() {
        if (Current is null || Current.IsInstrumentBusy) {
            return;
        }
        Current.Shutdown();
        var product = Current.Product;
        Current = null;
        _dialog.ShowInfo($"{product} の検査を中断しました。測定器は解除しました。\n途中保存は残してあるので、もう一度選ぶと続きから再開できます。");
    }

    /// <summary>アプリを閉じる前に呼ぶ。検査中なら、測定器を解除する</summary>
    public void Shutdown() {
        Current?.Shutdown();
    }
}
