using System.ComponentModel;
using System.Windows;
using InspectionTools.App.ViewModels;

namespace InspectionTools.App;

public partial class InstrumentSelectionWindow : Window {

    private readonly InstrumentSelectionViewModel _viewModel;

    public InstrumentSelectionWindow(InstrumentSelectionViewModel viewModel) {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = viewModel;
    }

    // 接続や解除の最中は閉じない（途中で閉じると、測定器の状態が分からなくなるため）
    private void Window_Closing(object? sender, CancelEventArgs e) {
        e.Cancel = _viewModel.IsBusy;
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
}
