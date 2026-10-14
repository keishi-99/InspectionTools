using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Threading;
using InspectionTools.App.ViewModels;
using InspectionTools.Inspection;

namespace InspectionTools.App;

public partial class MainWindow : Window {

    private static readonly Brush OkBrush = CreateBrush(0xC8, 0xE6, 0xC9);
    private static readonly Brush NgBrush = CreateBrush(0xFF, 0xCD, 0xD2);
    private static readonly Brush DoneBrush = CreateBrush(0xBB, 0xDE, 0xFB);

    private readonly MainViewModel _viewModel;

    // DCS の経過時間の表示を更新するタイマー
    private readonly DispatcherTimer _dcsTimer = new() { Interval = TimeSpan.FromSeconds(1) };

    // 表の選択を ViewModel の位置に合わせている最中は、選択変更のイベントを無視する（互いに呼び合うのを防ぐ）
    private bool _isSyncingSelection;


    public MainWindow(MainViewModel viewModel) {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = viewModel;
        viewModel.PropertyChanged += ViewModel_PropertyChanged;

        BuildGrid();
        Loaded += (_, _) => FocusInput();

        // DCS の経過時間を、1秒ごとに更新する
        _dcsTimer.Tick += (_, _) => _viewModel.UpdateDcsElapsed();
        _dcsTimer.Start();
    }

    private static Brush CreateBrush(byte r, byte g, byte b) {
        var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
        brush.Freeze();
        return brush;
    }

    private void ViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e) {
        switch (e.PropertyName) {
            case nameof(MainViewModel.Rows):
                BuildGrid();
                break;
            case nameof(MainViewModel.CurrentItemIndex):
            case nameof(MainViewModel.CurrentUnitIndex):
                SelectCurrentCell();
                FocusInput();
                break;
            case nameof(MainViewModel.IsInstrumentBusy) when !_viewModel.IsInstrumentBusy:
                // 通信が終わったら、入力欄に戻す（「値を取得」のあと、そのまま Enter で確定できるように）
                FocusInput();
                break;
        }
    }

    // 台数に合わせて列を作り直し、行のデータを結びつける
    private void BuildGrid() {
        _isSyncingSelection = true;
        try {
            ResultGrid.Columns.Clear();
            ResultGrid.Columns.Add(new DataGridTextColumn {
                Header = "検査項目",
                Binding = new Binding(nameof(RowViewModel.ItemName)),
                Width = 340,
            });
            for (var unitIndex = 0; unitIndex < _viewModel.UnitCount; unitIndex++) {
                ResultGrid.Columns.Add(new DataGridTextColumn {
                    Header = $"{unitIndex + 1}台目",
                    Binding = new Binding($"Cells[{unitIndex}].{nameof(CellViewModel.Display)}") { Mode = BindingMode.OneWay },
                    CellStyle = CreateCellStyle(unitIndex),
                    Width = 90,
                });
            }
            ResultGrid.ItemsSource = _viewModel.Rows;
        } finally {
            _isSyncingSelection = false;
        }
        SelectCurrentCell();
    }

    // 判定の状態によってマスの背景色を変える
    private Style CreateCellStyle(int unitIndex) {
        var style = new Style(typeof(DataGridCell), (Style)FindResource(typeof(DataGridCell)));
        var statePath = $"Cells[{unitIndex}].{nameof(CellViewModel.State)}";
        style.Triggers.Add(CreateStateTrigger(statePath, ResultState.Ok, OkBrush));
        style.Triggers.Add(CreateStateTrigger(statePath, ResultState.Ng, NgBrush));
        style.Triggers.Add(CreateStateTrigger(statePath, ResultState.Done, DoneBrush));
        return style;
    }

    private static DataTrigger CreateStateTrigger(string statePath, ResultState state, Brush brush) {
        var trigger = new DataTrigger { Binding = new Binding(statePath), Value = state };
        trigger.Setters.Add(new Setter(BackgroundProperty, brush));
        return trigger;
    }

    // ViewModel の現在位置のマスを、表の選択状態にする
    private void SelectCurrentCell() {
        if (ResultGrid.Columns.Count <= _viewModel.CurrentUnitIndex + 1 || _viewModel.Rows.Count <= _viewModel.CurrentItemIndex) {
            return;
        }

        _isSyncingSelection = true;
        try {
            var row = _viewModel.Rows[_viewModel.CurrentItemIndex];
            var column = ResultGrid.Columns[_viewModel.CurrentUnitIndex + 1];
            ResultGrid.SelectedCells.Clear();
            ResultGrid.SelectedCells.Add(new DataGridCellInfo(row, column));
            ResultGrid.CurrentCell = new DataGridCellInfo(row, column);
            ResultGrid.ScrollIntoView(row, column);
        } finally {
            _isSyncingSelection = false;
        }
    }

    // 利用者が表のマスを選んだら、ViewModel の現在位置を合わせる
    private async void ResultGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e) {
        if (_isSyncingSelection || ResultGrid.SelectedCells.Count != 1) {
            return;
        }
        // 通信中は位置を変えられない（ViewModel が受け付けない）ので、キー操作で動いた選択を元の位置に戻す
        if (_viewModel.IsInstrumentBusy) {
            SelectCurrentCell();
            return;
        }

        var cell = ResultGrid.SelectedCells[0];
        var unitIndex = cell.Column.DisplayIndex - 1;
        var itemIndex = _viewModel.Rows.ToList().IndexOf((RowViewModel)cell.Item);
        // 項目名の列（先頭）は選ばれても位置を変えない
        if (unitIndex >= 0 && itemIndex >= 0) {
            await _viewModel.MoveToAsync(itemIndex, unitIndex);
        }
    }

    // アプリを閉じるとき、測定器に接続していれば解除する
    private void Window_Closing(object? sender, CancelEventArgs e) {
        _viewModel.Shutdown();
    }

    // 台数の入力欄から離れたときも、入力された台数を反映する
    private void UnitCountTextBox_LostFocus(object sender, RoutedEventArgs e) => _viewModel.ApplyUnitCountTextCommand.Execute(null);

    private void FocusInput() {
        if (!_viewModel.IsValueInput) {
            return;
        }
        // 表示の切り替えが終わってからフォーカスを移す
        Dispatcher.BeginInvoke(() => {
            InputTextBox.Focus();
            InputTextBox.SelectAll();
        }, System.Windows.Threading.DispatcherPriority.Input);
    }
}
