using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using InspectionTools.App.Input;
using InspectionTools.App.ViewModels;
using InspectionTools.Inspection;

namespace InspectionTools.App;

public partial class MainWindow : Window {

    private static readonly Brush OkBrush = CreateBrush(0xC8, 0xE6, 0xC9);
    private static readonly Brush NgBrush = CreateBrush(0xFF, 0xCD, 0xD2);
    private static readonly Brush DoneBrush = CreateBrush(0xBB, 0xDE, 0xFB);

    private readonly ShellViewModel _shell;
    private MainViewModel? _viewModel;
    private readonly KeyBindings _keyBindings;
    private readonly IDialogService _dialog;

    // 全体のホットキー（前面に出すキー）。登録するとシステム全体でそのキーが奪われるので、ボタンで有効にしたときだけ登録する
    private readonly GlobalHotkey _globalHotkey = new();

    // 画面のキー（keybindings.json の割り当て）を使うか。無効のときは、通常のキー入力に戻り、操作はマウスだけになる
    private bool _isLocalKeysEnabled = true;

    // DCS の経過時間の表示を更新するタイマー
    private readonly DispatcherTimer _dcsTimer = new() { Interval = TimeSpan.FromSeconds(1) };

    // 表の選択を ViewModel の位置に合わせている最中は、選択変更のイベントを無視する（互いに呼び合うのを防ぐ）
    private bool _isSyncingSelection;


    public MainWindow(ShellViewModel shell, KeyBindings keyBindings, IDialogService dialog) {
        InitializeComponent();
        _shell = shell;
        _keyBindings = keyBindings;
        _dialog = dialog;
        _globalHotkey.Pressed += BringToFront;
        UpdateKeySwitches();
        DataContext = shell;
        shell.PropertyChanged += Shell_PropertyChanged;

        // DCS の経過時間を、1秒ごとに更新する
        _dcsTimer.Tick += (_, _) => _viewModel?.UpdateDcsElapsed();
        _dcsTimer.Start();
    }

    // メニューと検査画面を切り替える（検査画面の ViewModel は、製品を選ぶたびに入れ替わる）
    private void Shell_PropertyChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName != nameof(ShellViewModel.Current)) {
            return;
        }

        if (_viewModel is not null) {
            _viewModel.PropertyChanged -= ViewModel_PropertyChanged;
        }
        _viewModel = _shell.Current;
        InspectionRoot.DataContext = _viewModel;
        InspectionRoot.Visibility = _viewModel is null ? Visibility.Collapsed : Visibility.Visible;
        if (_viewModel is null) {
            ResultGrid.ItemsSource = null;
            return;
        }

        _viewModel.PropertyChanged += ViewModel_PropertyChanged;
        BuildGrid();
        FocusInput();
    }

    // 全体のホットキーの有効・無効を切り替える。有効にできなかったとき（他のソフトが同じキーを使っているなど）は、理由を知らせて無効のままにする
    private void GlobalHotkeySwitch_Click(object sender, RoutedEventArgs e) {
        if (_globalHotkey.IsRegistered) {
            _globalHotkey.Unregister();
        } else if (_keyBindings.GlobalHotkey is { } chord
            && !_globalHotkey.TryRegister(new WindowInteropHelper(this).Handle, chord, out var error)) {
            _dialog.ShowError($"全体のホットキーを有効にできません。\n{error}");
        }
        UpdateKeySwitches();
        FocusInput();
    }

    // 画面のキーの有効・無効を切り替える
    private void LocalKeysSwitch_Click(object sender, RoutedEventArgs e) {
        _isLocalKeysEnabled = !_isLocalKeysEnabled;
        UpdateKeySwitches();
        FocusInput();
    }

    // 2つのスイッチの見た目と、キーの案内を、いまの状態に合わせる
    private void UpdateKeySwitches() {
        LocalKeysSwitch.IsChecked = _isLocalKeysEnabled;
        LocalKeysState.Text = _isLocalKeysEnabled ? "有効" : "無効";
        KeyHintTextBlock.Text = _isLocalKeysEnabled ? _keyBindings.HintText : "画面のキー: 無効（マウスで操作してください）";

        GlobalHotkeySwitch.IsChecked = _globalHotkey.IsRegistered;
        if (_keyBindings.GlobalHotkey is not { } chord) {
            GlobalHotkeyLabel.Text = "前面ホットキー";
            GlobalHotkeyState.Text = "未設定";
            GlobalHotkeySwitch.IsEnabled = false;
            return;
        }
        GlobalHotkeyLabel.Text = $"前面ホットキー [{chord}]";
        GlobalHotkeyState.Text = _globalHotkey.IsRegistered ? "有効" : "無効";
    }

    /// <summary>全体のホットキーが押されたとき、このウィンドウを前面に出してフォーカスする</summary>
    private void BringToFront() {
        if (WindowState == WindowState.Minimized) {
            WindowState = WindowState.Normal;
        }
        // 他のソフトが前面にあっても出せるように、一度だけ最前面にしてすぐ戻す
        Activate();
        Topmost = true;
        Topmost = false;
        FocusInput();
    }

    private void BackToMenuButton_Click(object sender, RoutedEventArgs e) => _shell.BackToMenuCommand.Execute(null);

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
        if (_viewModel is null) {
            return;
        }
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
        if (_viewModel is null) {
            return;
        }
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
        if (_viewModel is null || _isSyncingSelection || ResultGrid.SelectedCells.Count != 1) {
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

    // キー割り当てに合うキーを受けたら、対応する操作を行う（入力欄やボタンより先に、画面全体で受ける）
    private async void Window_PreviewKeyDown(object sender, KeyEventArgs e) {
        // メニューを表示しているとき、画面のキーを無効にしているときは、キー操作を使わない
        if (_viewModel is null || !_isLocalKeysEnabled) {
            return;
        }

        // 台数の入力欄では、キー割り当てを使わない（Enter で台数を確定し、数字を入力するため）
        if (UnitCountTextBox.IsKeyboardFocusWithin) {
            return;
        }

        // Alt を伴うキーは、Key.System で届くので、本来のキーを取り出す
        var key = e.Key == Key.System ? e.SystemKey : e.Key;
        if (!_keyBindings.TryFind(new KeyChord(key, Keyboard.Modifiers), out var action)) {
            return;
        }

        // 割り当てのあるキーは、入力欄やボタンに渡さない（Space でフォーカス中のボタンが押されるのを防ぐ）。押しっぱなしの繰り返しは、操作にしない
        e.Handled = true;
        if (e.IsRepeat) {
            return;
        }
        await _viewModel.ExecuteKeyActionAsync(action);
    }

    // アプリを閉じるとき、測定器に接続していれば解除する
    private void Window_Closing(object? sender, CancelEventArgs e) {
        _globalHotkey.Dispose();
        _shell.Shutdown();
    }

    // 台数の入力欄から離れたときも、入力された台数を反映する
    private void UnitCountTextBox_LostFocus(object sender, RoutedEventArgs e) => _viewModel?.ApplyUnitCountTextCommand.Execute(null);

    private void FocusInput() {
        if (_viewModel is null || !_viewModel.IsValueInput) {
            return;
        }
        // 表示の切り替えが終わってからフォーカスを移す
        Dispatcher.BeginInvoke(() => {
            InputTextBox.Focus();
            InputTextBox.SelectAll();
        }, System.Windows.Threading.DispatcherPriority.Input);
    }
}
