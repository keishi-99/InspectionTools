using InspectionTools.Common;
using System.Data;
using System.Windows;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;
using static InspectionTools.MainWindow;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.Product {
    /// <summary>
    /// EL9220UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL9220UserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DcsInstClass _instDcs = new();
        private readonly DmmInstClass _instDmm = new();
        private readonly FgInstClass _instFg = new();

        private record SwitchCommand {
            public DcsMode DcsMode { get; init; }
            public DmmMode DmmMode { get; init; }
            public string Text { get; init; } = string.Empty;
            public string Adc { get; init; } = string.Empty;
            public string Visa { get; init; } = string.Empty;
            public string Gpib { get; init; } = string.Empty;
            public bool Query { get; init; } = false;
        }
        private readonly Dictionary<InstClass, (SwitchCommand Init, List<SwitchCommand> Settings)> _dicCommands = [];

        public EL9220UserControl() {
            InitializeComponent();
        }

        #region IDisposable Implementation

        /// <summary>
        /// リソースの解放（IDisposableパターン）
        /// </summary>
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// リソースの解放処理
        /// </summary>
        /// <param name="disposing">マネージドリソースも解放する場合はtrue</param>
        protected virtual void Dispose(bool disposing) {
            if (_disposed) {
                return;
            }

            if (disposing) {
                // マネージドリソースの解放
                try {
                    // ホットキーのクリア
                    ClearHotKey();

                    // 計測器の解放
                    DisposeInstrument(_instDcs);
                    DisposeInstrument(_instDmm);
                    DisposeInstrument(_instFg);

                    // 辞書のクリア
                    _dicCommands.Clear();
                } catch (Exception ex) {
                    // Dispose中のエラーはログに記録するのみ
                    System.Diagnostics.Debug.WriteLine($"Dispose error: {ex.Message}");
                }
            }

            // アンマネージドリソースの解放（必要に応じて）
            // ...

            _disposed = true;
        }

        /// <summary>
        /// 個別の計測器インスタンスを解放
        /// </summary>
        private static void DisposeInstrument(InstClass instrument) {
            if (instrument == null) return;

            try {
                // 計測器がIDisposableを実装している場合
                if (instrument is IDisposable disposable) {
                    disposable.Dispose();
                }
                else {
                    // ResetPropertiesで状態をリセット
                    instrument.ResetProperties();
                }
            } catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine($"Instrument dispose error: {ex.Message}");
            }
        }

        /// <summary>
        /// オブジェクトが破棄済みかチェック
        /// </summary>
        private void ThrowIfDisposed() {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        /// <summary>
        /// ファイナライザ
        /// </summary>
        ~EL9220UserControl() {
            Dispose(false);
        }

        #endregion

        // 起動時
        private void LoadEvents() {
            ThrowIfDisposed();
            InstListImport();
            var parentWindow = Window.GetWindow(this);
            MainWindow.AdjustWindowSizeToUserControl(parentWindow);
        }
        private void InstListImport() {
            // デジタルマルチメータ、ファンクションジェネレータ、オシロスコープのコンボボックスを更新する
            MainWindow.UpdateComboBox(DcsComboBox, "パワーサプライ", [2]);
            MainWindow.UpdateComboBox(DmmComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(FgComboBox, "ファンクションジェネレータ", [2]);
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            MainWindow.IsProcessing = isVisible;
            MainGrid.IsEnabled = !isVisible;
            ProgressRing.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }

        // 選択した機器のVisaAddressを取得
        private void SelectInst() {
            MainWindow.GetVisaAddress(_instDcs, DcsComboBox);
            MainWindow.GetVisaAddress(_instDmm, DmmComboBox);
            MainWindow.GetVisaAddress(_instFg, FgComboBox);
        }

        // 機器設定辞書登録
        private void RegDictionary() {

            _dicCommands[_instDcs] =
                (
                    Init: new() { Visa = "*RST;:VOLT 3.6;*OPC?", Query = true },
                    Settings: [
                        new() { Text = "OFF",   Visa = $":OUTPUT OFF;*OPC?",    Query = true },
                        new() { Text = "ON",    Visa = $":OUTPUT ON;*OPC?",     Query = true },
                    ]
                );

            _dicCommands[_instDmm] =
                (
                    Init: new() { DmmMode = DmmMode.DCI, Adc = "*RST,F5,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instFg] =
                (
                    Init: new() { Visa = "*RST;:FREQ 2000;:VOLT 3.0VPP;:VOLT:OFFS 1.5;:OUTPUT OFF;*OPC?", Query = true },
                    Settings: [
                        new(){Text= "OFF",      Visa= ":FREQ 2000;:OUTPUT OFF;*OPC?", Query=true},
                        new(){Text= "2000",     Visa= ":OUTPUT ON;*OPC?", Query=true},
                        new(){Text= "400",      Visa= ":OUTPUT OFF;:FREQ 400;:OUTPUT ON;*OPC?", Query=true},
                        new(){Text= "50",       Visa= ":OUTPUT OFF;:FREQ 50;:OUTPUT ON;*OPC?", Query=true},
                    ]
                );
        }
        // 機器初期設定
        private void FormatSet() {
            (_instDcs.InstCommand, _instDcs.Query) = ResolveCommand(_dicCommands[_instDcs].Init, _instDcs.SignalType);
            (_instDmm.InstCommand, _instDmm.Query) = ResolveCommand(_dicCommands[_instDmm].Init, _instDmm.SignalType);
            (_instFg.InstCommand, _instFg.Query) = ResolveCommand(_dicCommands[_instFg].Init, _instFg.SignalType);
        }
        private static (string Cmd, bool Query) ResolveCommand(SwitchCommand sw, int signalType) {
            return signalType switch {
                1 => (sw.Adc, sw.Query),
                2 => (sw.Visa, sw.Query),
                3 => (sw.Gpib, sw.Query),
                _ => (string.Empty, false),
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {

            ThrowIfDisposed();

            try {
                _mainWindow?.SetButtonEnabled("ProductListButton", false);

                HotKeyCheckBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();

                RegDictionary();
                FormatSet();

                InstClass[] devices = [_instDcs, _instDmm, _instFg];

                await Task.Run(() =>
                    DeviceConnectionHelper.ConnectInParallelAsync(devices)
                );

                DcsComboBox.IsEnabled = false;
                DmmComboBox.IsEnabled = false;
                FgComboBox.IsEnabled = false;
                ConnectButton.IsEnabled = false;
                ReleaseButton.IsEnabled = true;

            } catch (AggregateException aex) {
                Release();
                var messages = string.Join("\n", aex.InnerExceptions.Select(e => e.Message));
                MessageBox.Show(messages, "接続エラー");
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー");
            } finally {
                VisibleProgressImage(false);
            }
        }

        // 解除
        private void Release() {
            VisibleProgressImage(false);

            _instDcs.ResetProperties();
            _instDmm.ResetProperties();
            _instFg.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
            DcsComboBox.IsEnabled = true;
            DmmComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyCheckBox.IsChecked = false;
        }

        // DMM測定値取得
        private async Task<decimal> ReadDmm(DmmInstClass dmmInstClass) {

            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var output = await InstrumentService.ReadDmmAsync(dmmInstClass);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // DCS切り替え
        private async Task SwitchDcs(DcsInstClass dcsInstClass, bool isNext) {

            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var settings = _dicCommands[dcsInstClass].Settings;
                dcsInstClass.SettingNumber = (dcsInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[dcsInstClass.SettingNumber];
                (dcsInstClass.InstCommand, dcsInstClass.Query) = ResolveCommand(sw, dcsInstClass.SignalType);

                await DeviceController.ConnectAsync(dcsInstClass);

                VisibleProgressImage(false);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // FG切り替え
        private async void RotationFg(FgInstClass fgInstClass, bool isNext) {
            ThrowIfDisposed();

            try {
                if (string.IsNullOrEmpty(fgInstClass.VisaAddress)) { return; }
                VisibleProgressImage(true);

                var settings = _dicCommands[fgInstClass].Settings;
                fgInstClass.SettingNumber = (fgInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[fgInstClass.SettingNumber];
                (fgInstClass.InstCommand, fgInstClass.Query) = ResolveCommand(sw, fgInstClass.SignalType);

                if (fgInstClass.InstCommand == string.Empty) { return; }

                await InstrumentService.RotateFgAsync(fgInstClass);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }

        // DCSローテーション
        private async void ActionHotkeyPeriod() {
            if (MainWindow.IsProcessing) { return; }
            await SwitchDcs(_instDcs, true);
        }
        private async void ActionHotkeyNumDivide() {
            if (MainWindow.IsProcessing) { return; }
            await SwitchDcs(_instDcs, true);
        }
        // DMM測定値コピー
        private async void ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000000).ToString());  // μA単位に変換
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private async void ActionHotkeyNumSubtract() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000000).ToString());  // μA単位に変換
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // FGローテーション
        private async void ActionHotkeyBackslash() {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, true);
        }
        private async void ActionHotkeyNumMultiply() {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, true);
        }

        // HotKeyの登録
        private void SetHotKey() {

            ThrowIfDisposed();

            MainWindow.HotkeysList.Clear();

            if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                    new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyNumSubtract, ActionHotkeyNumSubtract),
                ]);
            }
            if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    new(ModNone, HotkeyNumMultiply, ActionHotkeyNumMultiply),
                ]);
            }

            MainWindow.SetHotKey();
        }
        private static void ClearHotKey() {
            MainWindow.ClearHotKey();
        }

        // イベントハンドラ
        private void UserControl_Loaded(object sender, RoutedEventArgs e) { LoadEvents(); }
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void HotKeyCheckBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyCheckBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }
        private void UserControl_Unloaded(object sender, RoutedEventArgs e) { Dispose(); }
    }
}
