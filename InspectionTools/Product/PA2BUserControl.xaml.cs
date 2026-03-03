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
    /// PA2BUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class PA2BUserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DmmInstClass _instDmm = new();
        private readonly OscInstClass _instOsc = new();

        private record SwitchCommand {
            public DmmMode Mode { get; init; }
            public string Text { get; init; } = string.Empty;
            public string Adc { get; init; } = string.Empty;
            public string Visa { get; init; } = string.Empty;
            public string Gpib { get; init; } = string.Empty;
            public bool Query { get; init; } = false;
        }
        private readonly Dictionary<InstClass, (SwitchCommand Init, List<SwitchCommand> Settings)> _dicCommands = [];

        public PA2BUserControl() {
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
                    DisposeInstrument(_instDmm);
                    DisposeInstrument(_instOsc);

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
        ~PA2BUserControl() {
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
            MainWindow.UpdateComboBox(DmmComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(OscComboBox, "オシロスコープ", [2]);
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            MainWindow.IsProcessing = isVisible;
            MainGrid.IsEnabled = !isVisible;
            ProgressRing.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }
        // 選択した機器のVisaAddressを取得
        private void SelectInst() {
            MainWindow.GetVisaAddress(_instDmm, DmmComboBox);
            MainWindow.GetVisaAddress(_instOsc, OscComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicCommands[_instDmm] =
                (
                    Init: new() { Mode = DmmMode.DCV, Adc = "*RST,F1,R0,*OPC?", Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG:AUTO ON;*OPC?", Query = true },
                    Settings: [
                            new() { Mode = DmmMode.DCV,   Adc= "*RST,F1,R0,*OPC?",    Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG:AUTO ON;*OPC?", Query = true },
                            new() { Mode = DmmMode.DCI,   Adc= "*RST,F5,R6,*OPC?",    Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?", Query = true },
                        ]
                );

            _dicCommands[_instOsc] =
                (
                    Init: new() {
                        Visa =
                            """
                            *RST;
                            :HEADER 0;
                            :SELECT:CH1 1;CH2 1;CH3 1;CH4 1;
                            :HORIZONTAL:MAIN:SCALE 1.0E-5;
                            :HORIZONTAL:POSITION -12.0E-6;
                            :CH1:SCALE 2;POSITION 0;
                            :CH2:SCALE 2;POSITION 0;
                            :CH3:SCALE 5;POSITION 2;
                            :CH4:SCALE 5;POSITION -3;
                            :TRIGGER:MAIN:EDGE:SOURCE CH2;
                            :TRIGGER:MAIN:LEVEL 3.0;
                            :MEASUREMENT:MEAS1:TYPE PK2PK;SOURCE CH2;
                            :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH2;
                            :MEASUREMENT:MEAS3:TYPE MINIMUM;SOURCE CH2;
                            :MEASUREMENT:MEAS4:TYPE NONE;SOURCE MATH;
                            :MEASUREMENT:MEAS5:TYPE NONE;SOURCE MATH;
                            *OPC?
                            """,
                        Query = true
                    },
                    Settings: []
                );
        }
        // 機器初期設定
        private void FormatSet() {
            (_instDmm.InstCommand, _instDmm.Query) = ResolveCommand(_dicCommands[_instDmm].Init, _instDmm.SignalType);
            (_instOsc.InstCommand, _instOsc.Query) = ResolveCommand(_dicCommands[_instOsc].Init, _instOsc.SignalType);
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

                InstClass[] devices = [_instDmm, _instOsc];

                await Task.Run(() =>
                    DeviceConnectionHelper.ConnectInParallelAsync(devices)
                );

                DmmComboBox.IsEnabled = false;
                OscComboBox.IsEnabled = false;
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

            _instDmm.ResetProperties();
            _instOsc.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
            DmmComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyCheckBox.IsChecked = false;
        }

        // OSC測定値取得
        private async Task<decimal> ReadOsc(OscInstClass oscInstClass, int meas) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var output = await InstrumentService.ReadOscAsync(oscInstClass, meas);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
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
        // DMMローテーション
        private async Task SwitchDmm(DmmInstClass dmmInstClass, bool isNext) {

            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var settings = _dicCommands[dmmInstClass].Settings;
                dmmInstClass.SettingNumber = (dmmInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[dmmInstClass.SettingNumber];
                dmmInstClass.InstCommand = dmmInstClass.SignalType switch {
                    1 => sw.Adc,
                    2 => sw.Visa,
                    3 => sw.Gpib,
                    _ => string.Empty,
                };
                dmmInstClass.Query = sw.Query;

                await DeviceController.ConnectAsync(dmmInstClass);

            } finally {
                VisibleProgressImage(false);
            }
        }

        // OSC mes2値コピー
        private async void ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            var output = await ReadOsc(_instOsc, 2);
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry(output.ToString("0.00"));
            sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }
        // OSC mes3値コピー
        private async void ActionHotkeyBackslash() {
            if (MainWindow.IsProcessing) { return; }

            var output = await ReadOsc(_instOsc, 3);
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry((output * 1000).ToString("0"));
            sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }
        // DMM測定値コピー
        private async void ActionHotkeyPeriod() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm);

                var settings = _dicCommands[_instDmm].Settings;
                var sw = settings[_instDmm.SettingNumber];
                var outputValue = sw.Mode switch {
                    DmmMode.DCI => output * 1000,
                    DmmMode.DCV => output,
                    _ => output,
                };

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(outputValue.ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private async void ActionHotkeyNumMultiply() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm);

                var settings = _dicCommands[_instDmm].Settings;
                var sw = settings[_instDmm.SettingNumber];
                var outputValue = sw.Mode switch {
                    DmmMode.DCI => output * 1000,
                    DmmMode.DCV => output,
                    _ => output,
                };

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(outputValue.ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM切り替え
        private async void ActionHotkeyComma() {
            if (MainWindow.IsProcessing) { return; }

            try {

                await SwitchDmm(_instDmm, true);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private async void ActionHotkeyNumDivide() {
            if (MainWindow.IsProcessing) { return; }

            try {

                await SwitchDmm(_instDmm, true);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // HotKeyの登録
        private void SetHotKey() {
            ThrowIfDisposed();

            MainWindow.HotkeysList.Clear();

            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                    new(ModNone, HotkeyComma, ActionHotkeyComma),
                    new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
                    new(ModNone, HotkeyNumMultiply, ActionHotkeyNumMultiply),
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
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
