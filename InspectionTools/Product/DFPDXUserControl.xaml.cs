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
    /// DFPDXUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class DFPDXUserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DmmInstClass _instDmm01 = new();
        private readonly DmmInstClass _instDmm02 = new();
        private readonly DmmInstClass _instDmm03 = new();
        private readonly FgInstClass _instFg = new();
        private readonly OscInstClass _instOsc = new();

        private record SwitchCommand {
            public string Text { get; init; } = string.Empty;
            public string Adc { get; init; } = string.Empty;
            public string Visa { get; init; } = string.Empty;
            public string Gpib { get; init; } = string.Empty;
            public bool Query { get; init; } = false;
        }
        private readonly Dictionary<InstClass, (SwitchCommand Init, List<SwitchCommand> Settings)> _dicCommands = [];

        public DFPDXUserControl() {
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
                    DisposeInstrument(_instDmm01);
                    DisposeInstrument(_instDmm02);
                    DisposeInstrument(_instDmm03);
                    DisposeInstrument(_instFg);
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
        ~DFPDXUserControl() {
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
            MainWindow.UpdateComboBox(Dmm01ComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(Dmm02ComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(Dmm03ComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(FgComboBox, "ファンクションジェネレータ", [2]);
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
            MainWindow.GetVisaAddress(_instDmm01, Dmm01ComboBox);
            MainWindow.GetVisaAddress(_instDmm02, Dmm02ComboBox);
            MainWindow.GetVisaAddress(_instDmm03, Dmm03ComboBox);
            MainWindow.GetVisaAddress(_instFg, FgComboBox);
            MainWindow.GetVisaAddress(_instOsc, OscComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {

            _dicCommands[_instDmm01] =
                (
                    Init: new() { Adc = "*RST,F5,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 0.2;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instDmm02] =
                (
                    Init: new() { Adc = "*RST,F5,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 0.2;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instDmm03] =
                (
                    Init: new() { Adc = "*RST,F1,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG 20;*OPC?", Query = true },
                    Settings: []
                );
            _dicCommands[_instFg] =
                (
                    Init: new() { Visa = "*RST;:OUTPUT OFF;:FREQ 4000;:VOLT 0.44VPP;*OPC?", Query = true },
                    Settings: [
                        new() { Text = "OFF",   Visa = ":FREQ 4000;:OUTPUT OFF;*OPC?", Query = true },
                        new() { Text = "4000",  Visa = ":FREQ 4000;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "OFF",   Visa = ":OUTPUT OFF;*OPC?", Query = true },
                        new() { Text = "5",     Visa = ":OUTPUT OFF;:FREQ 5;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "100",   Visa = ":OUTPUT OFF;:FREQ 100;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "4000",  Visa = ":OUTPUT OFF;:FREQ 4000;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "6250",  Visa = ":OUTPUT OFF;:FREQ 6250;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "4000",  Visa = ":OUTPUT OFF;:FREQ 4000;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "100",   Visa = ":OUTPUT OFF;:FREQ 100;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "60",    Visa = ":OUTPUT OFF;:FREQ 60;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "72",    Visa = ":OUTPUT OFF;:FREQ 72;:OUTPUT ON;*OPC?", Query = true },
                    ]
                );

            _dicCommands[_instOsc] =
                (
                    Init: new() {
                        Visa =
                            """
                            *RST;:HEADER 0;
                            :ACQUIRE:MODE SAMPLE;NUMAVG 16;
                            :CH1:SCALE 1.0E-1;COUPLING DC;
                            :CURSOR:FUNCTION HBARS;SELECT:SOURCE CH1;:CURSOR:HBARS:POSITION1 2.0E-1;POSITION2 -2.0E-1;
                            :HORIZONTAL:MAIN:SCALE 1.0E-3;
                            :TRIGGER:MAIN:LEVEL 0.0E0;
                            :MEASUREMENT:MEAS1:TYPE MAXIMUM;SOURCE CH1;
                            :MEASUREMENT:MEAS2:TYPE MINIMUM;SOURCE CH1;
                            :MEASUREMENT:MEAS3:TYPE NWIDTH;SOURCE CH1;
                            *OPC?
                            """,
                        Query = true
                    },
                    Settings: [
                        new() { Text = "1[6]",      Visa = ":HORIZONTAL:MAIN:SCALE 1.0E-3;:CH1:SCALE 1.0E-1;:TRIGGER:MAIN:LEVEL 0.0E0;*OPC?", Query = true },
                        new() { Text = "2[7-1]",    Visa = ":HORIZONTAL:MAIN:SCALE 5.0E-2;*OPC?", Query = true },
                        new() { Text = "3[7-2]",    Visa = ":HORIZONTAL:MAIN:SCALE 2.5E-3;*OPC?", Query = true },
                        new() { Text = "4[7-3]",    Visa = ":HORIZONTAL:MAIN:SCALE 1.0E-4;:CH1:SCALE 1.0E-1;*OPC?", Query = true },
                        new() { Text = "5[8]",      Visa = ":HORIZONTAL:MAIN:SCALE 1.0E-4;:CH1:SCALE 5.0E-1;*OPC?", Query = true },
                        new() { Text = "6[9]",      Visa = ":HORIZONTAL:MAIN:SCALE 2.5E-3;:CH1:SCALE 2.0E0;:TRIGGER:MAIN:LEVEL 0.0E0;*OPC?", Query = true },
                        new() { Text = "7[10]",     Visa = ":HORIZONTAL:MAIN:SCALE 2.5E-4;:CH1:SCALE 2.0E0;:TRIGGER:MAIN:LEVEL 1.2E0;*OPC?", Query = true },
                        new() { Text = "8[11]",     Visa = ":HORIZONTAL:MAIN:SCALE 5.0E-4;:CH1:SCALE 1.0E0;:TRIGGER:MAIN:LEVEL 1.2E0;*OPC?", Query = true },
                    ]
                );
        }
        // 機器初期設定
        private void FormatSet() {
            (_instDmm01.InstCommand, _instDmm01.Query) = ResolveCommand(_dicCommands[_instDmm01].Init, _instDmm01.SignalType);
            (_instDmm02.InstCommand, _instDmm02.Query) = ResolveCommand(_dicCommands[_instDmm02].Init, _instDmm02.SignalType);
            (_instDmm03.InstCommand, _instDmm03.Query) = ResolveCommand(_dicCommands[_instDmm03].Init, _instDmm03.SignalType);
            (_instFg.InstCommand, _instFg.Query) = ResolveCommand(_dicCommands[_instFg].Init, _instFg.SignalType);
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
                ValidateDmmSelection();

                RegDictionary();
                FormatSet();

                await Task.Run(async () => {
                    InstClass[] devices = [_instDmm01, _instDmm02, _instDmm03, _instFg, _instOsc];
                    var tasks = devices.Select(async device => {
                        try {
                            await MainWindow.ConnectDeviceAsync(device);
                        } catch (Exception ex) {
                            throw new Exception($"[{device.Name}] 接続失敗: {ex.Message}", ex);
                        }
                    });
                    var whenAllTask = Task.WhenAll(tasks);
                    try {
                        await whenAllTask;
                    } catch {
                        if (whenAllTask.Exception != null)
                            throw whenAllTask.Exception;
                        throw;
                    }
                });

                if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                    FgRotateRangeTextBox.Text = "OFF";
                    FgRotateButton.IsEnabled = true;
                }
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                    OscRotateRangeTextBox.Text = "1";
                    OscRotateButton.IsEnabled = true;
                }

                Dmm01ComboBox.IsEnabled = false;
                Dmm02ComboBox.IsEnabled = false;
                Dmm03ComboBox.IsEnabled = false;
                FgComboBox.IsEnabled = false;
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
        // DMMのIDチェック処理
        private void ValidateDmmSelection() {
            var indices = new[] { _instDmm01.Index, _instDmm02.Index, _instDmm03.Index }
                .Where(i => i >= 1); // 未選択(0以下)は無視

            if (indices.Count() != indices.Distinct().Count()) {
                throw new Exception("同じ測定器が選択されています。");
            }
        }

        // 解除
        private void Release() {
            VisibleProgressImage(false);

            _instDmm01.ResetProperties();
            _instDmm02.ResetProperties();
            _instDmm03.ResetProperties();
            _instFg.ResetProperties();
            _instOsc.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
            Dmm01ComboBox.IsEnabled = true;
            Dmm02ComboBox.IsEnabled = true;
            Dmm03ComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyCheckBox.IsChecked = false;
            FgRotateButton.IsEnabled = false;
            OscRotateButton.IsEnabled = false;
            FgRotateRangeTextBox.Text = string.Empty;
            OscRotateRangeTextBox.Text = string.Empty;
        }

        // DMM測定値取得
        private async Task<decimal> ReadDmm(DmmInstClass dmmInstClass) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var output = await MainWindow.ReadDmm(dmmInstClass);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // OSC測定値取得
        private async Task<decimal> ReadOsc(OscInstClass oscInstClass, int meas) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var output = await MainWindow.ReadOsc(oscInstClass, meas);

                return output;

            } finally {
                VisibleProgressImage(false);
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
                fgInstClass.InstCommand = fgInstClass.SignalType switch {
                    1 => sw.Adc,
                    2 => sw.Visa,
                    3 => sw.Gpib,
                    _ => string.Empty,
                };
                fgInstClass.Query = sw.Query;

                if (fgInstClass.InstCommand == string.Empty) { return; }

                await MainWindow.RotationFgAsync(fgInstClass);

                FgRotateRangeTextBox.Text = sw.Text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        // OSC切り替え
        private async void RotationOsc(OscInstClass oscInstClass, bool isNext) {
            ThrowIfDisposed();

            try {
                if (string.IsNullOrEmpty(oscInstClass.VisaAddress)) { return; }
                VisibleProgressImage(true);

                var settings = _dicCommands[oscInstClass].Settings;
                oscInstClass.SettingNumber = (oscInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[oscInstClass.SettingNumber];
                oscInstClass.InstCommand = oscInstClass.SignalType switch {
                    1 => sw.Adc,
                    2 => sw.Visa,
                    3 => sw.Gpib,
                    _ => string.Empty,
                };
                oscInstClass.Query = sw.Query;

                if (oscInstClass.InstCommand == string.Empty) { return; }

                await MainWindow.RotationOscAsync(oscInstClass);

                OscRotateRangeTextBox.Text = sw.Text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }

        // FGローテーション
        private void ActionHotkeyAtsign() {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, true);
        }
        private void ActionHotkeyShiftAtsign() {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, false);
        }
        // DMM01測定値コピー
        private async void ActionHotkeyPeriod() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm01);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM02測定値コピー
        private async void ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm02);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM03測定値コピー
        private async void ActionHotkeyBackslash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm03);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(output.ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSCローテーション
        private void ActionHotkeyBracketL() {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, true);
        }
        private void ActionHotkeyShiftBracketL() {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, false);
        }
        // OSC meas1測定値コピー
        private async void ActionHotkeyBracketR() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadOsc(_instOsc, 1);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // HotKeyの登録
        private void SetHotKey() {
            ThrowIfDisposed();

            MainWindow.HotkeysList.Clear();

            if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm03.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                ]);
            }
            if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                    new(ModShift, HotkeyAtsign, ActionHotkeyShiftAtsign)
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketL, ActionHotkeyBracketL),
                    new(ModShift, HotkeyBracketL, ActionHotkeyShiftBracketL),
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR)
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

        private void FgRotateButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, true);
        }
        private void OscRotateButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, true);
        }

        private void UserControl_Unloaded(object sender, RoutedEventArgs e) { Dispose(); }

    }
}
