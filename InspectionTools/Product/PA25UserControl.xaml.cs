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
    /// PA25UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class PA25UserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DmmInstClass _instDmm = new();
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
        private readonly Dictionary<InstClass, (SwitchCommand Init, List<SwitchCommand> Settings)> _dicReverseCommands = [];

        public PA25UserControl() {
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
        ~PA25UserControl() {
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
            MainWindow.GetVisaAddress(_instDmm, DmmComboBox);
            MainWindow.GetVisaAddress(_instFg, FgComboBox);
            MainWindow.GetVisaAddress(_instOsc, OscComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicCommands[_instDmm] =
                (
                    Init: new() { Adc = "*RST,F1,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG 2;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instFg] =
                (
                    Init: new() { Visa = "*RST;:FREQ 1;:VOLT 0.44VPP;*OPC?", Query = true },
                    Settings: [
                        new() { Text = "OFF",   Visa = ":FREQ 1;:OUTPUT OFF;*OPC?", Query = true },
                        new() { Text = "27",    Visa = ":OUTPUT OFF;:FREQ 27;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "29",    Visa = ":OUTPUT OFF;:FREQ 29;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "400",   Visa = ":OUTPUT OFF;:FREQ 400;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "1",     Visa = ":OUTPUT OFF;:FREQ 1;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "10",    Visa = ":OUTPUT OFF;:FREQ 10;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "200",   Visa = ":OUTPUT OFF;:FREQ 200;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "OFF",   Visa = ":OUTPUT OFF;:FREQ 1;*OPC?", Query = true },
                        new() { Text = "2200",  Visa = ":OUTPUT OFF;:FREQ 2200;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "1000",  Visa = ":OUTPUT OFF;:FREQ 1000;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "OFF",   Visa = ":OUTPUT OFF;:FREQ 1;*OPC?", Query = true },
                        new() { Text = "1000",  Visa = ":OUTPUT OFF;:FREQ 1000;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "OFF",   Visa = ":OUTPUT OFF;:FREQ 1;*OPC?", Query = true },
                        new() { Text = "6250",  Visa = ":OUTPUT OFF;:FREQ 6250;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "3800",  Visa = ":OUTPUT OFF;:FREQ 3800;:OUTPUT ON;*OPC?", Query = true },
                        new() { Text = "1000",  Visa = ":OUTPUT OFF;:FREQ 1000;:OUTPUT ON;*OPC?", Query = true },
                    ]
                );

            _dicCommands[_instOsc] =
                (
                    Init: new() {
                        Visa =
                            """
                            *RST;:HEADER 0;
                            :CH1:SCALE 1.0E-1;
                            :TRIGGER:MAIN:LEVEL 3.0E-1;
                            :CURSOR:FUNCTION VBArs;SELECT:SOURCE CH1;:CURSOR:VBArs:POSITION1 -1.34E-3;POSITION2 1.16E-3;
                            :MEASUREMENT:MEAS1:TYPE PK2PK;SOURCE CH1;
                            :MEASUREMENT:MEAS1:TYPE NONE;SOURCE CH1;
                            *OPC?
                            """,
                        Query = true
                    },
                    Settings: [
                        new() {
                            Text = "1",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 5.0E-4;
                                :CH1:SCALE 1.0E-1;COUPLING DC;
                                :TRIGGER:MAIN:LEVEL 3.0E-1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "2",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 2.5E-1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "3",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 2.5E-2;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "4",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 1.0E-3;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "5",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 2.5E-4;
                                :CH1:SCALE 2.0E0;
                                :TRIGGER:MAIN:LEVEL 3.84E0;
                                :MEASUREMENT:MEAS1:TYPE PERIOD;SOURCE CH1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "6",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 5.0E-5;
                                :CH1:SCALE 1.0E1;
                                :MEASUREMENT:MEAS1:TYPE NWIDTH;SOURCE CH1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "7",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 5.0E-2;
                                :CH1:SCALE 2.0E0;
                                :MEASUREMENT:MEAS1:TYPE PWIDTH;SOURCE CH1;
                                :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "8",
                            Visa =
                                """                        
                                :MEASUREMENT:MEAS1:TYPE MINIMUM;SOURCE CH1;
                                :MEASUREMENT:MEAS2:TYPE NONE;SOURCE CH1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "9",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 1.0E-2;                        
                                :CH1:SCALE 2.0E-2;COUPLING AC;
                                :TRIGGER:MAIN:LEVEL 1.69E-1;
                                :MEASUREMENT:MEAS1:TYPE NONE;SOURCE CH1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "10",
                            Visa =
                                """
                                :CH1:COUPLING DC;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "11",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 2.50E-4;  
                                :CH1:SCALE 1.0E0;
                                :TRIGGER:MAIN:LEVEL 3.84E0;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "12",
                            Visa =
                                """
                                :CH1:SCALE 1.0E-1;COUPLING AC;
                                :TRIGGER:MAIN:LEVEL 8.44E-1;
                                *OPC?
                                """,
                            Query = true },
                    ]
                );

            _dicReverseCommands[_instOsc] =
                (
                    Init: new(),
                    Settings: [
                        new() {
                            Text = "1",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 5.0E-4;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "2",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 2.5E-1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "3",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 2.5E-2;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "4",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 1.0E-3;
                                :CH1:SCALE 1.0E-1;COUPLING DC;
                                :TRIGGER:MAIN:LEVEL 3.0E-1;
                                :MEASUREMENT:MEAS1:TYPE NONE;SOURCE CH1;                        
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "5",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 2.5E-4;
                                :CH1:SCALE 2.0E0;
                                :MEASUREMENT:MEAS1:TYPE PERIOD;SOURCE CH1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "6",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 5.0E-5;
                                :CH1:SCALE 1.0E1;
                                :MEASUREMENT:MEAS1:TYPE NWIDTH;SOURCE CH1;
                                :MEASUREMENT:MEAS2:TYPE NONE;SOURCE CH1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "7",
                            Visa =
                                """
                                :MEASUREMENT:MEAS1:TYPE PWIDTH;SOURCE CH1;
                                :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "8",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 5.0E-2;
                                :CH1:SCALE 2.0E0;
                                :TRIGGER:MAIN:LEVEL 3.84E0;
                                :MEASUREMENT:MEAS1:TYPE MINIMUM;SOURCE CH1;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "9",
                            Visa =
                                """
                                :CH1:SCALE 2.0E-2;COUPLING AC;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "10",
                            Visa =
                                """
                                :CH1:COUPLING DC;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "11",
                            Visa =
                                """
                                :CH1:SCALE 1.0E0;
                                :TRIGGER:MAIN:LEVEL 3.84E0;
                                *OPC?
                                """,
                            Query = true },
                        new() {
                            Text = "12",
                            Visa =
                                """
                                :HORIZONTAL:MAIN:SCALE 2.50E-4;
                                :CH1:SCALE 1.0E-1;COUPLING AC;
                                :TRIGGER:MAIN:LEVEL 8.44E-1;
                                *OPC?
                                """,
                            Query = true },
                    ]
                );
        }

        // 機器初期設定
        private void FormatSet() {
            (_instDmm.InstCommand, _instDmm.Query) = ResolveCommand(_dicCommands[_instDmm].Init, _instDmm.SignalType);
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

                InstClass[] devices = [_instDmm, _instFg, _instOsc];
                RegDictionary();
                FormatSet();
                var tasks = devices.Select(device => MainWindow.ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);

                if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                    _instDmm.CurrentMode = DmmMode.DCV;
                }
                if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                    FgRotateRangeTextBox.Text = "OFF";
                    FgRotateButton.IsEnabled = true;
                    FgRotateRButton.IsEnabled = true;
                }
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                    OscRotateRangeTextBox.Text = "1";
                    OscRotateButton.IsEnabled = true;
                    OscRotateRButton.IsEnabled = true;
                }

                DmmComboBox.IsEnabled = false;
                FgComboBox.IsEnabled = false;
                OscComboBox.IsEnabled = false;
                ConnectButton.IsEnabled = false;
                ReleaseButton.IsEnabled = true;

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

            _instFg.ResetProperties();
            _instOsc.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
            DmmComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyCheckBox.IsChecked = false;
            FgRotateButton.IsEnabled = false;
            FgRotateRButton.IsEnabled = false;
            OscRotateButton.IsEnabled = false;
            OscRotateRButton.IsEnabled = false;
            FgRotateRangeTextBox.Text = string.Empty;
            OscRotateRangeTextBox.Text = string.Empty;
        }

        // DMM切り替え
        private async Task SwitchDmm(DmmInstClass dmmInstClass, string func) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                (dmmInstClass.InstCommand, dmmInstClass.CurrentMode) = func switch {
                    "DCI" => dmmInstClass.SignalType switch {
                        1 => ("*RST,F5,R6,*OPC?", DmmMode.DCI),
                        2 => ("*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 1E-1;*OPC?", DmmMode.DCI),
                        _ => throw new ApplicationException(),
                    },
                    "DCV" => dmmInstClass.SignalType switch {
                        1 => ("*RST,F1,R6,*OPC?", DmmMode.DCV),
                        2 => ("*RST;:INIT:CONT 1;:VOLT:DC:RANG 10;*OPC?", DmmMode.DCV),
                        _ => throw new ApplicationException(),
                    },
                    _ => throw new ApplicationException(),
                };

                await MainWindow.ConnectDeviceAsync(dmmInstClass);

            } finally {
                VisibleProgressImage(false);
            }
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

                var dic = isNext ? _dicCommands : _dicReverseCommands;
                var settings = dic[oscInstClass].Settings;
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

        // DMM切替(DCV)
        private async void ActionHotkeyComma() {
            if (MainWindow.IsProcessing) { return; }

            try {
                await SwitchDmm(_instDmm, "DCV");
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM切替(DCI)
        private async void ActionHotkeyPeriod() {
            if (MainWindow.IsProcessing) { return; }

            try {
                await SwitchDmm(_instDmm, "DCI");
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM01測定値コピー
        private async void ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm);

                var value = _instDmm.CurrentMode switch {
                    DmmMode.DCV => output.ToString("0.0000"),
                    DmmMode.DCI => (output * 1000).ToString("0.0000"),
                    _ => output.ToString(""),
                };

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(value);
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // FGローテーション
        private void ActionHotkeyBracketR() {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, true);
        }
        private void ActionHotkeyShiftBracketR() {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, false);
        }
        private void ActionHotkeyNumMultiply() {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, true);
        }

        // OSCローテーション
        private void ActionHotkeyColon() {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, true);
        }
        private void ActionHotkeyShiftColon() {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, false);
        }
        private void ActionHotkeyNumDivide() {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, true);
        }

        // OSC meas測定値コピー
        private async void ActionHotkeyBackslash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var meas = _instOsc.SettingNumber switch {
                    4 or 5 or 7 => 1,
                    6 => 2,
                    _ => 0,
                };
                if (meas == 0) { return; }

                for (var i = 1; i <= meas; i++) {
                    var output = await ReadOsc(_instOsc, i);

                    var sim = new InputSimulator();
                    var value = _instOsc.SettingNumber switch {
                        4 => output * 1000,
                        5 => output * 1000000,
                        7 => output,
                        6 => i switch {
                            1 => output * 1000,
                            2 => output,
                            _ => output,
                        },
                        _ => output,
                    };
                    sim.Keyboard.TextEntry(value.ToString("0.000"));
                    await Task.Delay(100);
                    sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                }

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private async void ActionHotkeyNumAdd() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var meas = _instOsc.SettingNumber switch {
                    4 or 5 or 7 => 1,
                    6 => 2,
                    _ => 0,
                };
                if (meas == 0) { return; }

                for (var i = 1; i <= meas; i++) {
                    var output = await ReadOsc(_instOsc, i);

                    var sim = new InputSimulator();
                    var value = _instOsc.SettingNumber switch {
                        4 => output * 1000,
                        5 => output * 1000000,
                        7 => output,
                        6 => i switch {
                            1 => output * 1000,
                            2 => output,
                            _ => output,
                        },
                        _ => output,
                    };
                    sim.Keyboard.TextEntry(value.ToString("0.000"));
                    await Task.Delay(100);
                    sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                }

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
                    new(ModNone, HotkeyComma, ActionHotkeyComma),
                    new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                ]);
            }
            if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                    new(ModShift, HotkeyBracketR, ActionHotkeyShiftBracketR),
                    new(ModNone, HotkeyNumMultiply, ActionHotkeyNumMultiply),
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                    new(ModShift, HotkeyColon, ActionHotkeyShiftColon),
                    new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
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
        private void FgRotateRButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, false);
        }
        private void OscRotateButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, true);
        }
        private void OscRotateRButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, false);
        }
        private void UserControl_Unloaded(object sender, RoutedEventArgs e) { Dispose(); }

    }
}
