using InspectionTools.Common;
using static InspectionTools.Common.InstrumentHelper;
using System.Data;
using System.Windows;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;
using static InspectionTools.MainWindow;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.Product {
    /// <summary>
    /// EL0137UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL0137UserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        // MainWindowへの参照をセットする
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DmmInstClass _instDmm = new();
        private readonly OscInstClass _instOsc = new();
        private readonly InputSimulator _sim = new();

        private readonly Dictionary<InstClass, (SwitchCommand Init, List<SwitchCommand> Settings)> _dicCommands = [];

        public EL0137UserControl() {
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
                    InstrumentHelper.SafeDispose(_instDmm);
                    InstrumentHelper.SafeDispose(_instOsc);

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
        /// オブジェクトが破棄済みかチェック
        /// </summary>
        private void ThrowIfDisposed() {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        /// <summary>
        /// ファイナライザ
        /// </summary>
        ~EL0137UserControl() {
            Dispose(false);
        }

        #endregion

        // UserControl読み込み時に計測器一覧を更新してウィンドウサイズを調整する
        private void LoadEvents() {
            ThrowIfDisposed();
            InstListImport();
            var parentWindow = Window.GetWindow(this);
            MainWindow.AdjustWindowSizeToUserControl(parentWindow);
        }
        // 計測器カテゴリ別にコンボボックスのアイテムを更新する
        private void InstListImport() {
            // デジタルマルチメータ、ファンクションジェネレータ、オシロスコープのコンボボックスを更新する
            MainWindow.UpdateComboBox(DmmComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(OscComboBox, "オシロスコープ", [2]);
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            MainGrid.IsEnabled = !isVisible;
            _mainWindow?.ShowSpinner(isVisible);
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
                    Init: new() { DmmMode = DmmMode.DCI, Adc = "*RST,F5,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instOsc] =
                (
                    Init: new() {
                        Visa =
                            """
                            *RST;
                            :HEADER 0;
                            :CH1:SCALE 5.0E-1;POSITION -2.0E0;
                            :CH2:POSITION -2.0E0;
                            :CH3:POSITION -2.0E0;
                            :CURSOR:FUNCTION HBARS;SELECT:SOURCE CH1;:CURSOR:HBARS:POSITION1 1.5E0;POSITION2 1.5E0;
                            :HORIZONTAL:MAIN:SCALE 2.5E-3;
                            :TRIGGER:MAIN:LEVEL 1.0E0;
                            :MEASUREMENT:MEAS1:TYPE MAXIMUM;SOURCE CH1;
                            :MEASUREMENT:MEAS4:TYPE NONE;SOURCE MATH;
                            :MEASUREMENT:MEAS5:TYPE NONE;SOURCE MATH;
                            *OPC?
                            """,
                        Query = true
                    },
                    Settings: [
                        new() {
                            Text ="OC入力回路",
                            Visa =
                                """
                                :SELECT:CH1 1;CH2 0;
                                :CH1:SCALE 5.0E-1;POSITION -2.0E0;
                                :CURSOR:FUNCTION HBARS;SELECT:SOURCE CH1;
                                :CURSOR:HBARS:POSITION1 1.5E0;POSITION2 1.5E0;
                                :HORIZONTAL:MAIN:SCALE 2.5E-3;
                                :TRIGGER:MAIN:LEVEL 1.0E0;EDGE:SOURCE CH1;
                                :MEASUREMENT:MEAS1:TYPE MAXIMUM;SOURCE CH1;
                                :MEASUREMENT:MEAS2:TYPE NONE;SOURCE MATH;
                                :MEASUREMENT:MEAS3:TYPE NONE;SOURCE MATH;
                                *OPC?
                                """,
                            Query = true
                        },
                        new() {
                            Text ="電流入力回路",
                            Visa =
                                """
                                :CURSOR:HBARS:POSITION1 1.6E0;POSITION2 1.6E0;
                                *OPC?
                                """,
                            Query = true
                        },
                        new() {
                            Text ="電圧入力回路",
                            Visa =
                                """
                                :CURSOR:HBARS:POSITION1 1.3E0;POSITION2 1.3E0;
                                *OPC?
                                """,
                            Query = true
                        },
                        new() {
                            Text ="波高値",
                            Visa =
                                """
                                :CH1:SCALE 1.0E0;
                                :TRIGGER:MAIN:LEVEL 1.0E0;EDGE:SOURCE CH1;
                                :CURSOR:HBARS:POSITION1 3.0E0;POSITION2 2.0E0;
                                :MEASUREMENT:MEAS2:TYPE MINIMUM;SOURCE CH1;
                                *OPC?
                                """,
                            Query = true
                        },
                        new() {
                            Text ="1/1分周 CH1,2",
                            Visa =
                                """
                                :SELECT:CH1 1;CH2 1;
                                :CURSOR:SELECT:SOURCE CH1;
                                :MEASUREMENT:MEAS1:TYPE NWIDTH;SOURCE CH1;
                                :MEASUREMENT:MEAS2:TYPE PWIDTH;SOURCE CH2;
                                *OPC?
                                """,
                            Query = true
                        },
                        new() {
                            Text ="1/2~100分周 CH1,2",
                            Visa =
                                """
                                :SELECT:CH1 1;CH2 1;
                                :CURSOR:SELECT:SOURCE CH1;
                                :MEASUREMENT:MEAS1:TYPE PWIDTH;SOURCE CH1;
                                :MEASUREMENT:MEAS2:TYPE PWIDTH;SOURCE CH2;
                                *OPC?
                                """,
                            Query = true
                        },
                    ]
                );
        }
        // 機器初期設定
        private void FormatSet() {
            (_instDmm.InstCommand, _instDmm.Query) = ResolveCommand(_dicCommands[_instDmm].Init, _instDmm.SignalType);
            (_instOsc.InstCommand, _instOsc.Query) = ResolveCommand(_dicCommands[_instOsc].Init, _instOsc.SignalType);
        }
        // 機器接続
        private async Task ConnectInstAsync() {
            ThrowIfDisposed();

            try {
                _mainWindow?.SetButtonEnabled(ProductListButtonName, false);

                HotKeyCheckBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();

                RegDictionary();
                FormatSet();

                InstClass[] devices = [_instDmm, _instOsc];

                await DeviceConnectionHelper.ConnectInParallelAsync(devices);

                if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                    _instDmm.CurrentMode = _dicCommands[_instDmm].Init.DmmMode;
                }
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                    OscRotateRangeTextBox.Text = "OC入力回路";
                    OscRotateButton.IsEnabled = true;
                }

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

            _mainWindow?.SetButtonEnabled(ProductListButtonName, true);
            DmmComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyCheckBox.IsChecked = false;
            OscRotateButton.IsEnabled = false;
            OscRotateRangeTextBox.Text = string.Empty;
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
        // OSC切り替え
        private async Task RotationOsc(OscInstClass oscInstClass, bool isNext) {
            if (_disposed) return;

            try {
                if (string.IsNullOrEmpty(oscInstClass.VisaAddress)) { return; }
                VisibleProgressImage(true);

                var settings = _dicCommands[oscInstClass].Settings;
                oscInstClass.SettingNumber = (oscInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[oscInstClass.SettingNumber];
                (oscInstClass.InstCommand, oscInstClass.Query) = ResolveCommand(sw, oscInstClass.SignalType);

                if (oscInstClass.InstCommand == string.Empty) { return; }

                await InstrumentService.RotateOscAsync(oscInstClass);

                OscRotateRangeTextBox.Text = sw.Text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }

        // DMM01測定値コピー
        private async Task ActionHotkeyColon() => await ReadDmmAndSendAsync();
        private async Task ActionHotkeyNumDivide() => await ReadDmmAndSendAsync();

        // DMM測定値をuA単位に変換してキーボード入力としてEnterまで送信する
        private async Task ReadDmmAndSendAsync() {
            if (MainWindow.IsProcessing) { return; }
            try {
                var output = await ReadDmm(_instDmm);

                _sim.Keyboard.TextEntry((output * 1000000).ToString("0.000"));
                await Task.Delay(100);
                _sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSCローテーション
        private async Task ActionHotkeyBracketR() { if (MainWindow.IsProcessing) { return; } await RotationOsc(_instOsc, true); }
        private async Task ActionHotkeyNumMultiply() { if (MainWindow.IsProcessing) { return; } await RotationOsc(_instOsc, true); }
        // OSC meas1測定値コピー
        private async Task ActionHotkeySlash() => await ReadOscAndSendAsync(1);
        private async Task ActionHotkeyNumSubtract() => await ReadOscAndSendAsync(1);
        // OSC meas2測定値コピー
        private async Task ActionHotkeyBackslash() => await ReadOscAndSendAsync(2);
        private async Task ActionHotkeyNumAdd() => await ReadOscAndSendAsync(2);

        // OSC測定値をms単位に変換してキーボード入力としてEnterまで送信する
        private async Task ReadOscAndSendAsync(int meas) {
            if (MainWindow.IsProcessing) { return; }
            try {
                var output = await ReadOsc(_instOsc, meas);
                _sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
                await Task.Delay(100);
                _sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
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
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                    new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                    new(ModNone, HotkeyNumMultiply, ActionHotkeyNumMultiply),
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyNumSubtract, ActionHotkeyNumSubtract),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
                ]);
            }

            MainWindow.SetHotKey();
        }
        // 登録済みホットキーを解除する
        private static void ClearHotKey() {
            MainWindow.ClearHotKey();
        }

        // イベントハンドラ
        private void UserControl_Loaded(object sender, RoutedEventArgs e) { LoadEvents(); }
        private async void ConnectButton_Click(object sender, RoutedEventArgs e) { await ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void HotKeyCheckBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyCheckBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }

        private async void OscRotateButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            await RotationOsc(_instOsc, true);
        }

        private void UserControl_Unloaded(object sender, RoutedEventArgs e) { Dispose(); }

    }
}
