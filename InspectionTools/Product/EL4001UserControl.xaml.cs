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
    /// EL4001UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL4001UserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        // MainWindowへの参照をセットする
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DcsInstClass _instDcs = new();
        private readonly DmmInstClass _instDmm01 = new();
        private readonly DmmInstClass _instDmm02 = new();
        private readonly OscInstClass _instOsc = new();

        private readonly Dictionary<InstClass, (SwitchCommand Init, List<SwitchCommand> Settings)> _dicCommands = [];

        public EL4001UserControl() {
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
                    InstrumentHelper.SafeDispose(_instDcs);
                    InstrumentHelper.SafeDispose(_instDmm01);
                    InstrumentHelper.SafeDispose(_instDmm02);
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
        ~EL4001UserControl() {
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
            MainWindow.UpdateComboBox(DcsComboBox, "電流電圧発生器", [2, 3]);
            MainWindow.UpdateComboBox(Dmm01ComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(Dmm02ComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(OscComboBox, "オシロスコープ", [2]);
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            MainGrid.IsEnabled = !isVisible;
            _mainWindow?.ShowSpinner(isVisible);
        }

        // 選択した機器のVisaAddressを取得
        private void SelectInst() {
            MainWindow.GetVisaAddress(_instDcs, DcsComboBox);
            MainWindow.GetVisaAddress(_instDmm01, Dmm01ComboBox);
            MainWindow.GetVisaAddress(_instDmm02, Dmm02ComboBox);
            MainWindow.GetVisaAddress(_instOsc, OscComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicCommands[_instDcs] =
                (
                    Init: new() { DcsMode = DcsMode.Off, Visa = "SIR3,SOI+0,SBY", Gpib = "RCF5R6S0EO0E" },
                    Settings: [
                        new() { DcsMode = DcsMode.Off,  Text = "OFF",   Visa = "SOI+0MA,SBY",   Gpib = "F5R6S0EO0E" },
                        new() { DcsMode = DcsMode.On,   Text = "4.0mA", Visa = "SOI+4MA,OPR",   Gpib = "F5R6S4.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "20mA",  Visa = "SOI+20MA,OPR",  Gpib = "F5R6S20.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "4.0mA", Visa = "SOI+4MA,OPR",   Gpib = "F5R6S4.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "20mA",  Visa = "SOI+20MA,OPR",  Gpib = "F5R6S20.0E-3O1E" },
                        new() { DcsMode = DcsMode.Off,  Text = "OFF",   Visa = "SOI+0MA,SBY",   Gpib = "F5R6S0EO0E" },
                        new() { DcsMode = DcsMode.On,   Text = "22mA",  Visa = "SOI+22MA,OPR",  Gpib = "F5R6S22.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "20mA",  Visa = "SOI+20MA,OPR",  Gpib = "F5R6S20.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "12mA",  Visa = "SOI+12MA,OPR",  Gpib = "F5R6S12.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "4.0mA", Visa = "SOI+4MA,OPR",   Gpib = "F5R6S4.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "3.2mA", Visa = "SOI+3.2MA,OPR", Gpib = "F5R6S3.2E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "22mA",  Visa = "SOI+22MA,OPR",  Gpib = "F5R6S22.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "20mA",  Visa = "SOI+20MA,OPR",  Gpib = "F5R6S20.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "12mA",  Visa = "SOI+12MA,OPR",  Gpib = "F5R6S12.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "4.0mA", Visa = "SOI+4MA,OPR",   Gpib = "F5R6S4.0E-3O1E" },
                        new() { DcsMode = DcsMode.On,   Text = "3.2mA", Visa = "SOI+3.2MA,OPR", Gpib = "F5R6S3.2E-3O1E" },
                    ]
                );

            _dicCommands[_instDmm01] =
                (
                    Init: new() { DmmMode = DmmMode.DCV, Adc = "*RST,F1,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG 20;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instDmm02] =
                (
                    Init: new() { DmmMode = DmmMode.DCV, Adc = "*RST,F1,R5,*OPC?", Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG 2;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instOsc] =
                (
                    Init: new() {
                        Visa =
                            """
                            *RST;
                            :HEADER 0;
                            :CH1:PROBE 1.0E1;SCALE 2.0E0;
                            :HORIZONTAL:MAIN:SCALE 5.0E-2;POSITION 1.0E-1;
                            :TRIGGER:MAIN:LEVEL 7.2E-1;
                            :MEASUREMENT:MEAS1:TYPE PWIDTH;SOURCE CH1;
                            :MEASUREMENT:MEAS2:TYPE NWIDTH;SOURCE CH1;
                            :MEASUREMENT:MEAS3:TYPE NONE;SOURCE MATH;
                            :MEASUREMENT:MEAS4:TYPE NONE;SOURCE MATH;
                            :MEASUREMENT:MEAS5:TYPE NONE;SOURCE MATH;
                            *OPC?
                            """,
                        Query = true
                    },
                    Settings: [
                            new() { Text = "50ms", Visa = ":HORIZONTAL:MAIN:SCALE 5.0E-2;POSITION 1.0E-1;*OPC?", Query = true },
                            new() { Text = "500us", Visa = ":HORIZONTAL:MAIN:SCALE 5.0E-4;POSITION 0.0;*OPC?", Query = true },
                        ]
                );
        }
        // 機器初期設定
        private void FormatSet() {
            (_instDcs.InstCommand, _instDcs.Query) = ResolveCommand(_dicCommands[_instDcs].Init, _instDcs.SignalType);
            (_instDmm01.InstCommand, _instDmm01.Query) = ResolveCommand(_dicCommands[_instDmm01].Init, _instDmm01.SignalType);
            (_instDmm02.InstCommand, _instDmm02.Query) = ResolveCommand(_dicCommands[_instDmm02].Init, _instDmm02.SignalType);
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
                ValidateDmmSelection(_instDmm01.Index, _instDmm02.Index);

                RegDictionary();
                FormatSet();

                InstClass[] devices = [_instDcs, _instDmm01, _instDmm02, _instOsc];

                await DeviceConnectionHelper.ConnectInParallelAsync(devices);

                if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                    DcsNumberLabel.Text = "00";
                    DcsRangeLabel.Text = "OFF";
                    _instDcs.CurrentMode = _dicCommands[_instDcs].Init.DcsMode;
                }
                if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                    _instDmm01.CurrentMode = _dicCommands[_instDmm01].Init.DmmMode;
                }
                if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                    _instDmm02.CurrentMode = _dicCommands[_instDmm02].Init.DmmMode;
                }
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) { OscRangeLabel.Text = "50m"; }

                Dmm01ComboBox.IsEnabled = false;
                Dmm02ComboBox.IsEnabled = false;
                OscComboBox.IsEnabled = false;
                DcsComboBox.IsEnabled = false;
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
            _instDmm01.ResetProperties();
            _instDmm02.ResetProperties();
            _instOsc.ResetProperties();

            _mainWindow?.SetButtonEnabled(ProductListButtonName, true);
            Dmm01ComboBox.IsEnabled = true;
            Dmm02ComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            DcsComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyCheckBox.IsChecked = false;
            OscRangeLabel.Text = string.Empty;
            DcsNumberLabel.Text = string.Empty;
            DcsRangeLabel.Text = string.Empty;
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
        // DCSローテーション
        private async Task RotationDcs(DcsInstClass dcsInstClass, bool isNext) {
            if (_disposed) return;

            try {
                VisibleProgressImage(true);

                var settings = _dicCommands[dcsInstClass].Settings;
                dcsInstClass.SettingNumber = (dcsInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[dcsInstClass.SettingNumber];
                (dcsInstClass.InstCommand, dcsInstClass.Query) = ResolveCommand(sw, dcsInstClass.SignalType);

                await DeviceController.ConnectAsync(dcsInstClass);
                DcsNumberLabel.Text = dcsInstClass.SettingNumber.ToString("00");
                DcsRangeLabel.Text = sw.Text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            } finally {
                VisibleProgressImage(false);
            }
        }
        // OSCローテーション
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

                OscRangeLabel.Text = sw.Text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            } finally {
                VisibleProgressImage(false);
            }
        }

        // DMM01測定値コピー
        private async Task ActionHotkeyColon() => await ReadDmm01AndSendAsync();
        private async Task ActionHotkeyNumMultiply() => await ReadDmm01AndSendAsync();

        // DMM01測定値（DCV）をキーボード入力としてEnterまで送信する
        private async Task ReadDmm01AndSendAsync() {
            if (MainWindow.IsProcessing) { return; }
            try {
                var output = await ReadDmm(_instDmm01);
                new InputSimulator().Keyboard
                    .TextEntry(output.ToString("0.000"))
                    .Sleep(100)
                    .KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM02測定値コピー
        private async Task ActionHotkeyBracketR() => await ReadDmm02AndSendAsync();
        private async Task ActionHotkeyNumAdd() => await ReadDmm02AndSendAsync();

        // DMM02測定値（DCV）をキーボード入力としてEnterまで送信する
        private async Task ReadDmm02AndSendAsync() {
            if (MainWindow.IsProcessing) { return; }
            try {
                var output = await ReadDmm(_instDmm02);
                new InputSimulator().Keyboard
                    .TextEntry(output.ToString("0.0000"))
                    .Sleep(100)
                    .KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSCローテーション
        private async Task ActionHotkeyPeriod() {
            if (MainWindow.IsProcessing) { return; }
            await RotationOsc(_instOsc, true);
        }
        // OSC meas1測定値コピー
        private async Task ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadOsc(_instOsc, 1);

                new InputSimulator().Keyboard
                    .TextEntry((output * 1000).ToString("0.000"))
                    .Sleep(100)
                    .KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSC meas2測定値コピー
        private async Task ActionHotkeyBackslash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadOsc(_instOsc, 2);

                new InputSimulator().Keyboard
                    .TextEntry((output * 1000).ToString("0.000"))
                    .Sleep(100)
                    .KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DCSローテーション
        private async Task ActionHotkeyAtsign() {
            if (MainWindow.IsProcessing) { return; }
            await RotationDcs(_instDcs, true);
        }
        // DCSローテーション逆方向
        private async Task ActionHotkeyShiftAtsign() {
            if (MainWindow.IsProcessing) { return; }
            await RotationDcs(_instDcs, false);
        }
        private async Task ActionHotkeyNumDivide() {
            if (MainWindow.IsProcessing) { return; }
            await RotationDcs(_instDcs, true);
        }

        // HotKeyの登録
        private void SetHotKey() {
            ThrowIfDisposed();

            MainWindow.HotkeysList.Clear();

            if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                    new(ModNone, HotkeyNumMultiply, ActionHotkeyNumMultiply),
                    ]);
            }
            if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                        new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
                    ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                        new(ModNone, HotkeySlash, ActionHotkeySlash),
                        new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    ]);
            }
            if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                        new(ModShift, HotkeyAtsign, ActionHotkeyShiftAtsign),
                        new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
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
        private void UserControl_Unloaded(object sender, RoutedEventArgs e) { Dispose(); }

    }
}
