using InspectionTools.Common;
using static InspectionTools.Common.InstrumentHelper;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;
using static InspectionTools.MainWindow;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.Product {
    /// <summary>
    /// EL5000DCUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL5000DCUserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        // MainWindowへの参照をセットする
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly CntInstClass _instCnt = new();
        private readonly DcsInstClass _instDcs01 = new();
        private readonly DcsInstClass _instDcs02 = new();
        private readonly DmmInstClass _instDmm = new();
        private readonly FgInstClass _instFg = new();
        private readonly InputSimulator _sim = new();

        private readonly Dictionary<InstClass, (SwitchCommand Init, List<SwitchCommand> Settings)> _dicCommands = [];
        readonly Stopwatch _stopwatch = new();
        readonly DispatcherTimer _timer = new();

        public EL5000DCUserControl() {
            InitializeComponent();

            _timer.Interval = new TimeSpan(0, 0, 0, 1);
            _timer.Tick += new EventHandler(TimerMethod);
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
                _timer.Stop();
                _timer.Tick -= TimerMethod;

                // マネージドリソースの解放
                try {
                    // ホットキーのクリア
                    ClearHotKey();

                    // 計測器の解放
                    InstrumentHelper.SafeDispose(_instCnt);
                    InstrumentHelper.SafeDispose(_instDcs01);
                    InstrumentHelper.SafeDispose(_instDcs02);
                    InstrumentHelper.SafeDispose(_instFg);
                    InstrumentHelper.SafeDispose(_instDmm);

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
        ~EL5000DCUserControl() {
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
            MainWindow.UpdateComboBox(CntComboBox, "ユニバーサルカウンタ", [3]);
            MainWindow.UpdateComboBox(FgComboBox, "ファンクションジェネレータ", [2]);
            MainWindow.UpdateComboBox(Dcs01ComboBox, "パワーサプライ", [2]);
            MainWindow.UpdateComboBox(Dcs02ComboBox, "電流電圧発生器", [2]);
            MainWindow.UpdateComboBox(DmmComboBox, "デジタルマルチメータ", [1, 2]);
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            MainGrid.IsEnabled = !isVisible;
            _mainWindow?.ShowSpinner(isVisible);
        }

        // 選択した機器のVisaAddressを取得
        private void SelectInst() {
            MainWindow.GetVisaAddress(_instCnt, CntComboBox);
            MainWindow.GetVisaAddress(_instFg, FgComboBox);
            MainWindow.GetVisaAddress(_instDcs01, Dcs01ComboBox);
            MainWindow.GetVisaAddress(_instDcs02, Dcs02ComboBox);
            MainWindow.GetVisaAddress(_instDmm, DmmComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {

            _dicCommands[_instCnt] =
                (
                    Init: new() { Gpib = "*RST;:FUNC PWID;:INPA:COUP DC;:INPA:LPF ON;:INPA:SLOP NEG;:FRUN ON;" },
                    Settings: [
                            new() { Gpib= ":FRUN ON",   Query = false },
                    ]
                );

            _dicCommands[_instFg] =
                (
                    Init: new() { Visa = "*RST;:FUNC SQU;:FREQ 7E3;:VOLT 3.0VPP;:VOLT:OFFS 1.5;*OPC?", Query = true },
                    Settings: [
                            new() { Visa = ":OUTP OFF;*OPC?" },
                            new() { Visa = ":OUTP ON;*OPC?" },
                    ]
                );

            _dicCommands[_instDcs01] =
                (
                    Init: new() { DcsMode = DcsMode.Off, Visa = "*RST;:VOLT 30;*OPC?", Query = true },
                    Settings: [
                        new() { DcsMode = DcsMode.On,   Visa = ":OUTPUT ON;*OPC?",  Query = true },
                        new() { DcsMode = DcsMode.Off,  Visa = ":OUTPUT OFF;*OPC?", Query = true },
                    ]
                );

            _dicCommands[_instDcs02] =
                (
                    Init: new() { DcsMode = DcsMode.Off, Visa = "SIR3,SOI+0,SBY", Gpib = "RCF5R6S0EO0E" },
                    Settings: [
                        new() { DcsMode = DcsMode.Off,  Text = "OFF",   Visa = "SIR2,SOI+0MA,SBY" },
                        new() { DcsMode = DcsMode.On,   Text = "4.0mA", Visa = "SIR2,SOI+4MA,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "20mA",  Visa = "SIR2,SOI+20MA,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "5V",    Visa = "SVR5,SOV+5V,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "1V",    Visa = "SVR5,SOV+1V,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "4.0mA", Visa = "SIR2,SOI+4MA,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "20mA",  Visa = "SIR2,SOI+20MA,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "5V",    Visa = "SVR5,SOV+5V,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "1V",    Visa = "SVR5,SOV+1V,OPR" },
                        new() { DcsMode = DcsMode.Off,  Text = "OFF",   Visa = "SIR2,SOI+0MA,SBY" },
                        new() { DcsMode = DcsMode.On,   Text = "4.0mA", Visa = "SIR2,SOI+4MA,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "12mA",  Visa = "SIR2,SOI+12MA,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "20mA",  Visa = "SIR2,SOI+20MA,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "1V",    Visa = "SVR5,SOV+1V,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "3V",    Visa = "SVR5,SOV+3V,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "5V",    Visa = "SVR5,SOV+5V,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "4.0mA", Visa = "SIR2,SOI+4MA,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "12mA",  Visa = "SIR2,SOI+12MA,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "20mA",  Visa = "SIR2,SOI+20MA,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "1V",    Visa = "SVR5,SOV+1V,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "3V",    Visa = "SVR5,SOV+3V,OPR" },
                        new() { DcsMode = DcsMode.On,   Text = "5V",    Visa = "SVR5,SOV+5V,OPR" },
                    ]
                );

            _dicCommands[_instDmm] =
                (
                    Init: new() { DmmMode = DmmMode.DCV, Adc = "*RST,F1,R0,*OPC?", Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG:AUTO ON;*OPC?", Query = true },
                    Settings: [
                            new() { DmmMode = DmmMode.DCV,   Adc= "*RST,F1,R0,*OPC?",    Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG:AUTO ON;*OPC?", Query = true },
                            new() { DmmMode = DmmMode.DCI,   Adc= "*RST,F5,R6,*OPC?",    Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?", Query = true },
                        ]
                );
        }
        // 機器初期設定
        private void FormatSet() {
            (_instCnt.InstCommand, _instCnt.Query) = ResolveCommand(_dicCommands[_instCnt].Init, _instCnt.SignalType);
            (_instFg.InstCommand, _instFg.Query) = ResolveCommand(_dicCommands[_instFg].Init, _instFg.SignalType);
            (_instDcs01.InstCommand, _instDcs01.Query) = ResolveCommand(_dicCommands[_instDcs01].Init, _instDcs01.SignalType);
            (_instDcs02.InstCommand, _instDcs02.Query) = ResolveCommand(_dicCommands[_instDcs02].Init, _instDcs02.SignalType);
            (_instDmm.InstCommand, _instDmm.Query) = ResolveCommand(_dicCommands[_instDmm].Init, _instDmm.SignalType);
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

                InstClass[] devices = [_instCnt, _instFg, _instDcs01, _instDcs02, _instDmm];

                await DeviceConnectionHelper.ConnectInParallelAsync(devices);

                if (!string.IsNullOrEmpty(_instDcs01.VisaAddress)) {
                    _instDcs01.CurrentMode = _dicCommands[_instDcs01].Init.DcsMode;
                }
                if (!string.IsNullOrEmpty(_instDcs02.VisaAddress)) {
                    DcsNumberLabel.Text = "00";
                    DcsRangeLabel.Text = "OFF";
                    _instDcs02.CurrentMode = _dicCommands[_instDcs02].Init.DcsMode;
                }
                if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                    _instDmm.CurrentMode = _dicCommands[_instDmm].Init.DmmMode;
                }

                CntComboBox.IsEnabled = false;
                FgComboBox.IsEnabled = false;
                Dcs01ComboBox.IsEnabled = false;
                Dcs02ComboBox.IsEnabled = false;
                DmmComboBox.IsEnabled = false;
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

            _instCnt.ResetProperties();
            _instFg.ResetProperties();
            _instDcs01.ResetProperties();
            _instDcs02.ResetProperties();
            _instDmm.ResetProperties();

            _mainWindow?.SetButtonEnabled(ProductListButtonName, true);
            CntComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            Dcs01ComboBox.IsEnabled = true;
            Dcs02ComboBox.IsEnabled = true;
            DmmComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyCheckBox.IsChecked = false;
            DcsNumberLabel.Text = string.Empty;
            DcsRangeLabel.Text = string.Empty;
            _timer.Stop();
            _stopwatch.Reset();
            DcsTimer.Text = "00:00";
        }

        // Cnt測定値取得
        private async Task<decimal> ReadCnt(CntInstClass cntInstClass) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var output = await InstrumentService.ReadCntAsync(cntInstClass);

                var settings = _dicCommands[cntInstClass].Settings;
                if (settings.Count > cntInstClass.SettingNumber) {
                    var sw = settings[cntInstClass.SettingNumber];
                    (cntInstClass.InstCommand, cntInstClass.Query) = ResolveCommand(sw, cntInstClass.SignalType);
                    await DeviceController.ConnectAsync(cntInstClass);
                }

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // FGローテーション
        private async Task RotationFg(FgInstClass fgInstClass, bool isNext) {
            if (_disposed) return;

            try {
                VisibleProgressImage(true);

                var settings = _dicCommands[fgInstClass].Settings;
                fgInstClass.SettingNumber = (fgInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[fgInstClass.SettingNumber];
                (fgInstClass.InstCommand, fgInstClass.Query) = ResolveCommand(sw, fgInstClass.SignalType);

                await DeviceController.ConnectAsync(fgInstClass);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            } finally {
                VisibleProgressImage(false);
            }
        }
        // DCS01ON-OFF
        private async Task SwitchDcs01Async(DcsInstClass dcsInstClass, DcsMode mode) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var settings = _dicCommands[dcsInstClass].Settings;
                var sw = settings.FirstOrDefault(s => s.DcsMode == mode) ?? throw new InvalidOperationException($"'{mode}' に対応する設定が見つかりません。");
                (dcsInstClass.InstCommand, dcsInstClass.Query) = ResolveCommand(sw, dcsInstClass.SignalType);
                dcsInstClass.CurrentMode = mode;

                await DeviceController.ConnectAsync(dcsInstClass);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            } finally {
                VisibleProgressImage(false);
            }
        }
        // DCS01電流値取得
        private async Task<decimal> ReadDcs(DcsInstClass dcsInstClass) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var output = await InstrumentService.ReadDcsAsync(dcsInstClass);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // DCS02ローテーション
        private async Task RotationDcs02(DcsInstClass dcsInstClass, bool isNext) {
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

                if (dcsInstClass.SettingNumber != 0) {
                    _stopwatch.Restart();
                    if (!_timer.IsEnabled) {
                        _timer.Start();
                    }
                }
                else {
                    _timer.Stop();
                    _stopwatch.Reset();
                    DcsTimer.Text = "00:00";
                }

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                (dmmInstClass.InstCommand, dmmInstClass.Query) = ResolveCommand(sw, dmmInstClass.SignalType);
                dmmInstClass.CurrentMode = sw.DmmMode;

                await DeviceController.ConnectAsync(dmmInstClass);

            } finally {
                VisibleProgressImage(false);
            }
        }

        // CNT測定値コピー
        private async Task ActionHotkeyComma() => await ReadCntAndSendAsync();
        private async Task ActionHotkeyNumDivide() => await ReadCntAndSendAsync();
        // CNT測定値をms単位に変換してキーボード入力としてEnterまで送信する
        private async Task ReadCntAndSendAsync() {
            if (MainWindow.IsProcessing) { return; }
            try {
                var output = await ReadCnt(_instCnt);

                _sim.Keyboard.TextEntry((output * 1000).ToString("0.000"));
                await Task.Delay(100);
                _sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // FGローテーション
        private async Task ActionHotkeyColon() {
            if (MainWindow.IsProcessing) { return; }
            await RotationFg(_instFg, true);
        }

        // DCS01電源切り替え
        private async Task ActionHotkeyBracketL() {
            if (MainWindow.IsProcessing) { return; }
            await SwitchDcs01Async(_instDcs01, DcsMode.On);
        }
        private async Task ActionHotkeyAtsign() {
            if (MainWindow.IsProcessing) { return; }
            await SwitchDcs01Async(_instDcs01, DcsMode.Off);
        }
        // DCS01電流値値コピー
        private async Task ActionHotkeyPeriod() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDcs(_instDcs01);

                _sim.Keyboard.TextEntry((output * 1000).ToString("0.000"));
                await Task.Delay(100);
                _sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // DCS02ローテーション
        private async Task ActionHotkeyBracketR() {
            if (MainWindow.IsProcessing) { return; }
            await RotationDcs02(_instDcs02, true);
        }
        private async Task ActionHotkeyShiftBracketR() {
            if (MainWindow.IsProcessing) { return; }
            await RotationDcs02(_instDcs02, false);
        }
        private async Task ActionHotkeyNumMultiply() {
            if (MainWindow.IsProcessing) { return; }
            await RotationDcs02(_instDcs02, true);
        }

        // DMM測定値コピー
        private async Task ActionHotkeySlash() => await ReadDmmAndSendAsync();
        private async Task ActionHotkeyNumAdd() => await ReadDmmAndSendAsync();
        // DMM測定値をモードに応じた単位に変換してキーボード入力としてEnterまで送信する
        private async Task ReadDmmAndSendAsync() {
            if (MainWindow.IsProcessing) { return; }
            try {
                var output = await ReadDmm(_instDmm);

                var outputValue = _instDmm.CurrentMode switch {
                    DmmMode.DCI => output * 1000,
                    DmmMode.DCV => output,
                    _ => output,
                };

                _sim.Keyboard.TextEntry(outputValue.ToString("0.0000"));
                await Task.Delay(100);
                _sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM切り替え
        private async Task ActionHotkeyBackSlash() {
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

            if (!string.IsNullOrEmpty(_instCnt.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyComma, ActionHotkeyComma),
                    new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
                ]);
            }
            if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDcs01.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketL, ActionHotkeyBracketL),
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                    new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDcs02.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                    new(ModShift, HotkeyBracketR, ActionHotkeyShiftBracketR),
                    new(ModNone, HotkeyNumMultiply, ActionHotkeyNumMultiply),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackSlash),
                    new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd)
                ]);
            }
            MainWindow.SetHotKey();
        }
        // 登録済みホットキーを解除する
        private static void ClearHotKey() {
            MainWindow.ClearHotKey();
        }

        // タイマーTickごとにストップウォッチの経過時間をDCSタイマーラベルに表示する
        private void TimerMethod(object? sender, EventArgs e) {
            var result = _stopwatch.Elapsed;
            DcsTimer.Text = $@"{result.Minutes:00}:{result.Seconds:00}";
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
