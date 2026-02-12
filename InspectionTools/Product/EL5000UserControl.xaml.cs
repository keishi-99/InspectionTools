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
    /// EL5000UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL5000UserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly CntInstClass _instCnt = new();
        private readonly DcsInstClass _instDcs = new();
        private readonly DmmInstClass _instDmm01 = new();
        private readonly DmmInstClass _instDmm02 = new();
        private readonly DmmInstClass _instDmm03 = new();

        private record SwitchCommand {
            public string Text { get; init; } = string.Empty;
            public string Adc { get; init; } = string.Empty;
            public string Visa { get; init; } = string.Empty;
            public string Gpib { get; init; } = string.Empty;
            public bool Query { get; init; } = false;
        }
        private readonly Dictionary<InstClass, (SwitchCommand Init, List<SwitchCommand> Settings)> _dicCommands = [];

        public EL5000UserControl() {
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
                    DisposeInstrument(_instCnt);
                    DisposeInstrument(_instDcs);
                    DisposeInstrument(_instDmm01);
                    DisposeInstrument(_instDmm02);
                    DisposeInstrument(_instDmm03);

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
        ~EL5000UserControl() {
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
            MainWindow.UpdateComboBox(CntComboBox, "ユニバーサルカウンタ", [3]);
            MainWindow.UpdateComboBox(DcsComboBox, "電流電圧発生器", [2, 3]);
            MainWindow.UpdateComboBox(Dmm01ComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(Dmm02ComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(Dmm03ComboBox, "デジタルマルチメータ", [1, 2]);
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            MainWindow.IsProcessing = isVisible;
            MainGrid.IsEnabled = !isVisible;
            ProgressRing.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }

        // 選択した機器のVisaAddressを取得
        private void SelectInst() {
            MainWindow.GetVisaAddress(_instCnt, CntComboBox);
            MainWindow.GetVisaAddress(_instDcs, DcsComboBox);
            MainWindow.GetVisaAddress(_instDmm01, Dmm01ComboBox);
            MainWindow.GetVisaAddress(_instDmm02, Dmm02ComboBox);
            MainWindow.GetVisaAddress(_instDmm03, Dmm03ComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {

            _dicCommands[_instCnt] =
                (
                    Init: new() { Gpib = ":FUNC FINA;:FRUN ON;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instDmm01] =
                (
                    Init: new() { Adc = "*RST,F5,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instDmm02] =
                (
                    Init: new() { Adc = "*RST,F1,R5,*OPC?", Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG 2;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instDmm03] =
                (
                    Init: new() { Adc = "*RST,F5,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?", Query = true },
                    Settings: [
                            new() { Text = "DCI",   Adc= "*RST,F5,R6,*OPC?",    Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?" },
                            new() { Text = "DCV",   Adc= "*RST,F1,R5,*OPC?",    Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG 2;*OPC?" },
                        ]
                );

            _dicCommands[_instDcs] =
                (
                    Init: new() { Visa = "SIR3,SOI+0,SBY", Gpib = "RCF5R6S0EO0E" },
                    Settings: [
                        new() { Text = "OFF",   Visa = "SOI+0MA,SBY",   Gpib = "F5R6S0EO0E" },
                        new() { Text = "4.0mA", Visa = "SOI+4MA,OPR",   Gpib = "F5R6S4.0E-3O1E" },
                        new() { Text = "20mA",  Visa = "SOI+20MA,OPR",  Gpib = "F5R6S20.0E-3O1E" },
                        new() { Text = "4.0mA", Visa = "SOI+4MA,OPR",   Gpib = "F5R6S4.0E-3O1E" },
                        new() { Text = "20mA",  Visa = "SOI+20MA,OPR",  Gpib = "F5R6S20.0E-3O1E" },
                        new() { Text = "OFF",   Visa = "SOI+0MA,SBY",   Gpib = "F5R6S0EO0E" },
                        new() { Text = "20mA",  Visa = "SOI+20MA,OPR",  Gpib = "F5R6S20.0E-3O1E" },
                        new() { Text = "12mA",  Visa = "SOI+12MA,OPR",  Gpib = "F5R6S12.0E-3O1E" },
                        new() { Text = "4.0mA", Visa = "SOI+4MA,OPR",   Gpib = "F5R6S4.0E-3O1E" },
                        new() { Text = "20mA",  Visa = "SOI+20MA,OPR",  Gpib = "F5R6S20.0E-3O1E" },
                        new() { Text = "12mA",  Visa = "SOI+12MA,OPR",  Gpib = "F5R6S12.0E-3O1E" },
                        new() { Text = "4.0mA", Visa = "SOI+4MA,OPR",   Gpib = "F5R6S4.0E-3O1E" },
                    ]
                );
        }
        // 機器初期設定
        private void FormatSet() {
            (_instCnt.InstCommand, _instCnt.Query) = ResolveCommand(_dicCommands[_instCnt].Init, _instCnt.SignalType);
            (_instDcs.InstCommand, _instDcs.Query) = ResolveCommand(_dicCommands[_instDcs].Init, _instDcs.SignalType);
            (_instDmm01.InstCommand, _instDmm01.Query) = ResolveCommand(_dicCommands[_instDmm01].Init, _instDmm01.SignalType);
            (_instDmm02.InstCommand, _instDmm02.Query) = ResolveCommand(_dicCommands[_instDmm02].Init, _instDmm02.SignalType);
            (_instDmm03.InstCommand, _instDmm03.Query) = ResolveCommand(_dicCommands[_instDmm03].Init, _instDmm03.SignalType);
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

                InstClass[] devices = [_instCnt, _instDcs, _instDmm01, _instDmm02, _instDmm03];
                RegDictionary();
                FormatSet();
                var tasks = devices.Select(device => MainWindow.ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);

                if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                    DcsNumberLabel.Text = "00";
                    DcsRangeLabel.Text = "OFF";
                }

                CntComboBox.IsEnabled = false;
                DcsComboBox.IsEnabled = false;
                Dmm01ComboBox.IsEnabled = false;
                Dmm02ComboBox.IsEnabled = false;
                Dmm03ComboBox.IsEnabled = false;
                ConnectButton.IsEnabled = false;
                ReleaseButton.IsEnabled = true;

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

            _instCnt.ResetProperties();
            _instDcs.ResetProperties();
            _instDmm01.ResetProperties();
            _instDmm02.ResetProperties();
            _instDmm03.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
            CntComboBox.IsEnabled = true;
            DcsComboBox.IsEnabled = true;
            Dmm01ComboBox.IsEnabled = true;
            Dmm02ComboBox.IsEnabled = true;
            Dmm03ComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyCheckBox.IsChecked = false;
            DcsNumberLabel.Text = string.Empty;
            DcsRangeLabel.Text = string.Empty;
        }

        // Cnt測定値取得
        private async Task<decimal> ReadCnt(CntInstClass cntInstClass) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var output = await MainWindow.ReadCnt(cntInstClass);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // DCSローテーション
        private async void RotationDcs(DcsInstClass dcsInstClass, bool isNext) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var settings = _dicCommands[dcsInstClass].Settings;
                dcsInstClass.SettingNumber = (dcsInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[dcsInstClass.SettingNumber];
                dcsInstClass.InstCommand = dcsInstClass.SignalType switch {
                    1 => sw.Adc,
                    2 => sw.Visa,
                    3 => sw.Gpib,
                    _ => string.Empty,
                };
                dcsInstClass.Query = sw.Query;

                await MainWindow.ConnectDeviceAsync(dcsInstClass);
                DcsNumberLabel.Text = dcsInstClass.SettingNumber.ToString("00");
                DcsRangeLabel.Text = sw.Text;

                VisibleProgressImage(false);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
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

                await MainWindow.ConnectDeviceAsync(dmmInstClass);

            } finally {
                VisibleProgressImage(false);
            }
        }

        // CNT測定値コピー
        private async void ActionHotkeyComma() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadCnt(_instCnt);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(output.ToString());
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
                var output = await ReadCnt(_instCnt);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(output.ToString());
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DCSローテーション
        private void ActionHotkeyBracketL() {
            if (MainWindow.IsProcessing) { return; }
            RotationDcs(_instDcs, true);
        }
        private void ActionHotkeyShiftBracketL() {
            if (MainWindow.IsProcessing) { return; }
            RotationDcs(_instDcs, false);
        }
        private void ActionHotkeyNumDivide() {
            if (MainWindow.IsProcessing) { return; }
            RotationDcs(_instDcs, true);
        }
        // DMM01測定値コピー
        private async void ActionHotkeyPeriod() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm01);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(output.ToString("0.000"));
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
                var output = await ReadDmm(_instDmm01);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(output.ToString("0.000"));
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
                sim.Keyboard.TextEntry(output.ToString("0.0000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private async void ActionHotkeyNumAdd() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm02);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(output.ToString("0.0000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM03測定値コピー
        private async void ActionHotkeyBackSlash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm03);

                var commandType = _dicCommands[_instDmm03].Settings[_instDmm03.SettingNumber].Text;
                var outputValue = commandType switch {
                    "DCI" => output * 1000m,
                    "DCV" => output,
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
        // DMM03切り替え
        private async void ActionHotkeyBracketR() {
            if (MainWindow.IsProcessing) { return; }

            try {

                await SwitchDmm(_instDmm03, true);

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
                    new(ModNone, HotkeyNumMultiply, ActionHotkeyNumMultiply),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketL, ActionHotkeyBracketL),
                    new(ModShift, HotkeyBracketL, ActionHotkeyShiftBracketL),
                    new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                    new(ModNone, HotkeyNumSubtract, ActionHotkeyNumSubtract),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm03.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackSlash),
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
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
