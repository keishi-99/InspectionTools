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
    /// EL0122FIUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL0122FIUserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DcsInstClass _instDcs = new();
        private readonly DmmInstClass _instDmm01 = new();
        private readonly DmmInstClass _instDmm02 = new();
        private readonly FgInstClass _instFg = new();
        private readonly OscInstClass _instOsc = new();

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

        public EL0122FIUserControl() {
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
                    DisposeInstrument(_instDmm01);
                    DisposeInstrument(_instDmm02);
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
        ~EL0122FIUserControl() {
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
            MainWindow.UpdateComboBox(FgComboBox, "ファンクションジェネレータ", [2]);
            MainWindow.UpdateComboBox(OscComboBox, "オシロスコープ", [2]);
            MainWindow.UpdateComboBox(PsComboBox, "パワーサプライ", [2]);
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
            MainWindow.GetVisaAddress(_instFg, FgComboBox);
            MainWindow.GetVisaAddress(_instOsc, OscComboBox);
            MainWindow.GetVisaAddress(_instDcs, PsComboBox);
        }

        // 機器設定辞書登録
        private void RegDictionary() {

            _dicCommands[_instDcs] =
                (
                    Init: new() { DcsMode = DcsMode.Off, Visa = "*RST;:VOLT 24;*OPC?", Query = true },
                    Settings: [
                        new() { DcsMode = DcsMode.On,   Visa = ":OUTPUT ON;*OPC?",  Query = true },
                        new() { DcsMode = DcsMode.Off,  Visa = ":OUTPUT OFF;*OPC?", Query = true },
                    ]
                );

            _dicCommands[_instDmm01] =
                (
                    Init: new() { DmmMode = DmmMode.DCV, Adc = "*RST,R7,*OPC?", Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG 200;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instDmm02] =
                (
                    Init: new() { DmmMode = DmmMode.DCI, Adc = "*RST,F5,R7,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 0.2;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instFg] =
                (
                    Init: new() { Visa = "*RST;:FREQ 50;:VOLT 3.0VPP;:VOLT:OFFS 1.5;:SOUR:FUNC SQU;:OUTPUT ON;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instOsc] =
                (
                    Init: new() {
                        Visa =
                            """
                            *RST;:HEADER 0;
                            :SELECT:CH1 1;CH2 1;
                            :CH1:SCALE 1.0E1;POSITION 0.0E0;
                            :CH2:SCALE 1.0E1;POSITION -3.0E0;
                            :HORIZONTAL:MAIN:SCALE 5.0E-3;
                            :TRIGGER:MAIN:LEVEL 4.0E0;
                            :MEASUREMENT:MEAS1:TYPE PWIDTH;SOURCE CH1;
                            :MEASUREMENT:MEAS2:TYPE PERIod;SOURCE CH1;
                            *OPC?
                            """,
                        Query = true
                    },
                    Settings: []
                );
        }
        // 機器初期設定
        private void FormatSet() {
            (_instDcs.InstCommand, _instDcs.Query) = ResolveCommand(_dicCommands[_instDcs].Init, _instDcs.SignalType);
            (_instDmm01.InstCommand, _instDmm01.Query) = ResolveCommand(_dicCommands[_instDmm01].Init, _instDmm01.SignalType);
            (_instDmm02.InstCommand, _instDmm02.Query) = ResolveCommand(_dicCommands[_instDmm02].Init, _instDmm02.SignalType);
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
        private async Task ConnectInstAsync() {
            ThrowIfDisposed();

            try {
                _mainWindow?.SetButtonEnabled("ProductListButton", false);

                HotKeyCheckBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                ValidateDmmSelection();

                RegDictionary();
                FormatSet();

                InstClass[] devices = [_instDmm01, _instDmm02, _instFg, _instOsc, _instDcs];

                await Task.Run(() =>
                    DeviceConnectionHelper.ConnectInParallelAsync(devices)
                );

                if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                    _instDcs.CurrentMode = _dicCommands[_instDcs].Init.DcsMode;
                }
                if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                    _instDmm01.CurrentMode = _dicCommands[_instDmm01].Init.DmmMode;
                }
                if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                    _instDmm02.CurrentMode = _dicCommands[_instDmm02].Init.DmmMode;
                }

                Dmm01ComboBox.IsEnabled = false;
                Dmm02ComboBox.IsEnabled = false;
                FgComboBox.IsEnabled = false;
                OscComboBox.IsEnabled = false;
                PsComboBox.IsEnabled = false;
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
            var indices = new[] { _instDmm01.Index, _instDmm02.Index }
                .Where(i => i >= 1); // 未選択(0以下)は無視

            if (indices.Count() != indices.Distinct().Count()) {
                throw new InvalidOperationException("同じ測定器が選択されています。");
            }
        }

        // 解除
        private void Release() {
            VisibleProgressImage(false);

            _instDmm01.ResetProperties();
            _instDmm02.ResetProperties();
            _instFg.ResetProperties();
            _instOsc.ResetProperties();
            _instDcs.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
            Dmm01ComboBox.IsEnabled = true;
            Dmm02ComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            PsComboBox.IsEnabled = true;
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
        // 電源のON-OFF
        private async Task SwitchDcsAsync(DcsInstClass dcsInstClass, DcsMode mode) {
            ThrowIfDisposed();

            try {
                var settings = _dicCommands[dcsInstClass].Settings;
                var sw = settings.FirstOrDefault(s => s.DcsMode == mode) ?? throw new InvalidOperationException($"'{mode}' に対応する設定が見つかりません。");
                (dcsInstClass.InstCommand, dcsInstClass.Query) = ResolveCommand(sw, dcsInstClass.SignalType);
                dcsInstClass.CurrentMode = mode;

                await DeviceController.ConnectAsync(dcsInstClass);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // 全てのデータを処理するメソッド
        private async Task ProcessAllDataAsync(int delay) {
            if (MainWindow.IsProcessing) { return; }
            try {
                var sim = new InputSimulator();
                VisibleProgressImage(true);

                if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                    await SwitchDcsAsync(_instDcs, DcsMode.On);
                    await Task.Delay(delay);
                }

                // DMM01の値を取得
                if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                    var output = await ReadDmm(_instDmm01);

                    sim.Keyboard.TextEntry(output.ToString("0.00"));
                    await Task.Delay(100);
                    sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
                    await Task.Delay(300);
                }
                // DMM02の値を取得
                if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                    var output = await ReadDmm(_instDmm02);

                    sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
                    await Task.Delay(100);
                    sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
                    await Task.Delay(300);
                }
                // オシロスコープの値を取得
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                    var output = await ReadOsc(_instOsc, 1);

                    sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
                    await Task.Delay(100);
                    sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
                    await Task.Delay(300);

                    output = await ReadOsc(_instOsc, 2);
                    sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
                    await Task.Delay(100);
                    sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                    await Task.Delay(300);
                }

                if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                    await SwitchDcsAsync(_instDcs, DcsMode.Off);
                }

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }

        // DMM01測定値コピー
        private async void ActionHotkeyColon() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm01);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(output.ToString("0.00"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM02測定値コピー
        private async void ActionHotkeyBracketR() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm02);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSC1測定値コピー
        private async void ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadOsc(_instOsc, 1);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSC2測定値コピー
        private async void ActionHotkeyBackslash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadOsc(_instOsc, 2);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // 電源ON-OFF
        private async void ActionHotkeyAtsign() {
            await SwitchDcsAsync(_instDcs, DcsMode.On);
        }
        private async void ActionHotkeyBracketL() {
            await SwitchDcsAsync(_instDcs, DcsMode.Off);
        }
        // 一連の処理
        private async void ActionHotkeyComma() {
            if (int.TryParse(WaitTimeTextBox.Text, out var delay)) {
                await ProcessAllDataAsync(delay);
            }
            else {
                MessageBox.Show("数値に変換できません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private async void ActionHotkeyNumAdd() {
            if (int.TryParse(WaitTimeTextBox.Text, out var delay)) {
                await ProcessAllDataAsync(delay);
            }
            else {
                MessageBox.Show("数値に変換できません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // HotKeyの登録
        private void SetHotKey() {
            ThrowIfDisposed();

            MainWindow.HotkeysList.Clear();

            MainWindow.HotkeysList.AddRange([
                new(ModNone, HotkeyComma, ActionHotkeyComma),
                new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
            ]);

            if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                    new(ModNone, HotkeyBracketL, ActionHotkeyBracketL),
                ]);
            }

            MainWindow.SetHotKey();
        }
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
