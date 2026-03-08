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
    /// PAF5ampUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class PAF5ampUserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        // MainWindowへの参照をセットする
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DcsInstClass _instDcs = new();
        private readonly DmmInstClass _instDmm = new();
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

        public PAF5ampUserControl() {
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
                    InstrumentHelper.SafeDispose(_instDmm);
                    InstrumentHelper.SafeDispose(_instFg);
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
        ~PAF5ampUserControl() {
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
            MainWindow.UpdateComboBox(DcsComboBox, "パワーサプライ", [2]);
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
            MainWindow.GetVisaAddress(_instDcs, DcsComboBox);
            MainWindow.GetVisaAddress(_instDmm, DmmComboBox);
            MainWindow.GetVisaAddress(_instFg, FgComboBox);
            MainWindow.GetVisaAddress(_instOsc, OscComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicCommands[_instDmm] =
                (
                    Init: new() { DmmMode = DmmMode.DCI, Adc = "*RST,F5,R7,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?", Query = true },
                    Settings: []
                );
            _dicCommands[_instDcs] =
                (
                    Init: new() { DcsMode = DcsMode.Off, Visa = "*RST;:VOLT 1.5;*OPC?", Query = true },
                    Settings: [
                        new() { DcsMode = DcsMode.On, Visa = ":OUTPUT ON;*OPC?", Query = true },
                        new() { DcsMode = DcsMode.Off, Visa = ":OUTPUT OFF;*OPC?", Query = true },
                    ]
                );
            _dicCommands[_instFg] =
                (
                    Init: new() { Visa = "*RST;:FREQ 1.7E0;*OPC?", Query = true },
                    Settings: [
                        new() { Visa = ":FREQ 1.7E0;*OPC?", Query = true },
                        new() { Visa = ":FREQ 1.0E1;*OPC?", Query = true },
                        new() { Visa = ":FREQ 4.0E3;*OPC?", Query = true },
                    ]
                );

            _dicCommands[_instOsc] =
                (
                    Init: new() {
                        Visa =
                            """
                            *RST;
                            :HEADER 0;
                            :CH1:SCALE 1.0E-1;
                            :HORIZONTAL:MAIN:SCALE 5.0E-2;
                            :MEASUREMENT:MEAS1:TYPE PK2PK;SOURCE CH1;
                            *OPC?
                            """,
                        Query = true
                    },
                    Settings: [
                        new() { Visa = ":HORIZONTAL:MAIN:SCALE 5.0E-2;*OPC?", Query = true },
                        new() { Visa = ":HORIZONTAL:MAIN:SCALE 2.5E-2;*OPC?", Query = true },
                        new() { Visa = ":HORIZONTAL:MAIN:SCALE 5.0E-5;*OPC?", Query = true },
                    ]
                );
        }
        // 機器初期設定
        private void FormatSet() {
            (_instDmm.InstCommand, _instDmm.Query) = ResolveCommand(_dicCommands[_instDmm].Init, _instDmm.SignalType);
            (_instDcs.InstCommand, _instDcs.Query) = ResolveCommand(_dicCommands[_instDcs].Init, _instDcs.SignalType);
            (_instFg.InstCommand, _instFg.Query) = ResolveCommand(_dicCommands[_instFg].Init, _instFg.SignalType);
            (_instOsc.InstCommand, _instOsc.Query) = ResolveCommand(_dicCommands[_instOsc].Init, _instOsc.SignalType);
        }
        // 信号種別に応じたコマンド文字列とクエリフラグを返す
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
                _mainWindow?.SetButtonEnabled(ProductListButtonName, false);

                HotKeyCheckBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();

                RegDictionary();
                FormatSet();

                InstClass[] devices = [_instDcs, _instDmm, _instFg, _instOsc];

                await DeviceConnectionHelper.ConnectInParallelAsync(devices);

                if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                    _instDcs.CurrentMode = _dicCommands[_instDcs].Init.DcsMode;
                }
                if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                    _instDmm.CurrentMode = _dicCommands[_instDmm].Init.DmmMode;
                }

                DcsComboBox.IsEnabled = false;
                DmmComboBox.IsEnabled = false;
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

        // 解除
        private void Release() {
            VisibleProgressImage(false);

            _instDcs.ResetProperties();
            _instDmm.ResetProperties();
            _instFg.ResetProperties();
            _instOsc.ResetProperties();

            _mainWindow?.SetButtonEnabled(ProductListButtonName, true);
            DcsComboBox.IsEnabled = true;
            DmmComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyCheckBox.IsChecked = false;
        }

        // DCS切り替え
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
        // OSC切り替え
        private async void RotationOsc(OscInstClass oscInstClass, bool isNext) {
            ThrowIfDisposed();

            try {
                if (string.IsNullOrEmpty(oscInstClass.VisaAddress)) { return; }
                VisibleProgressImage(true);

                var settings = _dicCommands[oscInstClass].Settings;
                oscInstClass.SettingNumber = (oscInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[oscInstClass.SettingNumber];
                (oscInstClass.InstCommand, oscInstClass.Query) = ResolveCommand(sw, oscInstClass.SignalType);

                if (oscInstClass.InstCommand == string.Empty) { return; }

                await InstrumentService.RotateOscAsync(oscInstClass);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }

        // 電源ON
        private async void ActionHotkeyAtsign() {
            try {
                await SwitchDcsAsync(_instDcs, DcsMode.On);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // 電源OFF
        private async void ActionHotkeyBracketL() {
            try {
                await SwitchDcsAsync(_instDcs, DcsMode.Off);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM測定値コピー
        private async void ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000000).ToString("0"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // FGローテーション
        private void ActionHotkeyColon() {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, true);
        }
        // OSCローテーション
        private void ActionHotkeyBracketR() {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, true);
        }

        // HotKeyの登録
        private void SetHotKey() {
            ThrowIfDisposed();

            MainWindow.HotkeysList.Clear();

            if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                    new(ModNone, HotkeyBracketL, ActionHotkeyBracketL),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash)
                ]);
            }
            if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
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
