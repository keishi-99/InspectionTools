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
    /// EL9220UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL9220UserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        // MainWindowへの参照をセットする
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DcsInstClass _instDcs = new();
        private readonly DmmInstClass _instDmm = new();
        private readonly FgInstClass _instFg = new();

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
                    InstrumentHelper.SafeDispose(_instDcs);
                    InstrumentHelper.SafeDispose(_instDmm);
                    InstrumentHelper.SafeDispose(_instFg);

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
        ~EL9220UserControl() {
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
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            MainGrid.IsEnabled = !isVisible;
            _mainWindow?.ShowSpinner(isVisible);
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
                    Init: new() { DcsMode = DcsMode.Off, Visa = "*RST;:VOLT 3.6;*OPC?", Query = true },
                    Settings: [
                        new() { DcsMode = DcsMode.Off,  Visa = $":OUTPUT OFF;*OPC?",    Query = true },
                        new() { DcsMode = DcsMode.On,   Visa = $":OUTPUT ON;*OPC?",     Query = true },
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

                InstClass[] devices = [_instDcs, _instDmm, _instFg];

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

            _mainWindow?.SetButtonEnabled(ProductListButtonName, true);
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
                dcsInstClass.CurrentMode = sw.DcsMode;

                await DeviceController.ConnectAsync(dcsInstClass);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            } finally {
                VisibleProgressImage(false);
            }
        }
        // FG切り替え
        private async Task RotationFg(FgInstClass fgInstClass, bool isNext) {
            if (_disposed) return;

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
        private async Task ActionHotkeyPeriod() {
            if (MainWindow.IsProcessing) { return; }
            try {
                await SwitchDcs(_instDcs, true);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private async Task ActionHotkeyNumDivide() {
            if (MainWindow.IsProcessing) { return; }
            try {
                await SwitchDcs(_instDcs, true);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM測定値コピー
        private async Task ActionHotkeySlash() => await ReadDmmAndSendAsync();
        private async Task ActionHotkeyNumSubtract() => await ReadDmmAndSendAsync();

        // DMM測定値をμA単位に変換してキーボード入力としてEnterまで送信する
        private async Task ReadDmmAndSendAsync() {
            if (MainWindow.IsProcessing) { return; }
            try {
                var output = await ReadDmm(_instDmm);

                new InputSimulator().Keyboard
                    .TextEntry((output * 1000000).ToString())  // μA単位に変換
                    .Sleep(100)
                    .KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // FGローテーション
        private async Task ActionHotkeyBackslash() {
            if (MainWindow.IsProcessing) { return; }
            await RotationFg(_instFg, true);
        }
        private async Task ActionHotkeyNumMultiply() {
            if (MainWindow.IsProcessing) { return; }
            await RotationFg(_instFg, true);
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
