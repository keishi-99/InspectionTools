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
    /// EL4001UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL4001UserControl : UserControl, IMainWindowAware {

        private MainWindow? _mainWindow;
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DcsInstClass _instDcs = new();
        private readonly DmmInstClass _instDmm01 = new();
        private readonly DmmInstClass _instDmm02 = new();
        private readonly OscInstClass _instOsc = new();

        public EL4001UserControl() {
            InitializeComponent();
        }

        private Dictionary<int, (string cmd2, string cmd3, string text)> _dicSwitchDcs = [];

        // 起動時
        private void LoadEvents() {
            InstListImport();
            FormatSet();
            RegDictionary();
            var parentWindow = Window.GetWindow(this);
            MainWindow.AdjustWindowSizeToUserControl(parentWindow);
        }
        private void InstListImport() {
            // デジタルマルチメータ、ファンクションジェネレータ、オシロスコープのコンボボックスを更新する
            MainWindow.UpdateComboBox(DcsComboBox, "電流電圧発生器", [2, 3]);
            MainWindow.UpdateComboBox(Dmm01ComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(Dmm02ComboBox, "デジタルマルチメータ", [1, 2]);
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
            MainWindow.GetVisaAddress(_instDmm01, Dmm01ComboBox);
            MainWindow.GetVisaAddress(_instDmm02, Dmm02ComboBox);
            MainWindow.GetVisaAddress(_instOsc, OscComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicSwitchDcs = new Dictionary<int, (string cmd2, string cmd3, string text)>
            {
                { 0, ("SOI+0MA,SBY", "F5R6S0EO0E", "OFF") },
                { 1, ("SOI+4MA,OPR", "F5R6S4.0E-3O1E", "4.0mA") },
                { 2, ("SOI+20MA,OPR", "F5R6S20.0E-3O1E", "20mA") },
                { 3, ("SOI+4MA,OPR", "F5R6S4.0E-3O1E", "4.0mA") },
                { 4, ("SOI+20MA,OPR", "F5R6S20.0E-3O1E", "20mA") },
                { 5, ("SOI+0MA,SBY", "F5R6S0EO0E", "OFF") },
                { 6, ("SOI+22MA,OPR", "F5R6S22.0E-3O1E", "22mA") },
                { 7, ("SOI+20MA,OPR", "F5R6S20.0E-3O1E", "20mA") },
                { 8, ("SOI+12MA,OPR", "F5R6S12.0E-3O1E", "12mA") },
                { 9, ("SOI+4MA,OPR", "F5R6S4.0E-3O1E", "4.0mA") },
                { 10, ("SOI+3.2MA,OPR", "F5R6S3.2E-3O1E", "3.2mA") },
                { 11, ("SOI+22MA,OPR", "F5R6S22.0E-3O1E", "22mA") },
                { 12, ("SOI+20MA,OPR", "F5R6S20.0E-3O1E", "20mA") },
                { 13, ("SOI+12MA,OPR", "F5R6S12.0E-3O1E", "12mA") },
                { 14, ("SOI+4MA,OPR", "F5R6S4.0E-3O1E", "4.0mA") },
                { 15, ("SOI+3.2MA,OPR", "F5R6S3.2E-3O1E", "3.2mA") },
            };
        }
        // 機器初期設定
        private void FormatSet() {
            _instDcs.InstCommand = _instDcs.SignalType switch {
                2 => "SIR3,SOI+0,SBY,*OPC?",
                3 => "RCF5R6S0EO0E",
                _ => string.Empty,
            };
            _instDmm01.InstCommand = _instDmm01.SignalType switch {
                1 => "*RST,F1,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:VOLT:DC:RANG 20;*OPC?",
                _ => string.Empty,
            };
            _instDmm02.InstCommand = _instDmm02.SignalType switch {
                1 => "*RST,F1,R5,*OPC?",
                2 => "*RST;:INIT:CONT 1;:VOLT:DC:RANG 2;*OPC?",
                _ => string.Empty,
            };
            _instOsc.InstCommand = _instOsc.SignalType switch {
                2 => """
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
                _ => string.Empty,
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {
            try {
                _mainWindow?.SetButtonEnabled("ProductListButton", false);

                HotKeyCheckBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                CheckDmmId();
                FormatSet();

                InstClass[] devices = [_instDcs, _instDmm01, _instDmm02, _instOsc];
                var tasks = devices.Select(device => MainWindow.ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);

                if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                    DcsNumberLabel.Text = "00";
                    DcsRangeLabel.Text = "OFF";
                }
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) { OscRangeLabel.Text = "50m"; }

                Dmm01ComboBox.IsEnabled = false;
                Dmm02ComboBox.IsEnabled = false;
                OscComboBox.IsEnabled = false;
                DcsComboBox.IsEnabled = false;
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
        private void CheckDmmId() {
            var indices = new[] { _instDmm01.Index, _instDmm02.Index }
                .Where(i => i >= 1); // 未選択(0以下)は無視

            if (indices.Count() != indices.Distinct().Count()) {
                throw new Exception("同じ測定器が選択されています。");
            }
        }

        // 解除
        private void Release() {
            VisibleProgressImage(false);

            _instDcs.ResetProperties();
            _instDmm01.ResetProperties();
            _instDmm02.ResetProperties();
            _instOsc.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
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
            try {
                VisibleProgressImage(true);

                var output = await MainWindow.ReadOsc(oscInstClass, meas);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // DCSローテーション
        private async void RotationDcs(DcsInstClass dcsInstClass, bool isNext) {
            try {
                VisibleProgressImage(true);

                var maxSettingNumber = _dicSwitchDcs.Keys.Count;
                dcsInstClass.SettingNumber = (dcsInstClass.SettingNumber + (isNext ? 1 : -1) + maxSettingNumber) % maxSettingNumber;
                var settingNumber = dcsInstClass.SettingNumber;

                dcsInstClass.InstCommand = dcsInstClass.SignalType switch {
                    2 => _dicSwitchDcs[settingNumber].cmd2,
                    3 => _dicSwitchDcs[settingNumber].cmd3,
                    _ => throw new ApplicationException(),
                };

                await MainWindow.ConnectDeviceAsync(dcsInstClass);
                DcsNumberLabel.Text = settingNumber.ToString("00");
                DcsRangeLabel.Text = _dicSwitchDcs[settingNumber].text;

                VisibleProgressImage(false);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSCローテーション
        private async void RotationOsc(OscInstClass oscInstClass) {
            try {
                VisibleProgressImage(true);

                (oscInstClass.InstCommand, var rangeText) = oscInstClass.SettingNumber switch {
                    0 => (":HORIZONTAL:MAIN:SCALE 5.0E-4;POSITION 0.0;*OPC?", "500u"),
                    1 => (":HORIZONTAL:MAIN:SCALE 5.0E-2;POSITION 1.0E-1;*OPC?", "50ms"),
                    _ => throw new ApplicationException(),
                };

                await MainWindow.ConnectDeviceAsync(oscInstClass);

                OscRangeLabel.Text = rangeText;

                // 設定番号を反転
                oscInstClass.SettingNumber = (oscInstClass.SettingNumber == 0) ? 1 : 0;

                VisibleProgressImage(false);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // DMM01測定値コピー
        private async void ActionHotkeyColon() {
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
        private async void ActionHotkeyNumMultiply() {
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
        private async void ActionHotkeyBracketR() {
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
        // OSCローテーション
        private void ActionHotkeyPeriod() {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc);
        }
        // OSC meas1測定値コピー
        private async void ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadOsc(_instOsc, 1);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSC meas2測定値コピー
        private async void ActionHotkeyBackslash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadOsc(_instOsc, 2);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DCSローテーション
        private void ActionHotkeyAtsign() {
            if (MainWindow.IsProcessing) { return; }
            RotationDcs(_instDcs, true);
        }
        private void ActionHotkeyShiftAtsign() {
            if (MainWindow.IsProcessing) { return; }
            RotationDcs(_instDcs, false);
        }
        private void ActionHotkeyNumDivide() {
            if (MainWindow.IsProcessing) { return; }
            RotationDcs(_instDcs, true);
        }

        // HotKeyの登録
        private void SetHotKey() {
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
        private static void ClearHotKey() {
            MainWindow.ClearHotKey();
        }

        // イベントハンドラ
        private void UserControl_Loaded(object sender, RoutedEventArgs e) { LoadEvents(); }
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void HotKeyCheckBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyCheckBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }


    }
}
