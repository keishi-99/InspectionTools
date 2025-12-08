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
    public partial class EL0122FIUserControl : UserControl, IMainWindowAware {

        private MainWindow? _mainWindow;
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DcsInstClass _instDcs = new();
        private readonly DmmInstClass _instDmm01 = new();
        private readonly DmmInstClass _instDmm02 = new();
        private readonly FgInstClass _instFg = new();
        private readonly OscInstClass _instOsc = new();

        public EL0122FIUserControl() {
            InitializeComponent();
        }

        // 起動時
        private void LoadEvents() {
            InstListImport();
            FormatSet();
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
        // 機器初期設定
        private void FormatSet() {
            _instDmm01.InstCommand = _instDmm01.SignalType switch {
                1 => "*RST,R7,*OPC?",
                2 => "*RST;:INIT:CONT 1;:VOLT:DC:RANG 200;*OPC?",
                _ => string.Empty
            };
            _instDmm02.InstCommand = _instDmm02.SignalType switch {
                1 => "*RST,F5,R7,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 0.2;*OPC?",
                _ => string.Empty,
            };
            _instFg.InstCommand = _instFg.SignalType switch {
                2 => "*RST;:FREQ 50;:VOLT 3.0VPP;:VOLT:OFFS 1.5;:SOUR:FUNC SQU;:OUTPUT ON;*OPC?",
                _ => string.Empty,
            };
            _instOsc.InstCommand = _instOsc.SignalType switch {
                2 =>
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
                _ => string.Empty,
            };
            _instDcs.InstCommand = _instDcs.SignalType switch {
                2 => "*RST;:VOLT 24;*OPC?",
                _ => string.Empty,
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {
            try {
                _mainWindow?.SetButtonEnabled("ProductListButton", false);

                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                CheckDmmId();
                FormatSet();

                InstClass[] devices = [_instDmm01, _instDmm02, _instFg, _instOsc, _instDcs];
                await Task.Run(() => {
                    foreach (var device in devices) {
                        MainWindow.ConnectDevice(device);
                    }
                });

                Dmm01ComboBox.IsEnabled = false;
                Dmm02ComboBox.IsEnabled = false;
                FgComboBox.IsEnabled = false;
                OscComboBox.IsEnabled = false;
                PsComboBox.IsEnabled = false;
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
            HotKeyChekBox.IsChecked = false;
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
        // 電源のON-OFF
        private async Task SwitchDcsAsync(DcsInstClass dcsInstClass, string cmd) {
            try {
                dcsInstClass.InstCommand = $":OUTPUT {cmd};*OPC?";
                MainWindow.ConnectDevice(dcsInstClass);

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
                    await SwitchDcsAsync(_instDcs, "ON");
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
                    await SwitchDcsAsync(_instDcs, "OFF");
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
            await SwitchDcsAsync(_instDcs, "ON");
        }
        private async void ActionHotkeyBracketL() {
            await SwitchDcsAsync(_instDcs, "OFF");
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
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }


    }
}
