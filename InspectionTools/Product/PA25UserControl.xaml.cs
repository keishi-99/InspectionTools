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
    public partial class PA25UserControl : UserControl, IMainWindowAware {

        private MainWindow? _mainWindow;
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DmmInstClass _instDmm = new();
        private readonly FgInstClass _instFg = new();
        private readonly OscInstClass _instOsc = new();

        public PA25UserControl() {
            InitializeComponent();
        }

        private Dictionary<int, (string cmd, string text)> _dicSwitchFg = [];
        private Dictionary<int, (string cmd, string text)> _dicSwitchOsc = [];
        private Dictionary<int, (string cmd, string text)> _dicSwitchROsc = [];

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
        // 機器初期設定
        private void FormatSet() {
            _instDmm.InstCommand = _instDmm.SignalType switch {
                1 => "*RST,F1,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:VOLT:DC:RANG 2;*OPC?",
                _ => string.Empty,
            };
            _instFg.InstCommand = _instFg.SignalType switch {
                2 => "*RST;:FREQ 1;:VOLT 0.44VPP;*OPC?",
                _ => string.Empty,
            };
            _instOsc.InstCommand = _instOsc.SignalType switch {
                2 =>
                    """
                    *RST;:HEADER 0;
                    :CH1:SCALE 1.0E-1;
                    :TRIGGER:MAIN:LEVEL 3.0E-1;
                    :CURSOR:FUNCTION VBArs;SELECT:SOURCE CH1;:CURSOR:VBArs:POSITION1 -1.34E-3;POSITION2 1.16E-3;
                    :MEASUREMENT:MEAS1:TYPE PK2PK;SOURCE CH1;
                    :MEASUREMENT:MEAS1:TYPE NONE;SOURCE CH1;
                    *OPC?
                    """,
                _ => string.Empty,
            };
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicSwitchFg = new Dictionary<int, (string cmd, string text)> {
                { 0, (":FREQ 1;:OUTPUT OFF;*OPC?", "OFF") },
                { 1, (":OUTPUT OFF;:FREQ 27;:OUTPUT ON;*OPC?", "27") },
                { 2, (":OUTPUT OFF;:FREQ 29;:OUTPUT ON;*OPC?", "29") },
                { 3, (":OUTPUT OFF;:FREQ 400;:OUTPUT ON;*OPC?", "400") },
                { 4, (":OUTPUT OFF;:FREQ 1;:OUTPUT ON;*OPC?", "1") },
                { 5, (":OUTPUT OFF;:FREQ 10;:OUTPUT ON;*OPC?", "10") },
                { 6, (":OUTPUT OFF;:FREQ 200;:OUTPUT ON;*OPC?", "200") },
                { 7, (":OUTPUT OFF;:FREQ 1;*OPC?", "OFF") } ,
                { 8, (":OUTPUT OFF;:FREQ 2200;:OUTPUT ON;*OPC?", "2200") } ,
                { 9, (":OUTPUT OFF;:FREQ 1000;:OUTPUT ON;*OPC?", "1000") } ,
                { 10, (":OUTPUT OFF;:FREQ 1;*OPC?", "OFF") } ,
                { 11, (":OUTPUT OFF;:FREQ 1000;:OUTPUT ON;*OPC?", "1000") } ,
                { 12, (":OUTPUT OFF;:FREQ 1;*OPC?", "OFF") } ,
                { 13, (":OUTPUT OFF;:FREQ 6250;:OUTPUT ON;*OPC?", "6250") } ,
                { 14, (":OUTPUT OFF;:FREQ 3800;:OUTPUT ON;*OPC?", "3800") },
                { 15, (":OUTPUT OFF;:FREQ 1000;:OUTPUT ON;*OPC?", "1000") },
            };

            _dicSwitchOsc = new Dictionary<int, (string cmd, string text)> {
                { 0,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 5.0E-4;
                        :CH1:SCALE 1.0E-1;COUPLING DC;
                        :TRIGGER:MAIN:LEVEL 3.0E-1;
                        *OPC?
                        """,
                        "1")  },
                { 1,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 2.5E-1;
                        *OPC?
                        """,
                        "2")  },
                { 2,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 2.5E-2;
                        *OPC?
                        """,
                        "3")  },
                { 3,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 1.0E-3;
                        *OPC?
                        """,
                        "4")  },
                { 4,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 2.5E-4;
                        :CH1:SCALE 2.0E0;
                        :TRIGGER:MAIN:LEVEL 3.84E0;
                        :MEASUREMENT:MEAS1:TYPE PERIOD;SOURCE CH1;
                        *OPC?
                        """,
                        "5")  },
                { 5,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 5.0E-5;
                        :CH1:SCALE 1.0E1;
                        :MEASUREMENT:MEAS1:TYPE NWIDTH;SOURCE CH1;
                        *OPC?
                        """,
                        "6")  },
                { 6,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 5.0E-2;
                        :CH1:SCALE 2.0E0;
                        :MEASUREMENT:MEAS1:TYPE PWIDTH;SOURCE CH1;
                        :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH1;
                        *OPC?
                        """,
                        "7")  },
                { 7,  (
                        """                        
                        :MEASUREMENT:MEAS1:TYPE MINIMUM;SOURCE CH1;
                        :MEASUREMENT:MEAS2:TYPE NONE;SOURCE CH1;
                        *OPC?
                        """,
                        "8")  },
                { 8,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 1.0E-2;                        
                        :CH1:SCALE 2.0E-2;COUPLING AC;
                        :TRIGGER:MAIN:LEVEL 1.69E-1;
                        :MEASUREMENT:MEAS1:TYPE NONE;SOURCE CH1;
                        *OPC?
                        """,
                        "9")  },
                { 9,  (
                        """
                        :CH1:COUPLING DC;
                        *OPC?
                        """,
                        "10")  },
                { 10,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 2.50E-4;  
                        :CH1:SCALE 1.0E0;
                        :TRIGGER:MAIN:LEVEL 3.84E0;
                        *OPC?
                        """,
                        "11")  },
                { 11,  (
                        """
                        :CH1:SCALE 1.0E-1;COUPLING AC;
                        :TRIGGER:MAIN:LEVEL 8.44E-1;
                        *OPC?
                        """,
                        "12")  },
            };

            _dicSwitchROsc = new Dictionary<int, (string cmd, string text)> {
                { 0,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 5.0E-4;
                        *OPC?
                        """,
                        "1")  },
                { 1,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 2.5E-1;
                        *OPC?
                        """,
                        "2")  },
                { 2,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 2.5E-2;
                        *OPC?
                        """,
                        "3")  },
                { 3,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 1.0E-3;
                        :CH1:SCALE 1.0E-1;COUPLING DC;
                        :TRIGGER:MAIN:LEVEL 3.0E-1;
                        :MEASUREMENT:MEAS1:TYPE NONE;SOURCE CH1;
                        
                        *OPC?
                        """,
                        "4")  },
                { 4,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 2.5E-4;
                        :CH1:SCALE 2.0E0;
                        :MEASUREMENT:MEAS1:TYPE PERIOD;SOURCE CH1;
                        *OPC?
                        """,
                        "5")  },
                { 5,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 5.0E-5;
                        :CH1:SCALE 1.0E1;
                        :MEASUREMENT:MEAS1:TYPE NWIDTH;SOURCE CH1;
                        :MEASUREMENT:MEAS2:TYPE NONE;SOURCE CH1;
                        *OPC?
                        """,
                        "6")  },
                { 6,  (
                        """
                        :MEASUREMENT:MEAS1:TYPE PWIDTH;SOURCE CH1;
                        :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH1;
                        *OPC?
                        """,
                        "7")  },
                { 7,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 5.0E-2;
                        :CH1:SCALE 2.0E0;
                        :TRIGGER:MAIN:LEVEL 3.84E0;
                        :MEASUREMENT:MEAS1:TYPE MINIMUM;SOURCE CH1;
                        *OPC?
                        """,
                        "8")  },
                { 8,  (
                        """
                        :CH1:SCALE 2.0E-2;COUPLING AC;
                        *OPC?
                        """,
                        "9")  },
                { 9,  (
                        """
                        :CH1:COUPLING DC;
                        *OPC?
                        """,
                        "10")  },
                { 10,  (
                        """
                        :CH1:SCALE 1.0E0;
                        :TRIGGER:MAIN:LEVEL 3.84E0;
                        *OPC?
                        """,
                        "11")  },
                { 11,  (
                        """
                        :HORIZONTAL:MAIN:SCALE 2.50E-4;
                        :CH1:SCALE 1.0E-1;COUPLING AC;
                        :TRIGGER:MAIN:LEVEL 8.44E-1;
                        *OPC?
                        """,
                        "12")  },
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {
            try {
                _mainWindow?.SetButtonEnabled("ProductListButton", false);

                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                FormatSet();

                InstClass[] devices = [_instDmm, _instFg, _instOsc];
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
            HotKeyChekBox.IsChecked = false;
            FgRotateButton.IsEnabled = false;
            FgRotateRButton.IsEnabled = false;
            OscRotateButton.IsEnabled = false;
            OscRotateRButton.IsEnabled = false;
            FgRotateRangeTextBox.Text = string.Empty;
            OscRotateRangeTextBox.Text = string.Empty;
        }

        // DMM切り替え
        private async Task SwitchDmm(DmmInstClass dmmInstClass, string func) {
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
            try {
                if (string.IsNullOrEmpty(fgInstClass.VisaAddress)) { return; }
                VisibleProgressImage(true);

                var fgMaxSettingNumber = _dicSwitchFg.Count;
                fgInstClass.SettingNumber = (fgInstClass.SettingNumber + (isNext ? 1 : -1) + fgMaxSettingNumber) % fgMaxSettingNumber;

                fgInstClass.InstCommand = _dicSwitchFg[fgInstClass.SettingNumber].cmd;

                if (fgInstClass.InstCommand == string.Empty) { return; }

                await MainWindow.RotationFgAsync(fgInstClass);

                FgRotateRangeTextBox.Text = _dicSwitchFg[fgInstClass.SettingNumber].text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        // OSC切り替え
        private async void RotationOsc(OscInstClass oscInstClass, bool isNext) {
            try {
                if (string.IsNullOrEmpty(oscInstClass.VisaAddress)) { return; }
                VisibleProgressImage(true);

                var oscMaxSettingNumber = _dicSwitchOsc.Count;
                oscInstClass.SettingNumber = (oscInstClass.SettingNumber + (isNext ? 1 : -1) + oscMaxSettingNumber) % oscMaxSettingNumber;

                var dic = isNext ? _dicSwitchOsc : _dicSwitchROsc;
                oscInstClass.InstCommand = dic[oscInstClass.SettingNumber].cmd;

                if (oscInstClass.InstCommand == string.Empty) { return; }

                await MainWindow.RotationOscAsync(oscInstClass);

                OscRotateRangeTextBox.Text = _dicSwitchOsc[oscInstClass.SettingNumber].text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
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

    }
}
