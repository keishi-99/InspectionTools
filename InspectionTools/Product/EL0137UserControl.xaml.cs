using InspectionTools.Common;
using System.Data;
using System.Windows;
using System.Windows.Threading;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;
using static InspectionTools.MainWindow;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.Product {
    /// <summary>
    /// EL0137UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL0137UserControl : UserControl, IMainWindowAware {

        private MainWindow? _mainWindow;
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DmmInstClass _instDmm = new();
        private readonly OscInstClass _instOsc = new();

        public EL0137UserControl() {
            InitializeComponent();
        }

        private Dictionary<int, (string cmd, string text)> _dicSwitchOsc = [];

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
            MainWindow.GetVisaAddress(_instOsc, OscComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicSwitchOsc = new Dictionary<int, (string cmd, string text)>
            {
                { 0,(
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
                        "OC入力回路"
                    )
                },
                { 1,(
                        """
                        :CURSOR:HBARS:POSITION1 1.6E0;POSITION2 1.6E0;
                        *OPC?
                        """,
                        "電流入力回路"
                    )
                },
                { 2,(
                        """
                        :CURSOR:HBARS:POSITION1 1.3E0;POSITION2 1.3E0;
                        *OPC?
                        """,
                        "電圧入力回路"
                    )
                },
                { 3,(
                        """
                        :CH1:SCALE 1.0E0;
                        :TRIGGER:MAIN:LEVEL 1.0E0;EDGE:SOURCE CH1;
                        :CURSOR:HBARS:POSITION1 3.0E0;POSITION2 2.0E0;
                        :MEASUREMENT:MEAS2:TYPE MINIMUM;SOURCE CH1;
                        *OPC?
                        """,
                        "波高値"
                    )
                },
                { 4,(
                        """
                        :SELECT:CH1 1;CH2 1;
                        :CURSOR:SELECT:SOURCE CH1;
                        :MEASUREMENT:MEAS1:TYPE NWIDTH;SOURCE CH1;
                        :MEASUREMENT:MEAS2:TYPE PWIDTH;SOURCE CH2;
                        *OPC?
                        """,
                        "1/1分周 CH1,2"
                    )
                },
                { 5,(
                        """
                        :SELECT:CH1 1;CH2 1;
                        :CURSOR:SELECT:SOURCE CH1;
                        :MEASUREMENT:MEAS1:TYPE PWIDTH;SOURCE CH1;
                        :MEASUREMENT:MEAS2:TYPE PWIDTH;SOURCE CH2;
                        *OPC?
                        """,
                        "1/2~100分周 CH1,2"
                    )
                }
            };
        }
        // 機器初期設定
        private void FormatSet() {
            _instDmm.InstCommand = _instDmm.SignalType switch {
                1 => "*RST,F5,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?",
                _ => string.Empty
            };
            _instOsc.InstCommand = _instOsc.SignalType switch {
                2 =>
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
                FormatSet();

                InstClass[] devices = [_instDmm, _instOsc];
                await Task.Run(() => {
                    foreach (var device in devices) {
                        MainWindow.ConnectDevice(device);
                    }
                });

                if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                    OscRotateRangeTextBox.Text = "OC入力回路";
                    OscRotateButton.IsEnabled = true;
                }

                DmmComboBox.IsEnabled = false;
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

            _instDmm.ResetProperties();
            _instOsc.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
            DmmComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyChekBox.IsChecked = false;
            OscRotateButton.IsEnabled = false;
            OscRotateRangeTextBox.Text = string.Empty;
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
        // OSC切り替え
        private async void RotationOsc(OscInstClass oscInstClass, bool isNext) {
            try {
                if (string.IsNullOrEmpty(oscInstClass.VisaAddress)) { return; }
                VisibleProgressImage(true);

                var oscMaxSettingNumber = _dicSwitchOsc.Count;
                oscInstClass.SettingNumber = (oscInstClass.SettingNumber + (isNext ? 1 : -1) + oscMaxSettingNumber) % oscMaxSettingNumber;

                oscInstClass.InstCommand = _dicSwitchOsc[oscInstClass.SettingNumber].cmd;

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

        // DMM01測定値コピー
        private async void ActionHotkeyColon() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000000).ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private async void ActionHotkeyNumDivide() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000000).ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSCローテーション
        private void ActionHotkeyBracketR() {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, true);
        }
        private void ActionHotkeyNumMultiply() {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, true);
        }
        // OSC meas1測定値コピー
        private async void ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadOsc(_instOsc, 1);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
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
                var output = await ReadOsc(_instOsc, 1);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
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
                sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
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

        // HotKeyの登録
        private void SetHotKey() {
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
        private static void ClearHotKey() {
            MainWindow.ClearHotKey();
        }

        // イベントハンドラ
        private void UserControl_Loaded(object sender, RoutedEventArgs e) { LoadEvents(); }
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }

        private void OscRotateButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, true);
        }


    }
}
