using InspectionTools.Common;
using System.Data;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;
using static InspectionTools.MainWindow;
using Brushes = System.Windows.Media.Brushes;
using MessageBox = System.Windows.MessageBox;
using RadioButton = System.Windows.Controls.RadioButton;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.Product {
    /// <summary>
    /// MassFlowUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class MassFlowUserControl : UserControl, IMainWindowAware {

        private MainWindow? _mainWindow;
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DcsInstClass _instDcs = new();
        private readonly DmmInstClass _instDmm = new();
        private readonly FgInstClass _instFg01 = new();
        private readonly FgInstClass _instFg02_1 = new();
        private readonly FgInstClass _instFg02_2 = new();
        private readonly OscInstClass _instOsc = new();

        public MassFlowUserControl() {
            InitializeComponent();
        }

        private Dictionary<int, (string text, string cmd2, string cmd3)> _dicSwitchDcs = [];
        private Dictionary<int, (string fg01, string fg02, string fg03_1, string fg03_2)> _dicSwitchFg = [];
        private Dictionary<int, (string fg01, string fg02, string fg03_1, string fg03_2)> _dicSwitchRFg = [];
        private Dictionary<int, string> _dicSwitchOsc = [];
        private Dictionary<int, string> _dicSwitchROsc = [];
        private Dictionary<int, string> _dicTextFgOsc = [];

        // FMRemoteのメニューアイテムID
        private const int MenuItemIdA1L = 32806;
        private const int MenuItemIdA1H = 32807;
        private const int MenuItemIdA2L = 32808;
        private const int MenuItemIdA2H = 32809;
        private const int MenuItemIdAnalogTrim = 32815;

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
            MainWindow.UpdateComboBox(DmmComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(Fg01ComboBox, "ファンクションジェネレータ", [3, 4]);
            MainWindow.UpdateComboBox(Fg02_1ComboBox, "ファンクションジェネレータ", [2]);
            MainWindow.UpdateComboBox(Fg02_2ComboBox, "ファンクションジェネレータ", [2]);
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
            MainWindow.GetVisaAddress(_instFg01, Fg01ComboBox);
            MainWindow.GetVisaAddress(_instFg02_1, Fg02_1ComboBox);
            MainWindow.GetVisaAddress(_instFg02_2, Fg02_2ComboBox);
            MainWindow.GetVisaAddress(_instOsc, OscComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            // DCS
            _dicSwitchDcs = new Dictionary<int, (string text, string cmd2, string cmd3)>
            {
                { 0, ("OFF","SVR5,SOV+0,SBY", "F1R5S0.0O0E") },
                { 1, ("2V","SVR5,SOV+2,OPR", "F1R5S2.0O1E") },
                { 2, ("8V","SVR5,SOV+8,OPR", "F1R5S8.0O1E") },
                { 3, ("1V","SVR5,SOV+1,OPR", "F1R5S1.0O1E") },
                { 4, ("7V","SVR5,SOV+7,OPR", "F1R5S7.0O1E") },
            };

            // FG
            _dicSwitchFg = new Dictionary<int, (string fg01, string fg02, string fg03_1, string fg03_2)>() {
                {
                    0,(
                    ":CHAN 2;:OUTP:STAT OFF;:CHAN 1;:OUTP:STAT OFF;:VOLT:OFFS 3.0;",
                    ":OUTP1:STAT OFF;:SOUR1:VOLT:OFFS 3.0;:OUTP2:STAT OFF;*OPC?",
                    ":OUTP OFF;:VOLT:OFFS 3.0;*OPC?",
                    ":OUTP OFF;*OPC?"
                    )
                },
                {
                    1,(
                    ":CHAN 2;:OUTP:STAT ON;:CHAN 1;:OUTP:STAT ON;",
                    ":OUTP1:STAT ON;:OUTP2:STAT ON;*OPC?",
                    ":OUTP ON;*OPC?",
                    ":OUTP ON;*OPC?"
                    )
                },
                {
                    2,(
                    ":CHAN 1;:OUTP:STAT OFF;:VOLT 5.0;:OUTP:STAT ON;",
                    ":SOUR1:VOLT:LEV:IMM:AMPL 5.0VPP;*OPC?",
                    ":VOLT 5.0VPP;*OPC?",
                    string.Empty
                    )
                },
                {
                    3,(
                    ":CHAN 1;:OUTP:STAT OFF;:VOLT 0.0;:VOLT:OFFS 4.0;:OUTP:STAT ON;",
                    ":SOUR1:VOLT:LEV:IMM:AMPL 0.0VPP;:SOUR1:VOLT:OFFS 4.0;*OPC?",
                    ":VOLT 0.0VPP;:VOLT:OFFS 4.0;*OPC?",
                    string.Empty
                    )
                },
                {
                    4,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
                {
                    5,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
                {
                    6,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
            };
            _dicSwitchRFg = new Dictionary<int, (string fg01, string fg02, string fg03_1, string fg03_2)>() {
                {
                    0,(
                    ":CHAN 2;:OUTP:STAT OFF;:CHAN 1;:OUTP:STAT OFF;",
                    ":OUTP1:STAT OFF;:OUTP2:STAT OFF;*OPC?",
                    ":OUTP OFF;*OPC?",
                    ":OUTP OFF;*OPC?"
                    )
                },
                {
                    1,(
                    ":CHAN 1;:OUTP:STAT OFF;:VOLT 0.0;:OUTP:STAT ON;",
                    ":SOUR1:VOLT:LEV:IMM:AMPL 0.0VPP;*OPC?",
                    ":VOLT 0.0VPP;*OPC?",
                    string.Empty
                    )
                },
                {
                    2,(
                    ":CHAN 1;:OUTP:STAT OFF;:VOLT 5.0;:VOLT:OFFS 3.0;:OUTP:STAT ON;",
                    ":SOUR1:VOLT:LEV:IMM:AMPL 5.0VPP;:SOUR1:VOLT:OFFS 3.0;*OPC?",
                    ":VOLT 5.0VPP;:VOLT:OFFS 3.0;*OPC?",
                    string.Empty
                    )
                },
                {
                    3,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
                {
                    4,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
                {
                    5,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
                {
                    6,(
                    ":CHAN 2;:OUTP:STAT ON;:CHAN 1;:OUTP:STAT OFF;:VOLT:OFFS 4.0;:OUTP:STAT ON;",
                    ":OUTP1:STAT ON;:OUTP2:STAT ON;:SOUR1:VOLT:OFFS 4.0;*OPC?",
                    ":OUTP ON;:VOLT:OFFS 4.0;*OPC?",
                    ":OUTP ON;*OPC?"
                    )
                },
            };

            // OSC
            _dicSwitchOsc = new Dictionary<int, string> {
                { 0, string.Empty },
                {
                    1,
                    """
                    :CH2:SCALE 5.0E0;
                    :CH3:SCALE 5.0E0;
                    :CH4:SCALE 5.0E0;
                    :HORIZONTAL:MAIN:SCALE 2.5E-3;
                    :TRIGGER:MAIN:LEVEL 2.0E0;
                    *OPC?
                    """
                },
                {
                    2,
                    """
                    :SELECT:CH1 1;CH2 0;CH4 0;
                    :MEASUREMENT:MEAS1:TYPE FREQUENCY;SOURCE CH1;
                    :MEASUREMENT:MEAS2:TYPE PK2PK;SOURCE CH1;
                    *OPC?
                    """
                },
                {
                    3,
                    """
                    :SELECT:CH1 0;CH3 1;
                    :MEASUREMENT:MEAS1:TYPE NONE;SOURCE CH3;
                    :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH3;
                    :HORIZONTAL:MAIN:SCALE 5.0E-2;
                    :TRIGGER:MAIN:EDGE:SOURCE CH3;
                    :TRIGGER:MAIN:LEVEL 1.0E1;
                    *OPC?
                    """
                },
                {
                    4,
                    """
                    :MEASUREMENT:MEAS2:TYPE MINIMUM;SOURCE CH3;
                    :ACQUIRE:MODE AVERAGE;NUMAVG 4;
                    :CH3:SCALE 5.0E-2;
                    :HORIZONTAL:MAIN:SCALE 2.5E-3;
                    :TRIGGER:MAIN:LEVEL 5.0E-2;
                    *OPC?
                    """
                },
                {
                    5,
                    """
                    :MEASUREMENT:MEAS2:TYPE FREQUENCY;SOURCE CH3;
                    :ACQUIRE:MODE SAMPLE;NUMAVG 16;
                    :CH3:SCALE 5.0E0;
                    :HORIZONTAL:MAIN:SCALE 5.0E-2;
                    :TRIGGER:MAIN:LEVEL 1.0E1;
                    *OPC?
                    """
                },
                {
                    6,
                    """
                    :SELECT:CH2 1;CH3 0;CH4 1;
                    :MEASUREMENT:MEAS1:TYPE MAXIMUM;SOURCE CH2;
                    :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH4;
                    :ACQUIRE:MODE SAMPLE;NUMAVG 16;
                    :CH2:SCALE 1.0E0;
                    :CH4:SCALE 1.0E0;
                    :HORIZONTAL:MAIN:SCALE 1.0E-2;
                    :TRIGGER:MAIN:EDGE:SOURCE CH1;
                    :TRIGGER:MAIN:LEVEL 0.0E0;
                    *OPC?
                    """
                }
            };
            _dicSwitchROsc = new Dictionary<int, string> {
                { 0, string.Empty },
                {
                    1,
                    """
                    :SELECT:CH1 0;CH2 1;CH3 0;CH4 1;
                    :MEASUREMENT:MEAS1:TYPE MAXIMUM;SOURCE CH2;
                    :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH4;
                    *OPC?
                    """
                },
                {
                    2,
                    """
                    :SELECT:CH1 1;CH3 0;CH2 0;CH4 0;
                    :MEASUREMENT:MEAS1:TYPE FREQUENCY;SOURCE CH1;
                    :MEASUREMENT:MEAS2:TYPE PK2PK;SOURCE CH1;
                    :CH2:SCALE 5.0E0;
                    :CH4:SCALE 5.0E0;
                    :HORIZONTAL:MAIN:SCALE 2.5E-3;
                    :TRIGGER:MAIN:EDGE:SOURCE CH1;
                    :TRIGGER:MAIN:LEVEL 2.0E0;
                    *OPC?
                    """
                },
                {
                    3,
                    """
                    :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH3;
                    :ACQUIRE:MODE SAMPLE;NUMAVG 16;
                    :CH3:SCALE 5.0E0;
                    :HORIZONTAL:MAIN:SCALE 5.0E-2;
                    :TRIGGER:MAIN:LEVEL 1.0E1;
                    *OPC?
                    """
                },
                {
                    4,
                    """
                    :MEASUREMENT:MEAS2:TYPE MINIMUM;SOURCE CH3;
                    :ACQUIRE:MODE AVERAGE;NUMAVG 4;
                    :CH3:SCALE 5.0E-2;
                    :HORIZONTAL:MAIN:SCALE 2.5E-3;
                    :TRIGGER:MAIN:LEVEL 5.0E-2;
                    *OPC?
                    """
                },
                {
                    5,
                    """
                    :SELECT:CH1 0;CH2 0;CH3 1;CH4 0;
                    :MEASUREMENT:MEAS1:TYPE NONE;SOURCE CH3;
                    :MEASUREMENT:MEAS2:TYPE FREQUENCY;SOURCE CH3;
                    :CH3:SCALE 5.0E0;:HORIZONTAL:MAIN:SCALE 5.0E-2;
                    :TRIGGER:MAIN:EDGE:SOURCE CH3;
                    :TRIGGER:MAIN:LEVEL 1.0E1;
                    *OPC?
                    """
                },
                {
                    6,
                    """
                    :CH2:SCALE 1.0E0;
                    :CH3:SCALE 5.0E-2;
                    :CH4:SCALE 1.0E0;
                    :HORIZONTAL:MAIN:SCALE 1.0E-2;
                    :TRIGGER:MAIN:LEVEL 0.0E0;
                    *OPC?
                    """
                }
            };

            _dicTextFgOsc = new Dictionary<int, string> {
                { 0, "0:OFF" },
                { 1, "1:小流量での動作確認" },
                { 2, "2:アンプ動作の確認" },
                { 3, "3:パルス出力回路の確認(最大値)" },
                { 4, "4:パルス出力回路の確認(最小値)" },
                { 5, "5:パルス出力回路の確認(周波数)" },
                { 6, "6:大流量での動作確認" },
            };
        }
        // 機器初期設定
        private void FormatSet() {
            // DCS
            _instDcs.InstCommand = _instDcs.SignalType switch {
                2 => "SVR5,SOV+0,SBY,*OPC?",
                3 => "RCF1R5S0.0O0E",
                _ => string.Empty,
            };
            // DMM
            _instDmm.InstCommand = _instDmm.SignalType switch {
                1 => "*RST,F5,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?",
                _ => string.Empty,
            };
            // FG
            _instFg01.InstCommand = _instFg01.SignalType switch {
                3 => "*RST;:CHAN 2;:MODE NORM;:FUNC:SHAP FSQU;:FREQ 100;:VOLT 0.0;:VOLT:OFFS 2.0;:OUTP:STAT OFF;:CHAN 1;:CHAN:MODE IND;:MODE NORM;:FUNC:SHAP FSQU;:FREQ 100;:VOLT 0.0;:VOLT:OFFS 3.0;:OUTP:STAT OFF;",
                4 => "*RST;:CHAN:MODE IND;:SOUR1:FUNC:SHAP SQU;SQU:DCYC 50PCT;:SOUR1:VOLT:LEV:IMM:AMPL 0.0VPP;:SOUR1:FREQ:CW 100HZ;:SOUR1:VOLT:OFFS 3.0;:SOUR2:FUNC:SHAP SQU;SQU:DCYC 50PCT;:SOUR2:VOLT:LEV:IMM:AMPL 0.0VPP;:SOUR2:FREQ:CW 100HZ;:SOUR2:VOLT:OFFS 2.0;:OUTP1:STAT OFF;:OUTP2:STAT OFF;*OPC?",
                _ => string.Empty,
            };
            _instFg02_1.InstCommand = _instFg02_1.SignalType switch {
                2 => "*RST;:FUNC SQU;:FREQ 100;:VOLT 0.0VPP;:VOLT:OFFS 3.0;*OPC?",
                _ => string.Empty,
            };
            _instFg02_2.InstCommand = _instFg02_2.SignalType switch {
                2 => "*RST;:FUNC SQU;:FREQ 100;:VOLT 0.0VPP;:VOLT:OFFS 2.0;*OPC?",
                _ => string.Empty,
            };
            // OSC
            _instOsc.InstCommand = _instOsc.SignalType switch {
                2 =>
                """
                *RST;
                :HEADER 0;
                :ACQUIRE:STATE 1;STOPAFTER RUNSTOP;
                :SELECT:CH1 0;CH2 1;CH4 1;
                :CH1:PROBE 1.0E1;CURRENTPROBE 1.0E1;SCALE 1.0E0;POSITION -2.0E0;
                :CH2:PROBE 1.0E1;CURRENTPROBE 1.0E1;SCALE 5.0E0;POSITION -2.0E0;
                :CH3:PROBE 1.0E1;CURRENTPROBE 1.0E1;SCALE 5.0E0;POSITION -3.0E0;
                :CH4:PROBE 1.0E1;CURRENTPROBE 1.0E1;SCALE 5.0E0;POSITION -3.0E0;
                :HORIZONTAL:MAIN:SCALE 2.5E-3;
                :HORIZONTAL:DELAY:SCALE 5.0E-9;POSITION 5.0E-4;
                :HORIZONTAL:SCALE 2.5E-3;
                :TRIGGER:MAIN:HOLDOFF:VALUE 5.0E-7;
                :TRIGGER:MAIN:VIDEO:SYNC FIELD;LINE 1;
                :TRIGGER:MAIN:LEVEL 2.0E0;
                :MEASUREMENT:MEAS1:TYPE MAXIMUM;SOURCE CH2;
                :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH4;
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

                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                CheckFgId();
                FormatSet();

                var value = int.Parse(FgNumberComboBox.Text);

                InstClass[] devices = value switch {
                    1 => [_instDcs, _instDmm, _instFg01, _instOsc],
                    2 => [_instDcs, _instDmm, _instFg02_1, _instFg02_2, _instOsc],
                    _ => [_instDcs, _instDmm, _instOsc] // 1, 2 以外の値の場合のデフォルト
                };

                var tasks = devices.Select(device => MainWindow.ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);

                if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                    DcsNumberTextBox.Text = _dicSwitchDcs[0].text;
                }
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress) || !string.IsNullOrEmpty(_instFg01.VisaAddress) || !string.IsNullOrEmpty(_instFg02_1.VisaAddress) || !string.IsNullOrEmpty(_instFg02_2.VisaAddress)) {
                    FgOscRotationButton.IsEnabled = true;
                    FgOscRotationRButton.IsEnabled = true;
                    FgOscNumberTextBox.Text = _dicTextFgOsc[0];
                }

                DcsComboBox.IsEnabled = false;
                DmmComboBox.IsEnabled = false;
                FgNumberComboBox.IsEnabled = false;
                Fg01ComboBox.IsEnabled = false;
                Fg02_1ComboBox.IsEnabled = false;
                Fg02_2ComboBox.IsEnabled = false;
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
        // FGのIDチェック処理
        private void CheckFgId() {
            var indices = new[] { _instFg02_1.Index, _instFg02_2.Index }
                .Where(i => i >= 1); // 未選択(0以下)は無視

            if (indices.Count() != indices.Distinct().Count()) {
                throw new Exception("同じ測定器が選択されています。");
            }
        }

        // 解除
        private void Release() {
            VisibleProgressImage(false);

            _instDcs.ResetProperties();
            _instDmm.ResetProperties();
            _instFg01.ResetProperties();
            _instFg02_1.ResetProperties();
            _instFg02_2.ResetProperties();
            _instOsc.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
            DcsComboBox.IsEnabled = true;
            DmmComboBox.IsEnabled = true;
            FgNumberComboBox.IsEnabled = true;
            Fg01ComboBox.IsEnabled = true;
            Fg02_1ComboBox.IsEnabled = true;
            Fg02_2ComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyChekBox.IsChecked = false;

            DcsNumberTextBox.Text = string.Empty;
            DcsOffButton.IsEnabled = true;
            Dcs2VButton.IsEnabled = true;
            Dcs8VButton.IsEnabled = true;
            Dcs1VButton.IsEnabled = true;
            Dcs7VButton.IsEnabled = true;

            FgOscRotationButton.IsEnabled = false;
            FgOscRotationRButton.IsEnabled = false;

            FgOscNumberTextBox.Text = string.Empty;
        }

        // DCS切り替え
        private async void SwitchDcs(int i) {
            VisibleProgressImage(true);
            try {
                if (_instFg01.SettingNumber != 0 || _instFg02_1.SettingNumber != 0 || _instOsc.SettingNumber != 0) {
                    MessageBox.Show("FG&OSCがONになっています。");
                    return;
                }

                var hWnd = IntPtr.Zero;
                if (i != 0) {
                    hWnd = FindWindow(null, "マルチ流量計渦 [V01.08]");
                    SetForegroundWindow(hWnd);
                }

                await SwitchDcsAsync(_instDcs, i);

                // テキストボックス更新
                DcsNumberTextBox.Text = _dicSwitchDcs[i].text;

                // OFF以外はボタン無効化
                FgOscRotationButton.IsEnabled = (i == 0) &&
                    (
                    !string.IsNullOrEmpty(_instFg01.VisaAddress) ||
                    !string.IsNullOrEmpty(_instFg02_1.VisaAddress) ||
                    !string.IsNullOrEmpty(_instOsc.VisaAddress)
                    );
                FgOscRotationRButton.IsEnabled = (i == 0) &&
                    (
                    !string.IsNullOrEmpty(_instFg01.VisaAddress) ||
                    !string.IsNullOrEmpty(_instFg02_1.VisaAddress) ||
                    !string.IsNullOrEmpty(_instOsc.VisaAddress)
                    );

                // 非同期処理完了後、UIスレッドでPostMessageを送信
                var menuItemID = i switch {
                    0 => 0,
                    1 => MenuItemIdA1L,
                    2 => MenuItemIdA1H,
                    3 => MenuItemIdA2L,
                    4 => MenuItemIdA2H,
                    _ => throw new ArgumentOutOfRangeException(nameof(i), "無効なメニューアイテムID")
                };
                if (hWnd != IntPtr.Zero) {
                    _ = PostMessage(hWnd, WmCommand, menuItemID, 0);
                }
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        private async Task SwitchDcsAsync(InstClass instClass, int i) {
            if (string.IsNullOrEmpty(instClass.VisaAddress)) { return; }

            instClass.SettingNumber = i;
            var settingNumber = instClass.SettingNumber;

            instClass.InstCommand = instClass.SignalType switch {
                2 => _dicSwitchDcs[settingNumber].cmd2,
                3 => _dicSwitchDcs[settingNumber].cmd3,
                _ => throw new ApplicationException(),
            };

            if (string.IsNullOrEmpty(instClass.InstCommand) || instClass.UsbDev is null) { return; }
            await MainWindow.ConnectDeviceAsync(instClass);
        }

        // FG&OSCローテーション
        private async void RotationFgOsc(bool isNext) {
            try {
                VisibleProgressImage(true);

                if (_instDcs.SettingNumber != 0) {
                    MessageBox.Show("DCSがONになっています。", "", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var value = int.Parse(FgNumberComboBox.Text);
                await (value switch {
                    1 => Task.Run(() => RotationFgOscAsync(_instFg01, _instOsc, isNext)),
                    2 => Task.Run(() => RotationFgOscAsync(_instFg02_1, _instFg02_2, _instOsc, isNext)),
                    _ => Task.CompletedTask // どれにも該当しない場合はすぐに完了するタスク
                });

                // テキストボックス更新
                var settingNumber = new InstClass[] { _instOsc, _instFg01, _instFg02_1 }
                  .FirstOrDefault(x => x.SignalType != 0)
                  ?.SettingNumber ?? 0;

                FgOscNumberTextBox.Text = _dicTextFgOsc[settingNumber];

                bool isSettingZero = _instOsc.SettingNumber == 0;
                DcsOffButton.IsEnabled = isSettingZero;
                Dcs2VButton.IsEnabled = isSettingZero;
                Dcs8VButton.IsEnabled = isSettingZero;
                Dcs1VButton.IsEnabled = isSettingZero;
                Dcs7VButton.IsEnabled = isSettingZero;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        private async Task RotationFgOscAsync(FgInstClass fgInstClass, OscInstClass oscInstClass, bool isNext) {
            var fgMaxSettingNumber = _dicSwitchFg.Count;
            var oscMaxSettingNumber = _dicSwitchOsc.Count;

            fgInstClass.SettingNumber = (fgInstClass.SettingNumber + (isNext ? 1 : -1) + fgMaxSettingNumber) % fgMaxSettingNumber;
            oscInstClass.SettingNumber = (oscInstClass.SettingNumber + (isNext ? 1 : -1) + oscMaxSettingNumber) % oscMaxSettingNumber;

            await ConnectAndSendCommand(fgInstClass, GetFgCommand(fgInstClass, isNext));
            await ConnectAndSendCommand(oscInstClass, GetOscCommand(oscInstClass, isNext));
        }
        private async Task RotationFgOscAsync(FgInstClass fgInstClass2_1, FgInstClass fgInstClass2_2, OscInstClass oscInstClass, bool isNext) {
            var fgMaxSettingNumber = _dicSwitchFg.Count;
            var oscMaxSettingNumber = _dicSwitchOsc.Count;

            fgInstClass2_1.SettingNumber = (fgInstClass2_1.SettingNumber + (isNext ? 1 : -1) + fgMaxSettingNumber) % fgMaxSettingNumber;
            oscInstClass.SettingNumber = (oscInstClass.SettingNumber + (isNext ? 1 : -1) + oscMaxSettingNumber) % oscMaxSettingNumber;

            await ConnectAndSendCommand(fgInstClass2_1, GetFgCommand(fgInstClass2_1, isNext));
            await ConnectAndSendCommand(fgInstClass2_2, GetFgCommand(fgInstClass2_2, isNext));
            await ConnectAndSendCommand(oscInstClass, GetOscCommand(oscInstClass, isNext));
        }
        private string GetFgCommand(FgInstClass fgInstClass, bool isNext) {
            var dic = isNext ? _dicSwitchFg : _dicSwitchRFg;
            var (fg01, fg02, fg03_1, fg03_2) = dic[fgInstClass.SettingNumber];

            return fgInstClass.SignalType switch {
                2 => fgInstClass.Tag switch {
                    "FG-1" => fg03_1,
                    "FG-2" => fg03_2,
                    _ => string.Empty
                },
                3 => fg01,
                4 => fg02,
                _ => string.Empty
            };
        }
        private string GetOscCommand(OscInstClass oscInstClass, bool isNext) {
            var dic = isNext ? _dicSwitchOsc : _dicSwitchROsc;
            return dic[oscInstClass.SettingNumber];
        }
        private static async Task ConnectAndSendCommand(InstClass instClass, string command) {
            if (!string.IsNullOrEmpty(instClass.VisaAddress) && !string.IsNullOrEmpty(command) && instClass.UsbDev is not null) {
                instClass.InstCommand = command;
                await MainWindow.ConnectDeviceAsync(instClass);
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

        // Serialロック
        private void SerialLockToggle() { SerialTextBox.IsEnabled = !SerialLockCheckBox.IsChecked ?? false; }
        // Serialインクリメント
        private void SerialIncrement(int i) {
            try {
                var serialText = SerialTextBox.Text;
                if (serialText.Length == 0) {
                    return;
                }
                if (serialText.Length != 8) {
                    throw new Exception("シリアル文字数が一致しません。");
                }
                var middleDigits = serialText.Substring(4, 3);
                if (int.TryParse(middleDigits, out var currentValue)) {
                    // 0〜999 の範囲で循環
                    currentValue = (currentValue + i + 1000) % 1000;

                    // 3桁に0埋め
                    var newValue = currentValue.ToString("000");

                    // 元のシリアル番号を再構築
                    var sb = new System.Text.StringBuilder();
                    sb.Append(serialText.AsSpan(0, 4));  // 先頭4文字
                    sb.Append(newValue);                  // 中間3桁
                    sb.Append(serialText.AsSpan(7));     // 7文字目以降

                    SerialTextBox.Text = sb.ToString();
                }
                else {
                    throw new Exception("シリアルの中間値が数値に変換できません。");
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // キャリブレーション値取得
        private static (string a1l, string a1h, string a2l, string a2h) GetCalibration() {
            var hWnd = FindWindow(null, "マルチ流量計渦 [V01.08]");
            if (hWnd == IntPtr.Zero) { return (string.Empty, string.Empty, string.Empty, string.Empty); }

            var childHandles = FindChildWindows(hWnd);

            var a1l = GetMessageText(childHandles[332], WmGetText, 256);
            var a1h = GetMessageText(childHandles[342], WmGetText, 256);
            var a2l = GetMessageText(childHandles[352], WmGetText, 256);
            var a2h = GetMessageText(childHandles[362], WmGetText, 256);

            return (a1l, a1h, a2l, a2h);
        }
        // 受信ボタン実行
        private static void ReceiveData(IntPtr hWnd) {
            var childHandles = FindChildWindows(hWnd);
            _ = PostMessage(childHandles[9], BmClick, 0, 0);    // 受信ボタンクリック

            var hWnd3 = IntPtr.Zero;
            var timeOut = 0;
            while (hWnd3 == IntPtr.Zero) {
                hWnd3 = FindWindow(null, "FMRemote2014");
                timeOut += 1;
                Thread.Sleep(100);
                if (timeOut > 30) {
                    MessageBox.Show("タイムアウトしました");
                    return;
                }
            }
            var childHandles2 = FindChildWindows(hWnd3);
            _ = PostMessage(childHandles2[0], BmClick, 0, 0);    // メッセージウィンドウ「はい」クリック
        }
        // アナログトリムオープン
        private async void AnalogTrimOpen(IntPtr hWnd) {
            if (MainWindow.IsProcessing) { return; }
            VisibleProgressImage(true);

            _ = PostMessage(hWnd, WmCommand, MenuItemIdAnalogTrim, 0);

            var hWnd3 = IntPtr.Zero;
            var timeOut = 0;
            while (hWnd3 == IntPtr.Zero) {
                hWnd3 = FindWindow(null, "FMRemote2014");
                timeOut += 1;
                Thread.Sleep(100);
                if (timeOut > 30) {
                    MessageBox.Show("タイムアウトしました");
                    VisibleProgressImage(false);
                    return;
                }
            }

            var childHandles = FindChildWindows(hWnd3);
            _ = PostMessage(childHandles[0], BmClick, 0, 0);    // メッセージウィンドウ「はい」クリック

            var hWnd2 = GetForegroundWindow();
            if (hWnd2 == IntPtr.Zero) { return; }

            (hWnd2, var windowText) = GetActiveWindow;

            while (windowText != "ANLOG TRIM") {
                await Task.Delay(500);
                (hWnd2, windowText) = GetActiveWindow;
            }

            await Task.Delay(500);

            ActivateAndBringToFront(hWnd2);

            var childHandles2 = FindChildWindows(hWnd2);
            _ = PostMessage(childHandles2[5], WmLButtonDown, 0, 0);    // 入力欄フォーカス

            VisibleProgressImage(false);
        }
        // アナログトリムスタート
        private async void AnalogTrimStart(IntPtr hWnd) {
            if (MainWindow.IsProcessing) { return; }
            VisibleProgressImage(true);

            var childHandles = FindChildWindows(hWnd);
            _ = PostMessage(childHandles[7], BmClick, 0, 0);    // STARTクリック

            await Task.Delay(500);

            ActivateAndBringToFront(hWnd);

            _ = PostMessage(childHandles[5], WmLButtonDown, 0, 0);  // 入力欄フォーカス
            _ = PostMessage(childHandles[5], EmSetSel, 0, -1);      // 入力欄全選択

            VisibleProgressImage(false);
        }
        // アナログトリムクローズ
        private void AnalogTrimClose(IntPtr hWnd) {
            if (MainWindow.IsProcessing) { return; }
            VisibleProgressImage(true);

            var childHandles = FindChildWindows(hWnd);
            _ = PostMessage(childHandles[0], BmClick, 0, 0);    // CLOSEクリック

            VisibleProgressImage(false);
        }
        // アナログトリム4mA
        private async void AnalogTrim4mA(IntPtr hWnd) {
            if (MainWindow.IsProcessing) { return; }
            VisibleProgressImage(true);

            var childHandles = FindChildWindows(hWnd);
            _ = PostMessage(childHandles[2], BmClick, 0, 0);    // 4mAクリック

            await Task.Delay(500);

            ActivateAndBringToFront(hWnd);

            _ = PostMessage(childHandles[5], WmLButtonDown, 0, 0);  // 入力欄フォーカス
            _ = PostMessage(childHandles[5], EmSetSel, 0, -1);      // 入力欄全選択

            VisibleProgressImage(false);
        }
        // アナログトリム20mA
        private async void AnalogTrim20mA(IntPtr hWnd) {
            if (MainWindow.IsProcessing) { return; }
            VisibleProgressImage(true);

            var childHandles = FindChildWindows(hWnd);
            _ = PostMessage(childHandles[3], BmClick, 0, 0);    // 20mAクリック

            await Task.Delay(500);

            ActivateAndBringToFront(hWnd);

            _ = PostMessage(childHandles[5], WmLButtonDown, 0, 0);  // 入力欄フォーカス
            _ = PostMessage(childHandles[5], EmSetSel, 0, -1);      // 入力欄全選択

            VisibleProgressImage(false);
        }
        // アクティブウィンドウ取得
        private static (IntPtr hWnd, string windowText) GetActiveWindow {
            get {
                var hWnd = GetForegroundWindow();
                var windowText = GetWindowText(hWnd);

                return (hWnd, windowText);
            }
        }

        [GeneratedRegex("^OP-650-10.*xlsm")]
        private static partial Regex MassFlowRegex();

        // アナログトリムスタート または OP-650-102をアクティブ
        private void ActionHotkeyNumAdd() {
            var (hWnd, windowText) = GetActiveWindow;
            switch (windowText.ToString()) {
                case "ANLOG TRIM": {
                        AnalogTrimStart(hWnd);
                        break;
                    }
                default: {
                        //すべてのプロセスを列挙する
                        foreach (var p in Process.GetProcesses()) {// メインウィンドウハンドルが存在し、かつタイトルが空でないことを確認
                            if (p.MainWindowHandle != IntPtr.Zero && !string.IsNullOrEmpty(p.MainWindowTitle)) {
                                //"OP-650-102"がメインウィンドウのタイトルに含まれているか調べる
                                if (MassFlowRegex().IsMatch(p.MainWindowTitle)) {
                                    //ウィンドウをアクティブにする
                                    var excelTitle = p.MainWindowTitle;
                                    var hWnd2 = FindWindow(null, excelTitle);
                                    ActivateAndBringToFront(hWnd2);
                                    break;
                                }
                            }
                        }
                        break;
                    }
            }
        }
        // OP-650-102 をアクティブ
        private static void ActionHotkeyMinus() {
            //すべてのプロセスを列挙する
            foreach (var p in Process.GetProcesses()) {
                // メインウィンドウハンドルが存在し、かつタイトルが空でないことを確認
                if (p.MainWindowHandle != IntPtr.Zero && !string.IsNullOrEmpty(p.MainWindowTitle)) {
                    //"OP-650-102"がメインウィンドウのタイトルに含まれているか調べる
                    if (MassFlowRegex().IsMatch(p.MainWindowTitle)) {
                        //ウィンドウをアクティブにする
                        var excelTitle = p.MainWindowTitle;
                        var hWnd = FindWindow(null, excelTitle);
                        ActivateAndBringToFront(hWnd);
                        break;
                    }
                }
            }
        }
        // マルチ流量計渦 [V01.08] をアクティブ
        private static void ActionHotkeyTilde() {
            var hWnd = FindWindow(null, "マルチ流量計渦 [V01.08]");
            ActivateAndBringToFront(hWnd);
        }
        // アナログトリムクローズ または マルチ流量計渦 [V01.08]をアクティブ
        private void ActionHotkeyNumSubtract() {
            var (hWnd, windowText) = GetActiveWindow;
            switch (windowText.ToString()) {
                case "ANLOG TRIM": {
                        AnalogTrimClose(hWnd);
                        break;
                    }
                default: {
                        string[] windowTitles = ["FMRemote2014", "マルチ流量計渦 [V01.08]"];

                        foreach (string title in windowTitles) {
                            IntPtr hWnd2 = FindWindow(null, title);
                            if (hWnd2 != IntPtr.Zero) {
                                ActivateAndBringToFront(hWnd2);
                                break;
                            }
                        }
                        break;
                    }
            }
        }
        // アナログトリム4mA または FG&OSCローテーション
        private void ActionHotkeyNumDivide() {
            var (hWnd, windowText) = GetActiveWindow;
            switch (windowText.ToString()) {
                case "ANLOG TRIM": {
                        AnalogTrim4mA(hWnd);
                        break;
                    }
                default: {
                        if (MainWindow.IsProcessing) { return; }
                        RotationFgOsc(false);
                        break;
                    }
            }
        }
        // アナログトリム20mA または FG&OSCローテーション
        private void ActionHotkeyNumMultiply() {
            var (hWnd, windowText) = GetActiveWindow;
            switch (windowText.ToString()) {
                case "ANLOG TRIM": {
                        AnalogTrim20mA(hWnd);
                        break;
                    }
                default: {
                        if (MainWindow.IsProcessing) { return; }
                        RotationFgOsc(true);
                        break;
                    }
            }
        }
        // 受信ボタン実行
        private static void ActionHotkeyAtsign() {
            var (hWnd, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                ReceiveData(hWnd);
            }
        }
        // Serialインクリメント＆コピー
        private void ActionHotkeyBracketL() {
            var (_, windowText) = GetActiveWindow;
            if (windowText.ToString() != "名前を付けて保存") {
                return;
            }
            if (string.IsNullOrEmpty(SerialTextBox.Text)) {
                return;
            }

            var sim = new InputSimulator();
            sim.Keyboard.TextEntry(SerialTextBox.Text);
            SerialIncrement(1);
        }
        // キャリブレーション値コピー
        private async void ActionHotkeySemiColon() {
            try {
                VisibleProgressImage(true);

                var (hWnd, windowText) = GetActiveWindow;
                if (!MassFlowRegex().IsMatch(windowText)) {
                    return;
                }
                (var a1l, var a1h, var a2l, var a2h) = GetCalibration();
                if (string.IsNullOrEmpty(a1l)) { a1l = "error"; }
                if (string.IsNullOrEmpty(a1h)) { a1h = "error"; }
                if (string.IsNullOrEmpty(a2l)) { a2l = "error"; }
                if (string.IsNullOrEmpty(a2h)) { a2h = "error"; }

                // 各クリップボード操作とキー送信のブロック
                await PerformClipboardAndSendKeys(a1l);
                await PerformClipboardAndSendKeys(a1h);
                await PerformClipboardAndSendKeys(a2l);
                await PerformClipboardAndSendKeys(a2h);

                static async Task PerformClipboardAndSendKeys(string text) {
                    await Task.Delay(500);
                    var sim = new InputSimulator();
                    sim.Keyboard.TextEntry(text);
                    sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                }
            } finally {
                VisibleProgressImage(false);
            }
        }
        // FG&OSCローテーション
        private void ActionHotkeyColon() {
            if (MainWindow.IsProcessing) { return; }
            RotationFgOsc(false);
        }
        private void ActionHotkeyBracketR() {
            if (MainWindow.IsProcessing) { return; }
            RotationFgOsc(true);
        }
        // DMM値コピー
        private async void ActionHotkeyPeriod() {
            if (MainWindow.IsProcessing) { return; }

            var output = await ReadDmm(_instDmm);
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry((output * 1000).ToString("0.0000"));
        }
        // OSC mes1値コピー
        private async void ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            var output = await ReadOsc(_instOsc, 1);
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry(output.ToString("0.0000"));
            sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }
        // OSC mes2値コピー
        private async void ActionHotkeyBackslash() {
            if (MainWindow.IsProcessing) { return; }

            var output = await ReadOsc(_instOsc, 2);
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry(output.ToString("0.0000"));
            sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }
        // DCS切り替え
        private void ActionHotkeyNum0() {
            var (_, windowText) = GetActiveWindow;

            var sim = new InputSimulator();

            switch (windowText.ToString()) {
                case "マルチ流量計渦 [V01.08]": {
                        SwitchDcs(0);
                        break;
                    }
                case "FMRemote2014": {
                        sim.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
                        return;
                    }
                default: {
                        sim.Keyboard.TextEntry("0");
                        break;
                    }
            }
        }
        private void ActionHotkeyNum1() {
            var (_, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                if (MainWindow.IsProcessing) { return; }
                SwitchDcs(1);
                return;
            }
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry("1");
        }
        private void ActionHotkeyNum2() {
            var (_, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                if (MainWindow.IsProcessing) { return; }
                SwitchDcs(2);
                return;
            }
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry("2");
        }
        private void ActionHotkeyNum3() {
            var (_, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                if (MainWindow.IsProcessing) { return; }
                SwitchDcs(3);
                return;
            }
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry("3");
        }
        private void ActionHotkeyNum4() {
            var (_, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                if (MainWindow.IsProcessing) { return; }
                SwitchDcs(4);
                return;
            }
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry("4");
        }
        private void ActionHotkeyNum7() {
            var (hWnd, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                AnalogTrimOpen(hWnd);
                return;
            }
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry("7");
        }


        // HotKeyの登録
        private void SetHotKey() {
            MainWindow.HotkeysList.Clear();

            MainWindow.HotkeysList.AddRange([
                new(ModNone,    HotkeyNumAdd,       ActionHotkeyNumAdd),
                new(ModNone,    HotkeyNumSubtract,  ActionHotkeyNumSubtract),
                new(ModNone,    HotkeyNumDivide,    ActionHotkeyNumDivide),
                new(ModNone,    HotkeyNumMultiply,  ActionHotkeyNumMultiply),
                new(ModNone,    HotkeyMinus,        ActionHotkeyMinus),
                new(ModNone,    HotkeyTilde,        ActionHotkeyTilde),
                new(ModNone,    HotkeyAtsign,       ActionHotkeyAtsign),
                new(ModNone,    HotkeyBracketL,     ActionHotkeyBracketL),
                new(ModNone,    HotkeySemiColon,    ActionHotkeySemiColon),
                new(ModNone,    HotkeyNum0,         ActionHotkeyNum0),
                new(ModNone,    HotkeyNum7,         ActionHotkeyNum7),
                new(ModNone,    HotkeyNum1,         ActionHotkeyNum1),
                new(ModNone,    HotkeyNum2,         ActionHotkeyNum2),
                new(ModNone,    HotkeyNum3,         ActionHotkeyNum3),
                new(ModNone,    HotkeyNum4,         ActionHotkeyNum4),
            ]);
            if (!string.IsNullOrEmpty(_instFg01.VisaAddress) || !string.IsNullOrEmpty(_instFg02_1.VisaAddress) || !string.IsNullOrEmpty(_instFg02_2.VisaAddress) || !string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                     new(ModNone, HotkeyColon, ActionHotkeyColon),
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                MainWindow.HotkeysList.Add(new(ModNone, HotkeyPeriod, ActionHotkeyPeriod));
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                     new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                ]);
            }

            MainWindow.SetHotKey();
        }
        private static void ClearHotKey() {
            MainWindow.ClearHotKey();
        }

        // 特定の親ウィンドウに属するすべての子ウィンドウのハンドルを取得するメソッド
        private static List<IntPtr> FindChildWindows(IntPtr hWndParent) {
            List<IntPtr> childWindows = [];
            EnumChildProc callback = new((hwnd, param) => {
                childWindows.Add(hwnd);
                return true;
            });

            EnumChildWindows(hWndParent, callback, IntPtr.Zero);
            return childWindows;
        }

        // ウィンドウをアクティブにして最前面に持ってくる
        private static void ActivateAndBringToFront(IntPtr hWnd) {
            if (hWnd == IntPtr.Zero) { return; }

            // ウィンドウが最小化されているか確認し、復元する
            if (IsIconic(hWnd)) {
                // SW_RESTORE は最小化されたウィンドウを元のサイズに戻し、アクティブにします。
                ShowWindow(hWnd, SwRestore);
                // 少し待機することで、ウィンドウが表示されるのを確実にする（必要な場合のみ）
                System.Threading.Thread.Sleep(50);
            }
            SetForegroundWindow(hWnd);
        }

        // FGパネルの表示切替
        private void FgPanelVisible() {
            var i = FgNumberComboBox.SelectedIndex;

            (FgNumberGrid1.Visibility, FgNumberGrid2.Visibility) = i switch {
                0 => (Visibility.Visible, Visibility.Collapsed),
                1 => (Visibility.Collapsed, Visibility.Visible),
                _ => (Visibility.Collapsed, Visibility.Collapsed) // 0と1以外の場合
            };
        }

        // イベントハンドラ
        private void UserControl_Loaded(object sender, RoutedEventArgs e) { LoadEvents(); }
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }

        private void DcsOffButton_Click(object sender, RoutedEventArgs e) { SwitchDcs(0); }
        private void Dcs2VButton_Click(object sender, RoutedEventArgs e) { SwitchDcs(1); }
        private void Dcs8VButton_Click(object sender, RoutedEventArgs e) { SwitchDcs(2); }
        private void Dcs1VButton_Click(object sender, RoutedEventArgs e) { SwitchDcs(3); }
        private void Dcs7VButton_Click(object sender, RoutedEventArgs e) { SwitchDcs(4); }

        private void FgOscRotationButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationFgOsc(true);
        }
        private void FgOscRotationRButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationFgOsc(false);
        }

        private void DcsRadioButton_Checked(object sender, RoutedEventArgs e) {
            if (sender is RadioButton checkedRadioButton) {
                checkedRadioButton.FontWeight = FontWeights.Bold;
                checkedRadioButton.Foreground = Brushes.Black;
            }
        }
        private void DcsRadioButton_Unchecked(object sender, RoutedEventArgs e) {
            if (sender is RadioButton checkedRadioButton) {
                checkedRadioButton.FontWeight = FontWeights.Regular;
                checkedRadioButton.Foreground = Brushes.Gray;
            }
        }
        private void FgOscRadioButton_Checked(object sender, RoutedEventArgs e) {
            if (sender is RadioButton checkedRadioButton) {
                checkedRadioButton.FontWeight = FontWeights.Bold;
                checkedRadioButton.Foreground = Brushes.Black;
            }
        }
        private void FgOscRadioButton_Unchecked(object sender, RoutedEventArgs e) {
            if (sender is RadioButton checkedRadioButton) {
                checkedRadioButton.FontWeight = FontWeights.Regular;
                checkedRadioButton.Foreground = Brushes.Gray;
            }
        }

        private void SerialLockCheckBox_Checked(object sender, RoutedEventArgs e) { SerialLockToggle(); }
        private void SerialLockCheckBox_Unchecked(object sender, RoutedEventArgs e) { SerialLockToggle(); }
        private void SerialBack_Click(object sender, RoutedEventArgs e) { SerialIncrement(-1); }
        private void SerialNext_Click(object sender, RoutedEventArgs e) { SerialIncrement(1); }

        private void FgNumberComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (!IsLoaded) return;
            FgPanelVisible();
        }


    }
}
