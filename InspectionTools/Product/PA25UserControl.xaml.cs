using InspectionTools.Common;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Interop;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;
using static InspectionTools.MainMenu.SubMenuUserControl;
using ComboBox = System.Windows.Controls.ComboBox;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.Product {
    /// <summary>
    /// PA25UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class PA25UserControl : UserControl, ISubMenuAware {

        private MainMenu.SubMenuUserControl? _subMenu;
        public void SetSubMenuControl(MainMenu.SubMenuUserControl? subMenu) {
            _subMenu = subMenu;
        }

        private IntPtr _hWnd = IntPtr.Zero;

        private readonly DmmInstClass _instDmm;
        private readonly FgInstClass _instFg;
        private readonly OscInstClass _instOsc;

        public PA25UserControl() {
            InitializeComponent();
            _instDmm = new();
            _instFg = new();
            _instOsc = new();
            LoadEvents();
            // 親ウィンドウを取得
            var parentWindow = Window.GetWindow(this);
            if (parentWindow != null) {
                _hWnd = new WindowInteropHelper(parentWindow).Handle;
            }

            Loaded += (s, e) => AdjustWindowSizeToUserControl();
        }
        private void AdjustWindowSizeToUserControl() {
            var parentWindow = Window.GetWindow(this);
            if (parentWindow != null) {
                parentWindow.SizeToContent = SizeToContent.WidthAndHeight;
            }
        }

        private const int TimeOut = 3;    //タイムアウトまでの時間(sec)

        internal DataTable _dataTable = new();

        private Dictionary<int, (string cmd, string text)> _dicSwitchFg = [];
        private Dictionary<int, (string cmd, string text)> _dicSwitchOsc = [];
        private Dictionary<int, (string cmd, string text)> _dicSwitchROsc = [];

        private readonly List<Hotkey> _hotkeys = [];
        private HwndSource? _source;

        private volatile bool _isProcessing = false;

        private static readonly SemaphoreSlim s_semaphore = new(1, 1); // 最大1つの接続

        // 起動時
        private void LoadEvents() {
            InstListImport();
            FormatSet();
            RegDictionary();
        }
        private void InstListImport() {
            const string XmlFilePath = "VisaAddress.xml";
            if (!System.IO.File.Exists(XmlFilePath)) {
                MessageBox.Show($"{XmlFilePath}が見つかりません。");
                return;
            }

            using DataSet dataSet = new();
            dataSet.ReadXml("VisaAddress.xml");
            _dataTable = dataSet.Tables[0];

            // デジタルマルチメータ、ファンクションジェネレータ、オシロスコープのコンボボックスを更新する
            UpdateComboBox(DmmComboBox, "デジタルマルチメータ", [1, 2], "[DMM]");
            UpdateComboBox(FgComboBox, "ファンクションジェネレータ", [2], "[FG]");
            UpdateComboBox(OscComboBox, "オシロスコープ", [2], "[OSC]");
        }
        private void UpdateComboBox(ComboBox comboBox, string category, List<int> signalTypes, string name) {
            if (_dataTable == null) {
                return;
            }

            var collection = new List<string> { name };

            foreach (var signalType in signalTypes) {
                var rows = _dataTable.Select($"Category = '{category}' AND SignalType = {signalType}");
                foreach (var d in rows) {
                    collection.Add(d["Name"].ToString() ?? string.Empty);
                }
            }
            comboBox.ItemsSource = collection;
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            _isProcessing = isVisible;
            MainGrid.IsEnabled = !isVisible;
            ProgressRing.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }

        // 選択した機器のVisaAddressを取得
        private void SelectInst() {
            GetVisaAddress(_instDmm, DmmComboBox);
            GetVisaAddress(_instFg, FgComboBox);
            GetVisaAddress(_instOsc, OscComboBox);
        }
        private void GetVisaAddress(InstClass instClass, ComboBox comboBox) {
            instClass.ResetProperties();

            instClass.Name = comboBox.Text;
            instClass.Index = comboBox.SelectedIndex;

            if (instClass.Index <= 0) { return; }

            var dRows = _dataTable.Select($"Name = '{instClass.Name}'");
            instClass.Category = dRows[0]["Category"] as string ?? string.Empty;
            instClass.VisaAddress = dRows[0]["VisaAddress"] as string ?? string.Empty;
            instClass.SignalType = dRows[0]["SignalType"] != DBNull.Value ? Convert.ToInt32(dRows[0]["SignalType"]) : 0;
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
                _subMenu?.SetButtonEnabled("ProductListButton", false);

                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                FormatSet();

                InstClass[] devices = [_instDmm, _instFg, _instOsc];
                var tasks = devices.Select(device => ConnectDeviceAsync(device));
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
        // デバイス接続
        private static async Task<string> ConnectDeviceAsync(InstClass instClass) {
            return instClass.Index < 1
                ? ""
                : instClass.SignalType switch {
                    1 => await ConnectDeviceAdcAsync(instClass),
                    2 or 4 => await ConnectDeviceVisaAsync(instClass, true),
                    3 => await ConnectDeviceVisaAsync(instClass, false),
                    _ => throw new ApplicationException(),
                };
        }
        // Visa接続
        private static async Task<string> ConnectDeviceVisaAsync(InstClass instClass, bool hasInput) {
            return await Task.Run(() => {
                using var usbDev = new USBDeviceManager();
                usbDev.OpenDev(instClass.VisaAddress);
                usbDev.OutputDev(instClass.InstCommand);
                return hasInput ? usbDev.InputDev() : "";
            });
        }
        // ADC接続
        private static async Task<string> ConnectDeviceAdcAsync(InstClass instClass) {
            await s_semaphore.WaitAsync();
            try {
                uint hDev = 0;
                var rcvDt = "";
                uint rcvLen = 50;
                var id = uint.Parse(instClass.VisaAddress);
                try {
                    if (AusbWrapper.Start(TimeOut) != 0 || AusbWrapper.Open(ref hDev, id) != 0) { throw new Exception("開始できません"); }
                    if (!string.IsNullOrEmpty(instClass.InstCommand)) {
                        if (AusbWrapper.Write(hDev, instClass.InstCommand) != 0) { throw new Exception("コマンドの送信に失敗しました"); }
                    }
                    if (AusbWrapper.Read(hDev, ref rcvDt, ref rcvLen) != 0) { throw new Exception("メッセージの受信に失敗しました"); }
                } finally {
                    _ = AusbWrapper.Close(hDev);
                    _ = AusbWrapper.End();
                }
                return rcvDt;
            } finally {
                s_semaphore.Release();
            }
        }

        // 解除
        private void Release() {
            VisibleProgressImage(false);

            _instFg.ResetProperties();
            _instOsc.ResetProperties();

            _subMenu?.SetButtonEnabled("ProductListButton", true);
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
        private async Task SwitchDmm(DmmInstClass instClass, string func) {
            try {
                VisibleProgressImage(true);

                (instClass.InstCommand, instClass.CurrentMode) = func switch {
                    "DCI" => instClass.SignalType switch {
                        1 => ("*RST,F5,R6,*OPC?", DmmMode.DCI),
                        2 => ("*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 1E-1;*OPC?", DmmMode.DCI),
                        _ => throw new ApplicationException(),
                    },
                    "DCV" => instClass.SignalType switch {
                        1 => ("*RST,F1,R6,*OPC?", DmmMode.DCV),
                        2 => ("*RST;:INIT:CONT 1;:VOLT:DC:RANG 10;*OPC?", DmmMode.DCV),
                        _ => throw new ApplicationException(),
                    },
                    _ => throw new ApplicationException(),
                };

                await ConnectDeviceAsync(instClass);

            } finally {
                VisibleProgressImage(false);
            }
        }
        // DMM測定値取得
        private async Task<decimal> ReadDmm(DmmInstClass instClass) {
            try {
                VisibleProgressImage(true);

                instClass.InstCommand = instClass.SignalType switch {
                    1 => string.Empty,
                    2 => "FETC?",
                    _ => throw new ApplicationException(),
                };

                var result = await ConnectDeviceAsync(instClass);
                decimal.TryParse(result, NumberStyles.AllowExponent | NumberStyles.Float, CultureInfo.InvariantCulture, out var output);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // FG切り替え
        private async void RotationFg(bool isNext) {
            try {
                if (string.IsNullOrEmpty(_instFg.VisaAddress)) { return; }
                VisibleProgressImage(true);

                await RotationFgAsync(_instFg, isNext);

                FgRotateRangeTextBox.Text = _dicSwitchFg[_instFg.SettingNumber].text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        private async Task RotationFgAsync(InstClass instClass, bool isNext) {

            var fgMaxSettingNumber = _dicSwitchFg.Count;
            _instFg.SettingNumber = (_instFg.SettingNumber + (isNext ? 1 : -1) + fgMaxSettingNumber) % fgMaxSettingNumber;

            instClass.InstCommand = _dicSwitchFg[_instFg.SettingNumber].cmd;

            if (instClass.InstCommand == string.Empty) { return; }

            await ConnectDeviceAsync(instClass);
        }
        // OSC切り替え
        private async void RotationOsc(bool isNext) {
            try {
                if (string.IsNullOrEmpty(_instOsc.VisaAddress)) { return; }
                VisibleProgressImage(true);

                await RotationOscAsync(_instOsc, isNext);

                OscRotateRangeTextBox.Text = _dicSwitchOsc[_instOsc.SettingNumber].text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        private async Task RotationOscAsync(InstClass instClass, bool isNext) {

            var oscMaxSettingNumber = _dicSwitchOsc.Count;
            _instOsc.SettingNumber = (_instOsc.SettingNumber + (isNext ? 1 : -1) + oscMaxSettingNumber) % oscMaxSettingNumber;

            var dic = isNext ? _dicSwitchOsc : _dicSwitchROsc;
            instClass.InstCommand = dic[_instOsc.SettingNumber].cmd;

            if (instClass.InstCommand == string.Empty) { return; }

            await ConnectDeviceAsync(instClass);

        }
        // OSC測定値取得
        private async Task<decimal> ReadOsc(InstClass instClass, int oscMeas) {
            try {
                VisibleProgressImage(true);

                instClass.InstCommand = $"MEASU:MEAS{oscMeas}:VAL?";
                var result = await ConnectDeviceAsync(instClass);
                decimal.TryParse(result, NumberStyles.AllowExponent | NumberStyles.Float, CultureInfo.InvariantCulture, out var output);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }

        // DMM切替(DCV)
        private async void ActionHotkeyComma() {
            if (_isProcessing) { return; }

            try {
                await SwitchDmm(_instDmm, "DCV");
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM切替(DCI)
        private async void ActionHotkeyPeriod() {
            if (_isProcessing) { return; }

            try {
                await SwitchDmm(_instDmm, "DCI");
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM01測定値コピー
        private async void ActionHotkeySlash() {
            if (_isProcessing) { return; }

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
            if (_isProcessing) { return; }
            RotationFg(true);
        }
        private void ActionHotkeyShiftBracketR() {
            if (_isProcessing) { return; }
            RotationFg(false);
        }
        private void ActionHotkeyNumMultiply() {
            if (_isProcessing) { return; }
            RotationFg(true);
        }

        // OSCローテーション
        private void ActionHotkeyColon() {
            if (_isProcessing) { return; }
            RotationOsc(true);
        }
        private void ActionHotkeyShiftColon() {
            if (_isProcessing) { return; }
            RotationOsc(false);
        }
        private void ActionHotkeyNumDivide() {
            if (_isProcessing) { return; }
            RotationOsc(true);
        }

        // OSC meas測定値コピー
        private async void ActionHotkeyBackslash() {
            if (_isProcessing) { return; }

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
            if (_isProcessing) { return; }

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

        // HotkKeyの登録
        private void SetHotKey() {
            _hotkeys.Clear();
            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyComma, ActionHotkeyComma),
                    new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                ]);
            }
            if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                    new(ModShift, HotkeyBracketR, ActionHotkeyShiftBracketR),
                    new(ModNone, HotkeyNumMultiply, ActionHotkeyNumMultiply),
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                    new(ModShift, HotkeyColon, ActionHotkeyShiftColon),
                    new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
                ]);
            }

            // 親ウィンドウを取得
            var parentWindow = Window.GetWindow(this);
            if (parentWindow != null) {
                _hWnd = new WindowInteropHelper(parentWindow).Handle;
                _source = HwndSource.FromHwnd(_hWnd);
                _source.AddHook(HwndHook);
            }

            // ホットキーを登録
            foreach (var hotkey in _hotkeys) {
                RegisterHotKey(_hWnd, hotkey.Id, hotkey.Modifier, (uint)hotkey.VirtualKey);
            }
        }
        private void ClearHotKey() {
            foreach (var hotkey in _hotkeys) {
                UnregisterHotKey(_hWnd, hotkey.Id);
            }
            _hotkeys.Clear();
        }
        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
            if (msg == WmHotKey) {
                int id = wParam.ToInt32();

                var hotkey = _hotkeys.FirstOrDefault(h => h.Id == id);
                hotkey?.Action.Invoke(); // ホットキーに設定されたアクションを実行
                handled = true;
            }
            return IntPtr.Zero;
        }

        // イベントハンドラ
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }

        private void FgRotateButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            RotationFg(true);
        }
        private void FgRotateRButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            RotationFg(false);
        }
        private void OscRotateButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            RotationOsc(true);
        }
        private void OscRotateRButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            RotationOsc(false);
        }

    }
}
