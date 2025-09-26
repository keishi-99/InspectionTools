using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Threading;
using Tools.Common;
using Tools.Common.InstList;
using WindowsInput;
using static Tools.Common.Win32Wrapper;

namespace PAF5 {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        private IntPtr _hWnd = IntPtr.Zero;

        private readonly InstClass _instDmm01;
        private readonly InstClass _instDmm02;
        private readonly InstClass _instFg;
        private readonly InstClass _instOsc;

        public ObservableCollection<string> Dmm1List { get; } = [];
        public ObservableCollection<string> Dmm2List { get; } = [];
        public ObservableCollection<string> FgList { get; } = [];
        public ObservableCollection<string> OscList { get; } = [];

        public MainWindow() {
            InitializeComponent();
            _instDmm01 = new();
            _instDmm02 = new();
            _instFg = new();
            _instOsc = new();
            LoadEvents();
            // Window が完全に作られたあとにハンドルを取得
            Loaded += (s, e) => { _hWnd = new WindowInteropHelper(this).Handle; };

            _timer = new DispatcherTimer {
                Interval = TimeSpan.FromSeconds(1)
            };
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        public class InstClass {
            internal USBDeviceManager UsbDev { get; set; } = new();
            public string Category { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public string VisaAddress { get; set; } = string.Empty;
            public int SignalType { get; set; } = 0;
            public int Index { get; set; } = 0;
            public string InstCommand { get; set; } = string.Empty;
            public int SettingNumber { get; set; } = 0;

            public void ResetProperties() {
                UsbDev = new();
                Name = string.Empty;
                Category = string.Empty;
                VisaAddress = string.Empty;
                SignalType = 0;
                Index = 0;
                InstCommand = string.Empty;
                SettingNumber = 0;
            }
            public void Dispose() {
                // UsbDevの解放処理
                UsbDev?.Dispose();
            }
        }

        private const int TimeOut = 3;    //タイムアウトまでの時間(sec)

        internal DataTable _dataTable = new();

        private Dictionary<int, (string cmd, string text)> _dicSwitchFg = [];
        private Dictionary<int, (string cmd, string text)> _dicSwitchOsc = [];
        private readonly List<Hotkey> _hotkeys = [];
        private HwndSource? _source;

        private volatile bool _isProcessing = false;

        private static readonly SemaphoreSlim s_semaphore = new(1, 1); // 最大1つの接続

        // タイマー
        private readonly DispatcherTimer _timer;

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
            UpdateComboBox(Dmm01ComboBox, Dmm1List, "デジタルマルチメータ", [1, 2], "[DMM1]");
            UpdateComboBox(Dmm02ComboBox, Dmm2List, "デジタルマルチメータ", [1, 2], "[DMM2]");
            UpdateComboBox(FgComboBox, FgList, "ファンクションジェネレータ", [2], "[FG]");
            UpdateComboBox(OscComboBox, OscList, "オシロスコープ", [2], "[OSC]");
        }
        private void UpdateComboBox(ComboBox comboBox, ObservableCollection<string> collection, string category, List<int> signalTypes, string name) {
            if (_dataTable == null) return;
            collection.Clear();
            collection.Add(name);

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
            GetVisaAddress(_instDmm01, Dmm01ComboBox);
            GetVisaAddress(_instDmm02, Dmm02ComboBox);
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
        // 機器リスト表示
        private void ShowInstList() {
            InstListWindow frm1 = new(_dataTable);
            frm1.ShowDialog();
            InstListImport();
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicSwitchFg = new Dictionary<int, (string cmd, string text)> {
                { 0,  (":FREQ 20;:OUTPUT OFF;*OPC?", "OFF")  },
                { 1,  (":FREQ 20;:OUTPUT ON;*OPC?", "20")  },
                { 2, (":OUTPUT OFF;:FREQ 200;:OUTPUT ON;*OPC?", "200") },
                { 3, (":OUTPUT OFF;:FREQ 2000;:OUTPUT ON;*OPC?", "2000")  },
                { 4,  (":OUTPUT OFF;:FREQ 6250;:OUTPUT ON;*OPC?", "6250")  },
                { 5,  (":OUTPUT OFF;:FREQ 5000;:OUTPUT ON;*OPC?", "5000")  },
                { 6,  (":OUTPUT OFF;:FREQ 1000;:OUTPUT ON;*OPC?", "1000")  },
                { 7,  (":OUTPUT OFF;:FREQ 400;:OUTPUT ON;*OPC?", "400") } ,
                { 8,  (":OUTPUT OFF;:FREQ 0.51;:OUTPUT ON;*OPC?", "0.51[min][F]") },
                { 9,  (":OUTPUT OFF;:FREQ 1.0;:OUTPUT ON;*OPC?", "1[min][F]")  },
                { 10,  (":OUTPUT OFF;:FREQ 3.5;:OUTPUT ON;*OPC?", "3.5[min][7]")  },
                { 11, (":OUTPUT OFF;:FREQ 5.7;:OUTPUT ON;*OPC?", "5.7[min][7]")  },
                { 12, (":OUTPUT OFF;:FREQ 33;:OUTPUT ON;*OPC?", "33[max][7]") } ,
                { 13,  (":OUTPUT OFF;:FREQ 46;:OUTPUT ON;*OPC?", "46[max][7]")  },
                { 14,  (":OUTPUT OFF;:FREQ 5.7;:OUTPUT ON;*OPC?", "5.7[max][F]")  },
                { 15,  (":OUTPUT OFF;:FREQ 8.0;:OUTPUT ON;*OPC?", "8[max][F]")  }
            };

            _dicSwitchOsc = new Dictionary<int, (string cmd, string text)> {
                { 0,  (":HORIZONTAL:MAIN:SCALE 1.0E-3;*OPC?", "1")  },
                { 1,  (":HORIZONTAL:MAIN:SCALE 1.0E-2;*OPC?", "2")  },
                { 2,  (":HORIZONTAL:MAIN:SCALE 1.0E-3;*OPC?", "3")  },
                { 3,  (":HORIZONTAL:MAIN:SCALE 1.0E-4;*OPC?", "4")  }
            };
        }
        // 機器初期設定
        private void FormatSet() {
            _instDmm01.InstCommand = _instDmm01.SignalType switch {
                1 => "*RST,R7,*OPC?",
                2 => "*RST;:INIT:CONT 1;:VOLT:DC:RANG 200;*OPC?",
                _ => string.Empty
            };
            _instDmm02.InstCommand = _instDmm02.SignalType switch {
                1 => "*RST,F5,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?",
                _ => string.Empty,
            };
            _instFg.InstCommand = _instFg.SignalType switch {
                2 => "*RST;:FREQ 20;:VOLT 0.44VPP;*OPC?",
                _ => string.Empty,
            };
            _instOsc.InstCommand = _instOsc.SignalType switch {
                2 =>
                    """
                    *RST;:HEADER 0;
                    :ACQUIRE:MODE AVERAGE;
                    :CH1:SCALE 1.0E-1;COUPLING DC;
                    :CURSOR:FUNCTION HBARS;SELECT:SOURCE CH1;:CURSOR:HBARS:POSITION1 2.0E-1;POSITION2 -2.0E-1;
                    :HORIZONTAL:MAIN:SCALE 1.0E-3;
                    :MEASUREMENT:MEAS1:TYPE PK2PK;SOURCE CH1;
                    *OPC?
                    """,
                _ => string.Empty,
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {
            try {
                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                CheckDmmId();
                FormatSet();

                var devices = new[] { _instDmm01, _instDmm02, _instFg, _instOsc };
                var tasks = devices.Select(device => ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);

                if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                    FgRotateRangeTextBox.Text = "OFF";
                    FgRotateButton.IsEnabled = true;
                }
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                    OscRotateRangeTextBox.Text = "1";
                    OscRotateButton.IsEnabled = true;
                }

                Dmm01ComboBox.IsEnabled = false;
                Dmm02ComboBox.IsEnabled = false;
                FgComboBox.IsEnabled = false;
                OscComboBox.IsEnabled = false;
                ConnectButton.IsEnabled = false;
                ReleaseButton.IsEnabled = true;
                InstListButton.IsEnabled = false;

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

            _instDmm01.ResetProperties();
            _instDmm02.ResetProperties();
            _instFg.ResetProperties();
            _instOsc.ResetProperties();

            Dmm01ComboBox.IsEnabled = true;
            Dmm02ComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            InstListButton.IsEnabled = true;
            HotKeyChekBox.IsChecked = false;
            FgRotateButton.IsEnabled = false;
            OscRotateButton.IsEnabled = false;
            FgRotateRangeTextBox.Text = string.Empty;
            OscRotateRangeTextBox.Text = string.Empty;
        }

        // DMM測定値取得
        private async Task<decimal> ReadDmm(InstClass instClass) {
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
                Activate();
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
                Activate();
            } finally {
                VisibleProgressImage(false);
            }
        }
        private async Task RotationOscAsync(InstClass instClass, bool isNext) {

            var oscMaxSettingNumber = _dicSwitchOsc.Count;
            _instOsc.SettingNumber = (_instOsc.SettingNumber + (isNext ? 1 : -1) + oscMaxSettingNumber) % oscMaxSettingNumber;

            instClass.InstCommand = _dicSwitchOsc[_instOsc.SettingNumber].cmd;

            if (instClass.InstCommand == string.Empty) { return; }

            await ConnectDeviceAsync(instClass);

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
        // DMM01測定値コピー
        private async void ActionHotkeyPeriod() {
            if (_isProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm01);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(output.ToString("0.00"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM02測定値コピー
        private async void ActionHotkeySlash() {
            if (_isProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm02);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000000).ToString("0.00"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSCローテーション
        private void ActionHotkeyColon() {
            if (_isProcessing) { return; }
            RotationOsc(true);
        }
        private void ActionHotkeyNumDivide() {
            if (_isProcessing) { return; }
            RotationOsc(true);
        }
        // OSC meas1測定値コピー
        private async void ActionHotkeyBackslash() {
            if (_isProcessing) { return; }

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

        // HotkKeyの登録
        private void SetHotKey() {
            _hotkeys.Clear();
            if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                _hotkeys.AddRange([
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
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
                ]);
            }


            var helper = new WindowInteropHelper(this);
            _source = HwndSource.FromHwnd(helper.Handle);
            _source.AddHook(HwndHook);

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
        private void InstListButton_Click(object sender, RoutedEventArgs e) { ShowInstList(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }
        private void TopMostCheckBox_Checked(object sender, RoutedEventArgs e) { Topmost = true; }
        private void TopMostCheckBox_Unchecked(object sender, RoutedEventArgs e) { Topmost = false; }
        private void Timer_Tick(object? sender, EventArgs e) { Time.Text = DateTime.Now.ToString("HH:mm:ss"); }

        private void FgRotateButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            RotationFg(true);
        }
        private void OscRotateButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            RotationOsc(true);
        }
    }
}