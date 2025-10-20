using InspectionTools.Common;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;

namespace InspectionTools.Product {
    /// <summary>
    /// DFPDXUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class DFPDXUserControl : UserControl {

        private readonly IntPtr _hWnd = IntPtr.Zero;

        private readonly InstClass _instDmm01;
        private readonly InstClass _instDmm02;
        private readonly InstClass _instDmm03;
        private readonly InstClass _instFg;
        private readonly InstClass _instOsc;

        public ObservableCollection<string> Dmm1List { get; } = [];
        public ObservableCollection<string> Dmm2List { get; } = [];
        public ObservableCollection<string> Dmm3List { get; } = [];
        public ObservableCollection<string> FgList { get; } = [];
        public ObservableCollection<string> OscList { get; } = [];

        public DFPDXUserControl() {
            InitializeComponent();
            _instDmm01 = new();
            _instDmm02 = new();
            _instDmm03 = new();
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
            UpdateComboBox(Dmm03ComboBox, Dmm3List, "デジタルマルチメータ", [1, 2], "[DMM3]");
            UpdateComboBox(FgComboBox, FgList, "ファンクションジェネレータ", [2], "[FG]");
            UpdateComboBox(OscComboBox, OscList, "オシロスコープ", [2], "[OSC]");
        }
        private void UpdateComboBox(ComboBox comboBox, ObservableCollection<string> collection, string category, List<int> signalTypes, string name) {
            if (_dataTable == null) {
                return;
            }

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
            GetVisaAddress(_instDmm03, Dmm03ComboBox);
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
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicSwitchFg = new Dictionary<int, (string cmd, string text)>
            {
                { 0, (":FREQ 4000;:OUTPUT OFF;*OPC?", "OFF") },
                { 1, (":FREQ 4000;:OUTPUT ON;*OPC?", "4000") },
                { 2, (":OUTPUT OFF;*OPC?", "OFF") },
                { 3, (":OUTPUT OFF;:FREQ 5;:OUTPUT ON;*OPC?", "5") },
                { 4, (":OUTPUT OFF;:FREQ 100;:OUTPUT ON;*OPC?", "100") },
                { 5, (":OUTPUT OFF;:FREQ 4000;:OUTPUT ON;*OPC?", "4000")},
                { 6, (":OUTPUT OFF;:FREQ 6250;:OUTPUT ON;*OPC?", "6250") },
                { 7, (":OUTPUT OFF;:FREQ 4000;:OUTPUT ON;*OPC?", "4000") },
                { 8, (":OUTPUT OFF;:FREQ 100;:OUTPUT ON;*OPC?", "100") },
                { 9, (":OUTPUT OFF;:FREQ 60;:OUTPUT ON;*OPC?", "60") },
                { 10, (":OUTPUT OFF;:FREQ 72;:OUTPUT ON;*OPC?", "72") }
            };

            _dicSwitchOsc = new Dictionary<int, (string cmd, string text)>
            {
                { 0, (":HORIZONTAL:MAIN:SCALE 1.0E-3;:CH1:SCALE 1.0E-1;:TRIGGER:MAIN:LEVEL 0.0E0;*OPC?", "1[6]") },
                { 1, (":HORIZONTAL:MAIN:SCALE 5.0E-2;*OPC?", "2[7-1]") },
                { 2, (":HORIZONTAL:MAIN:SCALE 2.5E-3;*OPC?", "3[7-2]") },
                { 3, (":HORIZONTAL:MAIN:SCALE 1.0E-4;:CH1:SCALE 1.0E-1;*OPC?", "4[7-3]") },
                { 4, (":HORIZONTAL:MAIN:SCALE 1.0E-4;:CH1:SCALE 5.0E-1;*OPC?", "5[8]") },
                { 5, (":HORIZONTAL:MAIN:SCALE 2.5E-3;:CH1:SCALE 2.0E0;:TRIGGER:MAIN:LEVEL 0.0E0;*OPC?", "6[9]") },
                { 6, (":HORIZONTAL:MAIN:SCALE 2.5E-4;:CH1:SCALE 2.0E0;:TRIGGER:MAIN:LEVEL 1.2E0;*OPC?", "7[10]") },
                { 7, (":HORIZONTAL:MAIN:SCALE 5.0E-4;:CH1:SCALE 1.0E0;:TRIGGER:MAIN:LEVEL 1.2E0;*OPC?", "8[11]") }
            };
        }
        // 機器初期設定
        private void FormatSet() {
            _instDmm01.InstCommand = _instDmm01.SignalType switch {
                1 => "*RST,F5,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 0.2;*OPC?",
                _ => string.Empty,
            };
            _instDmm02.InstCommand = _instDmm02.SignalType switch {
                1 => "*RST,F5,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 0.2;*OPC?",
                _ => string.Empty,
            };
            _instDmm03.InstCommand = _instDmm03.SignalType switch {
                1 => "*RST,F1,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:VOLT:DC:RANG 20;*OPC?",
                _ => string.Empty,
            };
            _instFg.InstCommand = _instFg.SignalType switch {
                2 => "*RST;:OUTPUT OFF;:FREQ 4000;:VOLT 0.44VPP;*OPC?",
                _ => string.Empty,
            };
            _instOsc.InstCommand = _instOsc.SignalType switch {
                2 => """
                    *RST;:HEADER 0;
                    :ACQUIRE:MODE SAMPLE;NUMAVG 16;
                    :CH1:SCALE 1.0E-1;COUPLING DC;
                    :CURSOR:FUNCTION HBARS;SELECT:SOURCE CH1;:CURSOR:HBARS:POSITION1 2.0E-1;POSITION2 -2.0E-1;
                    :HORIZONTAL:MAIN:SCALE 1.0E-3;
                    :TRIGGER:MAIN:LEVEL 0.0E0;
                    :MEASUREMENT:MEAS1:TYPE MAXIMUM;SOURCE CH1;
                    :MEASUREMENT:MEAS2:TYPE MINIMUM;SOURCE CH1;
                    :MEASUREMENT:MEAS3:TYPE NWIDTH;SOURCE CH1;
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

                var devices = new[] { _instDmm01, _instDmm02, _instDmm03, _instFg, _instOsc };
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
                Dmm03ComboBox.IsEnabled = false;
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
        // DMMのIDチェック処理
        private void CheckDmmId() {
            var indices = new[] { _instDmm01.Index, _instDmm02.Index, _instDmm03.Index }
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
            _instDmm03.ResetProperties();
            _instFg.ResetProperties();
            _instOsc.ResetProperties();

            Dmm01ComboBox.IsEnabled = true;
            Dmm02ComboBox.IsEnabled = true;
            Dmm03ComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
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
        private void ActionHotkeyAtsign() {
            if (_isProcessing) { return; }
            RotationFg(true);
        }
        private void ActionHotkeyShiftAtsign() {
            if (_isProcessing) { return; }
            RotationFg(false);
        }
        // DMM01測定値コピー
        private async void ActionHotkeyPeriod() {
            if (_isProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm01);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.000"));
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
                sim.Keyboard.TextEntry((output * 1000).ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM03測定値コピー
        private async void ActionHotkeyBackslash() {
            if (_isProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm03);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(output.ToString("0.000"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSCローテーション
        private void ActionHotkeyBracketL() {
            if (_isProcessing) { return; }
            RotationOsc(true);
        }
        private void ActionHotkeyShiftracketL() {
            if (_isProcessing) { return; }
            RotationOsc(false);
        }
        // OSC meas1測定値コピー
        private async void ActionHotkeyBracketR() {
            if (_isProcessing) { return; }

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
            if (!string.IsNullOrEmpty(_instDmm03.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                ]);
            }
            if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                    new(ModShift, HotkeyAtsign, ActionHotkeyShiftAtsign)
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyBracketL, ActionHotkeyBracketL),
                    new(ModShift, HotkeyBracketL, ActionHotkeyShiftracketL),
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR)
                ]);
            }

            // 親ウィンドウを取得
            var parentWindow = Window.GetWindow(this);
            if (parentWindow != null) {
                var helper = new WindowInteropHelper(parentWindow).Handle;
                _source = HwndSource.FromHwnd(helper);
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
        private void TopMostCheckBox_Checked(object sender, RoutedEventArgs e) {
            var parentWindow = Window.GetWindow(this);
            parentWindow.Topmost = true;
        }
        private void TopMostCheckBox_Unchecked(object sender, RoutedEventArgs e) {
            var parentWindow = Window.GetWindow(this);
            parentWindow.Topmost = false;
        }

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
