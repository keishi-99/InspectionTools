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
    /// PAF5ampUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class PAF5ampUserControl : UserControl {

        private readonly IntPtr _hWnd = IntPtr.Zero;

        private readonly InstClass _instDcs;
        private readonly InstClass _instDmm;
        private readonly InstClass _instFg;
        private readonly InstClass _instOsc;

        public ObservableCollection<string> DcsList { get; } = [];
        public ObservableCollection<string> DmmList { get; } = [];
        public ObservableCollection<string> FgList { get; } = [];
        public ObservableCollection<string> OscList { get; } = [];

        public PAF5ampUserControl() {
            InitializeComponent();
            _instDcs = new();
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

        private Dictionary<int, string> _dicSwitchFg = [];
        private Dictionary<int, string> _dicSwitchOsc = [];

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
            UpdateComboBox(DcsComboBox, DcsList, "パワーサプライ", [2], "[DCS]");
            UpdateComboBox(DmmComboBox, DmmList, "デジタルマルチメータ", [1, 2], "[DMM]");
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
            GetVisaAddress(_instDcs, DcsComboBox);
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
                1 => "*RST,F5,R7,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?",
                _ => string.Empty,
            };
            _instFg.InstCommand = _instFg.SignalType switch {
                2 => "*RST;:FREQ 1.7E0;*OPC?",
                _ => string.Empty,
            };
            _instOsc.InstCommand = _instOsc.SignalType switch {
                2 =>
                    """
                    *RST;
                    :HEADER 0;
                    :CH1:SCALE 1.0E-1;
                    :HORIZONTAL:MAIN:SCALE 5.0E-2;
                    :MEASUREMENT:MEAS1:TYPE PK2PK;SOURCE CH1;
                    *OPC?
                    """,
                _ => string.Empty,
            };
            _instDcs.InstCommand = _instDcs.SignalType switch {
                2 => "*RST;:VOLT 1.5;*OPC?",
                _ => string.Empty,
            };
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicSwitchFg = new Dictionary<int, string> {
                { 0, ":FREQ 1.7E0;*OPC?" },
                { 1, ":FREQ 1.0E1;*OPC?" },
                { 2, ":FREQ 4.0E3;*OPC?" },
            };

            _dicSwitchOsc = new Dictionary<int, string> {
                { 0, ":HORIZONTAL:MAIN:SCALE 5.0E-2;*OPC?" },
                { 1, ":HORIZONTAL:MAIN:SCALE 2.5E-2;*OPC?" },
                { 2, ":HORIZONTAL:MAIN:SCALE 5.0E-5;*OPC?" },
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {
            try {
                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                FormatSet();

                var devices = new[] { _instDcs, _instDmm, _instFg, _instOsc };
                var tasks = devices.Select(device => ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);

                DcsComboBox.IsEnabled = false;
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

            _instDcs.ResetProperties();
            _instDmm.ResetProperties();
            _instFg.ResetProperties();
            _instOsc.ResetProperties();

            DcsComboBox.IsEnabled = true;
            DmmComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyChekBox.IsChecked = false;
        }

        // 電源のON-OFF
        private async Task SwitchDcsAsync(InstClass instClass, string cmd) {
            try {
                instClass.InstCommand = $":OUTPUT {cmd};*OPC?";
                await ConnectDeviceAsync(instClass);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
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

            instClass.InstCommand = _dicSwitchFg[_instFg.SettingNumber];

            if (instClass.InstCommand == string.Empty) { return; }

            await ConnectDeviceAsync(instClass);
        }
        // OSC切り替え
        private async void RotationOsc(bool isNext) {
            try {
                if (string.IsNullOrEmpty(_instOsc.VisaAddress)) { return; }
                VisibleProgressImage(true);

                await RotationOscAsync(_instOsc, isNext);

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

            instClass.InstCommand = _dicSwitchOsc[_instOsc.SettingNumber];

            if (instClass.InstCommand == string.Empty) { return; }

            await ConnectDeviceAsync(instClass);
        }

        // 電源ON-OFF
        private async void ActionHotkeyAtsign() {
            await SwitchDcsAsync(_instDcs, "ON");
        }
        private async void ActionHotkeyBracketL() {
            await SwitchDcsAsync(_instDcs, "OFF");
        }
        // DMM測定値コピー
        private async void ActionHotkeySlash() {
            if (_isProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000000).ToString("0"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.DOWN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // FGローテーション
        private void ActionHotkeyColon() {
            if (_isProcessing) { return; }
            RotationFg(true);
        }
        // OSCローテーション
        private void ActionHotkeyBracketR() {
            if (_isProcessing) { return; }
            RotationOsc(true);
        }

        // HotkKeyの登録
        private void SetHotKey() {
            _hotkeys.Clear();

            if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                    new(ModNone, HotkeyBracketL, ActionHotkeyBracketL),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash)
                ]);
            }
            if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
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


    }
}
