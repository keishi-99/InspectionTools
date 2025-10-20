using InspectionTools.Common;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;
using static InspectionTools.MainMenu.SubMenuUserControl;

namespace InspectionTools.Product {
    /// <summary>
    /// EL0122FIUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL0122FIUserControl : UserControl, ISubMenuAware {

        private MainMenu.SubMenuUserControl? _subMenu;
        public void SetSubMenuControl(MainMenu.SubMenuUserControl? subMenu) {
            _subMenu = subMenu;
        }

        private readonly IntPtr _hWnd = IntPtr.Zero;

        private readonly InstClass _instDmm01;
        private readonly InstClass _instDmm02;
        private readonly InstClass _instFg;
        private readonly InstClass _instOsc;
        private readonly InstClass _instPs;

        public ObservableCollection<string> Dmm1List { get; } = [];
        public ObservableCollection<string> Dmm2List { get; } = [];
        public ObservableCollection<string> FgList { get; } = [];
        public ObservableCollection<string> OscList { get; } = [];
        public ObservableCollection<string> PsList { get; } = [];

        public EL0122FIUserControl() {
            InitializeComponent();
            _instDmm01 = new();
            _instDmm02 = new();
            _instFg = new();
            _instOsc = new();
            _instPs = new();
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

        private readonly List<Hotkey> _hotkeys = [];
        private HwndSource? _source;

        private volatile bool _isProcessing = false;

        private static readonly SemaphoreSlim s_semaphore = new(1, 1); // 最大1つの接続

        // 起動時
        private void LoadEvents() {
            InstListImport();
            FormatSet();
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
            UpdateComboBox(Dmm01ComboBox, Dmm1List, "デジタルマルチメータ", [1, 2], "[DMM-V]");
            UpdateComboBox(Dmm02ComboBox, Dmm2List, "デジタルマルチメータ", [1, 2], "[DMM-A]");
            UpdateComboBox(FgComboBox, FgList, "ファンクションジェネレータ", [2], "[FG]");
            UpdateComboBox(OscComboBox, OscList, "オシロスコープ", [2], "[OSC]");
            UpdateComboBox(PsComboBox, PsList, "パワーサプライ", [2], "[DCS]");
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
            GetVisaAddress(_instFg, FgComboBox);
            GetVisaAddress(_instOsc, OscComboBox);
            GetVisaAddress(_instPs, PsComboBox);
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
            _instPs.InstCommand = _instPs.SignalType switch {
                2 => "*RST;:VOLT 24;*OPC?",
                _ => string.Empty,
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {
            try {
                _subMenu?.SetButtonEnabled("ProductListButton", false);
                _subMenu?.SetButtonEnabled("InstListButton", false);

                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                CheckDmmId();
                FormatSet();

                var devices = new[] { _instDmm01, _instDmm02, _instFg, _instOsc, _instPs };
                var tasks = devices.Select(device => ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);

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
            _instPs.ResetProperties();

            _subMenu?.SetButtonEnabled("ProductListButton", true);
            _subMenu?.SetButtonEnabled("InstListButton", true);
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
        private static async Task<decimal> ReadDmm(InstClass instClass) {

            instClass.InstCommand = instClass.SignalType switch {
                1 => string.Empty,
                2 => "FETC?",
                _ => throw new ApplicationException(),
            };

            var result = await ConnectDeviceAsync(instClass);
            decimal.TryParse(result, NumberStyles.AllowExponent | NumberStyles.Float, CultureInfo.InvariantCulture, out var output);

            return output;
        }
        // OSC測定値取得
        private static async Task<decimal> ReadOsc(InstClass instClass, int oscMeas) {

            instClass.InstCommand = $"MEASU:MEAS{oscMeas}:VAL?";
            var result = await ConnectDeviceAsync(instClass);
            decimal.TryParse(result, NumberStyles.AllowExponent | NumberStyles.Float, CultureInfo.InvariantCulture, out var output);

            return output;
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

        // 全てのデータを処理するメソッド
        private async Task ProcessAllDataAsync(int delay) {
            if (_isProcessing) { return; }
            try {
                var sim = new InputSimulator();
                VisibleProgressImage(true);

                if (!string.IsNullOrEmpty(_instPs.VisaAddress)) {
                    await SwitchDcsAsync(_instPs, "ON");
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

                if (!string.IsNullOrEmpty(_instPs.VisaAddress)) {
                    await SwitchDcsAsync(_instPs, "OFF");
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
        private async void ActionHotkeyBracketR() {
            if (_isProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm02);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.00"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSC1測定値コピー
        private async void ActionHotkeySlash() {
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
        // OSC2測定値コピー
        private async void ActionHotkeyBackslash() {
            if (_isProcessing) { return; }

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
            await SwitchDcsAsync(_instPs, "ON");
        }
        private async void ActionHotkeyBracketL() {
            await SwitchDcsAsync(_instPs, "OFF");
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

        // HotkKeyの登録
        private void SetHotKey() {
            _hotkeys.Clear();

            _hotkeys.AddRange([
                new(ModNone, HotkeyComma, ActionHotkeyComma),
                new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
            ]);

            if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                ]);
            }
            if (!string.IsNullOrEmpty(_instPs.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                    new(ModNone, HotkeyBracketL, ActionHotkeyBracketL),
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
