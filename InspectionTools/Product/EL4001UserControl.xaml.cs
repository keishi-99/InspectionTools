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
    /// EL4001UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL4001UserControl : UserControl, ISubMenuAware {

        private MainMenu.SubMenuUserControl? _subMenu;
        public void SetSubMenuControl(MainMenu.SubMenuUserControl? subMenu) {
            _subMenu = subMenu;
        }

        private IntPtr _hWnd = IntPtr.Zero;

        private readonly DcsInstClass _instDcs;
        private readonly DmmInstClass _instDmm01;
        private readonly DmmInstClass _instDmm02;
        private readonly OscInstClass _instOsc;

        public EL4001UserControl() {
            InitializeComponent();
            _instDcs = new();
            _instDmm01 = new();
            _instDmm02 = new();
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

        private Dictionary<int, (string cmd2, string cmd3, string text)> _dicSwitchDcs = [];
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
            UpdateComboBox(DcsComboBox, "電流電圧発生器", [2, 3], "[DCS]");
            UpdateComboBox(Dmm01ComboBox, "デジタルマルチメータ", [1, 2], "[DMM1]");
            UpdateComboBox(Dmm02ComboBox, "デジタルマルチメータ", [1, 2], "[DMM2]");
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
            GetVisaAddress(_instDcs, DcsComboBox);
            GetVisaAddress(_instDmm01, Dmm01ComboBox);
            GetVisaAddress(_instDmm02, Dmm02ComboBox);
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
                _subMenu?.SetButtonEnabled("ProductListButton", false);

                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                CheckDmmId();
                FormatSet();

                InstClass[] devices = [_instDcs, _instDmm01, _instDmm02, _instOsc];
                var tasks = devices.Select(device => ConnectDeviceAsync(device));
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
            _instDmm01.ResetProperties();
            _instDmm02.ResetProperties();
            _instOsc.ResetProperties();

            _subMenu?.SetButtonEnabled("ProductListButton", true);
            Dmm01ComboBox.IsEnabled = true;
            Dmm02ComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            DcsComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyChekBox.IsChecked = false;
            OscRangeLabel.Text = string.Empty;
            DcsNumberLabel.Text = string.Empty;
            DcsRangeLabel.Text = string.Empty;
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
        // DCSローテーション
        private async void RotationDcs(InstClass instClass, bool isNext) {
            try {
                VisibleProgressImage(true);

                var maxSettingNumber = _dicSwitchDcs.Keys.Count;
                instClass.SettingNumber = (instClass.SettingNumber + (isNext ? 1 : -1) + maxSettingNumber) % maxSettingNumber;
                var settingNumber = instClass.SettingNumber;

                instClass.InstCommand = instClass.SignalType switch {
                    2 => _dicSwitchDcs[settingNumber].cmd2,
                    3 => _dicSwitchDcs[settingNumber].cmd3,
                    _ => throw new ApplicationException(),
                };

                await ConnectDeviceAsync(instClass);
                DcsNumberLabel.Text = settingNumber.ToString("00");
                DcsRangeLabel.Text = _dicSwitchDcs[settingNumber].text;

                VisibleProgressImage(false);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // OSCローテーション
        private async void RotationOsc(InstClass instClass) {
            try {
                VisibleProgressImage(true);

                (instClass.InstCommand, var rangeText) = instClass.SettingNumber switch {
                    0 => (":HORIZONTAL:MAIN:SCALE 5.0E-4;POSITION 0.0;*OPC?", "500u"),
                    1 => (":HORIZONTAL:MAIN:SCALE 5.0E-2;POSITION 1.0E-1;*OPC?", "50ms"),
                    _ => throw new ApplicationException(),
                };

                await ConnectDeviceAsync(instClass);

                OscRangeLabel.Text = rangeText;

                // 設定番号を反転
                instClass.SettingNumber = (instClass.SettingNumber == 0) ? 1 : 0;

                VisibleProgressImage(false);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // DMM01測定値コピー
        private async void ActionHotkeyColon() {
            if (_isProcessing) { return; }

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
            if (_isProcessing) { return; }

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
            if (_isProcessing) { return; }

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
            if (_isProcessing) { return; }

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
            if (_isProcessing) { return; }
            RotationOsc(_instOsc);
        }
        // OSC meas1測定値コピー
        private async void ActionHotkeySlash() {
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
        // OSC meas2測定値コピー
        private async void ActionHotkeyBackslash() {
            if (_isProcessing) { return; }

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
            if (_isProcessing) { return; }
            RotationDcs(_instDcs, true);
        }
        private void ActionHotkeyShiftAtsign() {
            if (_isProcessing) { return; }
            RotationDcs(_instDcs, false);
        }
        private void ActionHotkeyNumDivide() {
            if (_isProcessing) { return; }
            RotationDcs(_instDcs, true);
        }

        // HotkKeyの登録
        private void SetHotKey() {
            _hotkeys.Clear();
            if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                    new(ModNone, HotkeyNumMultiply, ActionHotkeyNumMultiply),
                    ]);
            }
            if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                        new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
                    ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                        new(ModNone, HotkeySlash, ActionHotkeySlash),
                        new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    ]);
            }
            if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                        new(ModShift, HotkeyAtsign, ActionHotkeyShiftAtsign),
                        new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
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


    }
}
