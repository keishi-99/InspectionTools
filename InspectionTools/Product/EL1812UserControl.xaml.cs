using InspectionTools.Common;
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
    /// EL1812UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL1812UserControl : UserControl, ISubMenuAware {

        private MainMenu.SubMenuUserControl? _subMenu;
        public void SetSubMenuControl(MainMenu.SubMenuUserControl? subMenu) {
            _subMenu = subMenu;
        }

        private readonly DmmInstClass _instDmm = new();
        private readonly FgInstClass _instFg = new();
        private readonly OscInstClass _instOsc = new();

        public EL1812UserControl() {
            InitializeComponent();
        }
        private void AdjustWindowSizeToUserControl() {
            var parentWindow = Window.GetWindow(this);
            if (parentWindow != null) {
                parentWindow.SizeToContent = SizeToContent.WidthAndHeight;
            }
        }

        private const int TimeOut = 3;    //タイムアウトまでの時間(sec)

        private Dictionary<int, (string cmd, string text)> _dicSwitchFg = [];
        private Dictionary<int, (string cmd, string text)> _dicSwitchFgR = [];
        private Dictionary<int, (string cmd, string text)> _dicSwitchOsc = [];

        private volatile bool _isProcessing = false;

        private static readonly SemaphoreSlim s_semaphore = new(1, 1); // 最大1つの接続

        private int _rotateMenuNumber = 0;
        private List<string> _listRotateMenuTitle = [];

        // 起動時
        private void LoadEvents() {
            InstListImport();
            FormatSet();
            RegDictionary();
            RegMenuTitle();
            AdjustWindowSizeToUserControl();
        }
        private void InstListImport() {

            // デジタルマルチメータ、ファンクションジェネレータ、オシロスコープのコンボボックスを更新する
            UpdateComboBox(DmmComboBox, "デジタルマルチメータ", [1, 2], "[DMM]");
            UpdateComboBox(FgComboBox, "ファンクションジェネレータ", [3], "[FG]");
            UpdateComboBox(OscComboBox, "オシロスコープ", [2], "[OSC]");
        }
        private static void UpdateComboBox(ComboBox comboBox, string category, List<int> signalTypes, string name) {
            if (VisaAddressDataTable == null) {
                return;
            }

            var collection = new List<string> { name };

            foreach (var signalType in signalTypes) {
                var rows = VisaAddressDataTable.Select($"Category = '{category}' AND SignalType = {signalType}");
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
        private static void GetVisaAddress(InstClass instClass, ComboBox comboBox) {
            instClass.ResetProperties();

            instClass.Name = comboBox.Text;
            instClass.Index = comboBox.SelectedIndex;

            if (instClass.Index <= 0) { return; }

            var dRows = VisaAddressDataTable.Select($"Name = '{instClass.Name}'");
            instClass.Category = dRows[0]["Category"] as string ?? string.Empty;
            instClass.VisaAddress = dRows[0]["VisaAddress"] as string ?? string.Empty;
            instClass.SignalType = dRows[0]["SignalType"] != DBNull.Value ? Convert.ToInt32(dRows[0]["SignalType"]) : 0;
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            // FG
            _dicSwitchFg = new Dictionary<int, (string cmd, string text)>()
            {
                { 0, ("SIG 0;OMO 1;BTY 1;FRQ 2E+02;", "00" ) },
                { 1, ("SIG 1;MRK 79;TRG 1;", "01" ) },
                { 2, ("MRK 1;TRG 1;", "02" ) },
                { 3, ("MRK 839;TRG 1;", "03" ) },
                { 4, ("MRK 1;TRG 1;", "04" ) },
                { 5, ("MRK 78;TRG 1;", "05" ) },
                { 6, ("MRK 2;TRG 1;", "06" ) },
                { 7, ("FRQ 3E+3;MRK 1E+5;TRG 1;", "07" ) },
                { 8, ("MRK 10;TRG 1;", "08" ) },
                { 9, ("MRK 1002;TRG 1;", "09" ) },
                { 10, ("OMO 0;MRK 1;FRQ 10;", "10" ) },
                { 11, ("SIG 0;", "11" ) },
                { 12, ("SIG 1;", "12" ) },
                { 13, ("SIG 0;", "13" ) },
                { 14, ("SIG 1;", "14" ) },
            };
            _dicSwitchFgR = new Dictionary<int, (string cmd, string text)> {
                { 0, ("SIG 0;HIV 4.5;LOV 2.0;FRQ 2E+02;", "00" ) },
                { 1, ("HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 79;", "01" ) },
                { 2, ("HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 1;", "02" ) },
                { 3, ("HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 839;", "03" ) },
                { 4, ("HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 1;", "04" ) },
                { 5, ("HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 78;", "05" ) },
                { 6, ("HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 2;", "06" ) },
                { 7, ("FRQ 3E+3;MRK 1E+5;", "07" ) },
                { 8, ("FRQ 3E+3;MRK 10;", "08" ) },
                { 9, ("BTY 1;FRQ 3E+3;MRK 1002;", "09" ) },
                { 10, ("SIG 1;FRQ 10;", "10" ) },
                { 11, ("SIG 0;FRQ 10;", "11" )},
                { 12, ("SIG 1;FRQ 10;", "12" ) },
                { 13, ("SIG 0;FRQ 10;", "13" ) },
                { 14, ("SIG 1;MRK 1;BTY 0;FRQ 10;", "14" ) }
            };
            // OSC
            _dicSwitchOsc = new Dictionary<int, (string cmd, string text)> {
                { 0,
                    (
                    """
                    :HORIZONTAL:MAIN:SCALE 5.0E-5;
                    :HORIZONTAL:MAIN:POSITION 10.0E-5;
                    *OPC?
                    """,
                    "150u"
                    )
                },
                { 1,
                    (
                    """
                    :HORIZONTAL:MAIN:SCALE 25.0E-5;
                    *OPC?
                    """,
                    "1m"
                    )
                },
                { 2,
                    (
                    """
                    :HORIZONTAL:MAIN:SCALE 10.0E-3;
                    :HORIZONTAL:MAIN:POSITION 20.0E-3;
                    *OPC?
                    """,
                    "50m"
                    )
                }
            };
        }
        // 機器初期設定
        private void FormatSet() {
            // DMM
            _instDmm.InstCommand = _instDmm.SignalType switch {
                1 => "*RST,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;*OPC?",
                _ => string.Empty,
            };
            // FG
            _instFg.InstCommand = _instFg.SignalType switch {
                3 => "*RST;OMO 1;BES 0;BTY 1;FNC 3;TRS 1;TRE 0;BSS 1;BSV -100.0;FRQ 2E+02;HIV 4.5;LOV 2.0;MRK 79;SIG 0;",
                _ => string.Empty,
            };
            // OSC
            _instOsc.InstCommand = _instOsc.SignalType switch {
                2 => """
                    *RST;:HEADER 0;
                    :CH1:SCALE 5.0E0;POSITION -2.0E0;COUPLING DC;
                    :HORIZONTAL:MAIN:SCALE 5.0E-5;
                    :HORIZONTAL:MAIN:POSITION 10.0E-5;
                    :TRIGGER:MAIN:LEVEL 1.5E1;
                    :TRIGGER:MAIN:EDGE:SLOPE FALL;
                    :MEASUREMENT:MEAS1:TYPE NWIDTH;SOURCE CH1;
                    *OPC?
                    """,
                _ => string.Empty,
            };
        }
        // 検査項目名登録
        private void RegMenuTitle() {
            _listRotateMenuTitle = [
                "1.絶縁抵抗",
                "2.耐電圧",
                "3.検査準備",
                "4.外観構造",
                "5.発信器電源電圧",
                "6.バッチ動作",
                "7.バッチ量設定\r\nスケーラ-値積算\r\nパルス出力",
                "8.パルス未登録\r\nリーク警報\r\n過充填警報",
                "9.チャンネル選択",
                "10.外部入力",
                "11.パルス出力幅",
                "12.電源ON/OFF時\r\n設定値",
                "13.電圧、電流パルス(12V)",
                "14.バッチカウンタ初期化",
                "15.結果保存",
                "検査情報"
            ];
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

                if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                    FgRotateBackButton.IsEnabled = true;
                    FgRotateNextButton.IsEnabled = true;
                    FgOutputOnButton.IsEnabled = true;
                    FgOutputOffButton.IsEnabled = true;
                    FgTriggerButton.IsEnabled = true;
                    FgRotateRangeTextBox.Text = "00";
                }
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                    OscRotationButton.IsEnabled = true;
                    OscRotateRangeTextBox.Text = "150u";
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

            _instDmm.ResetProperties();
            _instFg.ResetProperties();
            _instOsc.ResetProperties();

            _subMenu?.SetButtonEnabled("ProductListButton", true);
            DmmComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyChekBox.IsChecked = false;
            FgRotateBackButton.IsEnabled = false;
            FgRotateNextButton.IsEnabled = false;
            FgOutputOnButton.IsEnabled = false;
            FgOutputOffButton.IsEnabled = false;
            FgTriggerButton.IsEnabled = false;
            OscRotationButton.IsEnabled = false;
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

            var dic = isNext ? _dicSwitchFg : _dicSwitchFgR;
            instClass.InstCommand = dic[_instFg.SettingNumber].cmd;

            if (instClass.InstCommand == string.Empty) { return; }

            await ConnectDeviceAsync(instClass);
        }
        // FG 出力ON/OFF
        private async void OutputFg(string cmd) {
            try {
                if (string.IsNullOrEmpty(_instFg.VisaAddress)) { return; }
                VisibleProgressImage(true);

                await OutputFgAsync(_instFg, cmd); ;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        private static async Task OutputFgAsync(InstClass instClass, string cmd) {
            instClass.InstCommand = $"{cmd};";
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

            instClass.InstCommand = _dicSwitchOsc[_instOsc.SettingNumber].cmd;

            if (instClass.InstCommand == string.Empty) { return; }

            await ConnectDeviceAsync(instClass);

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
        // 検査項目切り替え
        private void RotateMenu(int i) {

            var hWnd = FindWindow(null, "工程間検査");
            var hWnd2 = FindWindow(null, "検査情報画面");
            var hWnd3 = FindWindow(null, "確認");

            var windowToActivate = IntPtr.Zero;
            // ウィンドウが最小化されているか確認し、復元する
            if (hWnd3 != IntPtr.Zero) {
                windowToActivate = hWnd3;
            }
            else if (hWnd2 != IntPtr.Zero) {
                windowToActivate = hWnd2;
            }
            else if (hWnd != IntPtr.Zero) {
                windowToActivate = hWnd;
            }

            if (windowToActivate != IntPtr.Zero) {
                if (IsIconic(windowToActivate)) {
                    ShowWindow(windowToActivate, SwRestore);
                    System.Threading.Thread.Sleep(50);
                }
                SetForegroundWindow(windowToActivate);
            }
            else {
                return;
            }

            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);

            nint parentWindow;
            string itemName;

            switch (windowText.ToString()) {
                case "確認": {
                        return;
                    }
                case "検査情報画面": {
                        parentWindow = hWnd2;
                        itemName = "登録";
                        break;
                    }
                default: {
                        var n = _listRotateMenuTitle.Count;
                        _rotateMenuNumber = (n + ((_rotateMenuNumber + i) % n)) % n;
                        MenuRotateRangeTextBox.Text = (_rotateMenuNumber + 1).ToString("00");
                        var windowName = (_rotateMenuNumber != 15) ? "検査実施" : "検査情報";
                        parentWindow = FindWindowEx(hWnd, IntPtr.Zero, null, windowName);
                        itemName = _listRotateMenuTitle[_rotateMenuNumber];
                        break;
                    }
            }
            if (parentWindow == IntPtr.Zero) { return; }

            var targetWindow = FindWindowEx(parentWindow, IntPtr.Zero, null, itemName);
            if (targetWindow == IntPtr.Zero) { return; }

            _ = PostMessage(targetWindow, BmClick, 0, 0);

            nint hWnd4;
            List<nint> childHandles;
            switch (itemName) {
                case "5.発信器電源電圧":
                    hWnd4 = FindWindow(null, "発信器電源電圧");
                    //SetForegroundWindow(hWnd4);
                    childHandles = FindChildWindows(hWnd4);
                    _ = PostMessage(childHandles[5], EmSetSel, 0, -1);     // 入力欄全選択
                    _ = PostMessage(childHandles[5], WmClear, 0, -1);      // 入力欄消去
                    _ = PostMessage(childHandles[11], EmSetSel, 0, -1);     // 入力欄全選択
                    _ = PostMessage(childHandles[11], WmClear, 0, -1);      // 入力欄消去
                    _ = PostMessage(childHandles[11], WmLButtonDown, 0, 0);  // フォーカス
                    _ = PostMessage(childHandles[11], WmLButtonUp, 0, 0);  // フォーカス
                    break;
                case "10.外部入力":
                    hWnd4 = FindWindow(null, "外部入力");
                    //SetForegroundWindow(hWnd4);
                    childHandles = FindChildWindows(hWnd4);
                    _ = PostMessage(childHandles[11], WmLButtonDown, 0, 0);  // フォーカス
                    _ = PostMessage(hWnd4, WmLButtonUp, 0, 0);  // フォーカス
                    break;
                case "11.パルス出力幅":
                    hWnd4 = FindWindow(null, "パルス出力幅");
                    childHandles = FindChildWindows(hWnd4);
                    _ = PostMessage(childHandles[6], EmSetSel, 0, -1);     // 入力欄全選択
                    _ = PostMessage(childHandles[6], WmClear, 0, -1);      // 入力欄消去
                    _ = PostMessage(childHandles[12], EmSetSel, 0, -1);     // 入力欄全選択
                    _ = PostMessage(childHandles[12], WmClear, 0, -1);      // 入力欄消去
                    _ = PostMessage(childHandles[18], EmSetSel, 0, -1);     // 入力欄全選択
                    _ = PostMessage(childHandles[18], WmClear, 0, -1);      // 入力欄消去
                    _ = PostMessage(childHandles[23], WmLButtonDown, 0, 0);  // フォーカス
                    _ = PostMessage(hWnd4, WmLButtonUp, 0, 0);  // フォーカス
                    break;
                case "13.電圧、電流パルス(12V)":
                    break;
                case "14.バッチカウンタ初期化":
                    hWnd4 = FindWindow(null, "バッチカウンタ初期化");
                    childHandles = FindChildWindows(hWnd4);
                    _ = PostMessage(childHandles[15], WmLButtonDown, 0, 0);  // フォーカス
                    _ = PostMessage(hWnd4, WmLButtonUp, 0, 0);  // フォーカス
                    break;
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
        // Countロック
        private void CountLockToggle() { CountTextBox.IsEnabled = !CountLockCheckBox.IsChecked ?? false; }
        // Countインクリメント
        private void CountIncrement(int i) {
            if (!int.TryParse(CountTextBox.Text, out var intCount)) { return; }
            intCount += i;
            CountTextBox.Text = intCount.ToString();
        }
        // OSC切り替え
        private void ActionSwitchOscRange(bool isNext) {
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);
            if (windowText != "パルス出力幅") { return; }

            // 11.パルス出力幅でのみ有効
            if (_isProcessing) { return; }

            RotationOsc(isNext);
        }
        // OSC測定値コピー
        private async void ActionCopyOscValue() {
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);
            if (windowText != "パルス出力幅") { return; }

            // 11.パルス出力幅でのみ有効
            if (_isProcessing) { return; }

            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);

            var output = await ReadOsc(_instOsc, 1);
            (var value, var format) = _instOsc.SettingNumber switch {
                0 => (output * 1000000, "0.0"),
                1 or 2 => (output * 1000, "0.00"),
                _ => (output, "0.00"),
            };
            sim.Keyboard.TextEntry(value.ToString(format));
        }
        // DMM測定値コピー
        private async void ActionCopyDmmValue() {
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);
            if (windowText != "発信器電源電圧") { return; }

            // 5.発信器電源電圧でのみ有効
            if (_isProcessing) { return; }

            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);

            var output = await ReadDmm(_instDmm);
            sim.Keyboard.TextEntry(output.ToString("0.0"));
            Thread.Sleep(200);
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
        }
        // Serial貼り付け
        private void ActionPasteSerial() {
            if (string.IsNullOrEmpty(SerialTextBox.Text)) { return; }

            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);
            if (windowText != "検査情報画面") { return; }

            // 検査情報画面でのみ有効
            var childHandles = FindChildWindows(foregroundWindow);
            _ = PostMessage(childHandles[15], WmLButtonDown, 0, 0);  // フォーカス
            _ = PostMessage(childHandles[15], WmLButtonUp, 0, 0);  // フォーカス

            var sim = new InputSimulator();
            sim.Keyboard.TextEntry(SerialTextBox.Text);
            SerialIncrement(1);
        }
        // Count貼り付け
        private void ActionPasteCount() {
            if (string.IsNullOrEmpty(CountTextBox.Text)) { return; }

            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);
            if (windowText != "検査情報画面") { return; }

            // 検査情報画面でのみ有効
            var childHandles = FindChildWindows(foregroundWindow);
            _ = PostMessage(childHandles[6], WmLButtonDown, 0, 0);  // フォーカス
            _ = PostMessage(childHandles[6], WmLButtonUp, 0, 0);  // フォーカス

            var sim = new InputSimulator();
            sim.Keyboard.TextEntry(CountTextBox.Text);
            CountIncrement(1);
        }
        // FG切り替え
        private void ActionSwitchFg(bool isNext) {
            if (_isProcessing) { return; }
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            // FGを使用する項目でのみ有効
            var windowText = GetWindowText(foregroundWindow);
            if (windowText is
                not "バッチ動作" and
                not "スケーラ値" and
                not "警報" and
                not "パルス出力幅"
                ) { return; }
            RotationFg(isNext);
        }
        // OSC+FG切り替え
        private void ActionSwitchFgOsc(bool isNext) {
            if (_isProcessing) { return; }
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            // パルス出力幅 でのみ有効
            var windowText = GetWindowText(foregroundWindow);
            if (windowText is
                not "パルス出力幅"
                ) { return; }
            RotationFg(isNext);
            RotationOsc(isNext);
        }

        // メニュー切り替え
        private void ActionHotkeyBracketL() { RotateMenu(1); }
        private void ActionHotkeyNum7() { RotateMenu(1); }
        // Tabキー
        private static void ActionHotkeyAtsign() {
            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
        }
        private static void ActionHotkeyNum8() {
            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
        }
        // OSC切り替え
        private void ActionHotkeyBracketR() { ActionSwitchOscRange(true); }
        private void ActionHotkeyNum1() { ActionSwitchOscRange(true); }
        // OSC測定値コピー
        private void ActionHotkeyColon() { ActionCopyOscValue(); }
        private void ActionHotkeyNum6() { ActionCopyOscValue(); }
        // DMM測定値コピー
        private void ActionHotkeySemiColon() { ActionCopyDmmValue(); }
        private void ActionHotkeyNum9() { ActionCopyDmmValue(); }
        // Serial貼り付け
        private void ActionHotkeyPeriod() { ActionPasteSerial(); }
        private void ActionHotkeyNumSubtract() { ActionPasteSerial(); }
        // Count貼り付け
        private void ActionHotkeyComma() { ActionPasteCount(); }
        private void ActionHotkeyNumAdd() { ActionPasteCount(); }
        // FG切り替え
        private void ActionHotkeyBackslash() { ActionSwitchFg(true); }
        private void ActionHotkeyNum5() { ActionSwitchFg(true); }
        private void ActionHotkeyShiftBackslash() { ActionSwitchFg(false); }
        private void ActionHotkeyShiftNum5() { ActionSwitchFg(false); }
        private void ActionHotkeyNum2() {
            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            ActionSwitchFg(true);
        }
        private void ActionHotkeyNum3() {
            if (_isProcessing) { return; }

            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);

            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            // FGを使用する項目でのみ有効
            var windowText = GetWindowText(foregroundWindow);
            if (windowText is
                not "バッチ動作" and
                not "スケーラ値" and
                not "警報" and
                not "パルス出力幅"
                ) { return; }
            ActionSwitchFgOsc(true);
        }

        // HotkKeyの登録
        private void SetHotKey() {
            MainWindow.HotkeysList.Clear();

            MainWindow.HotkeysList.AddRange([
                new(ModNone, HotkeyBracketL, ActionHotkeyBracketL),
                new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                new(ModNone, HotkeyComma, ActionHotkeyComma),
                new(ModNone, HotkeyNum7, ActionHotkeyNum7),
                new(ModNone, HotkeyNum8, ActionHotkeyNum8),
                new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
                new(ModNone, HotkeyNumSubtract, ActionHotkeyNumSubtract),
            ]);
            if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    new(ModShift, HotkeyBackslash, ActionHotkeyShiftBackslash),
                    new(ModNone, HotkeyNum5, ActionHotkeyNum5),
                    new(ModShift, HotkeyNum5, ActionHotkeyShiftNum5),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeySemiColon, ActionHotkeySemiColon),
                    new(ModNone, HotkeyNum9,ActionHotkeyNum9),
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                    new(ModNone, HotkeyNum1, ActionHotkeyNum1),
                    new(ModNone, HotkeyNum6, ActionHotkeyNum6),
                ]);
            }
            if (!string.IsNullOrEmpty(_instFg.VisaAddress) && !string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyNum2, ActionHotkeyNum2),
                    new(ModNone, HotkeyNum3, ActionHotkeyNum3),
                ]);
            }

            MainWindow.Source = HwndSource.FromHwnd(MainWindow.HWnd);
            MainWindow.Source.AddHook(HwndHook);

            // ホットキーを登録
            foreach (var hotkey in MainWindow.HotkeysList) {
                RegisterHotKey(MainWindow.HWnd, hotkey.Id, hotkey.Modifier, (uint)hotkey.VirtualKey);
            }
        }
        private static void ClearHotKey() {
            foreach (var hotkey in MainWindow.HotkeysList) {
                UnregisterHotKey(MainWindow.HWnd, hotkey.Id);
            }
            MainWindow.HotkeysList.Clear();
        }
        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
            if (msg == WmHotKey) {
                int id = wParam.ToInt32();

                var hotkey = MainWindow.HotkeysList.FirstOrDefault(h => h.Id == id);
                hotkey?.Action.Invoke(); // ホットキーに設定されたアクションを実行
                handled = true;
            }
            return IntPtr.Zero;
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

        // イベントハンドラ
        private void UserControl_Loaded(object sender, RoutedEventArgs e) { LoadEvents(); }
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }

        private void OscRotationButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            RotationOsc(true);
        }
        private void FgRotateBackButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            RotationFg(false);
        }
        private void FgRotateNextButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            RotationFg(true);
        }
        private void FgOutputOnButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            OutputFg("SIG 1");
        }
        private void FgOutputOffButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            OutputFg("SIG 0");
        }
        private void FgTriggerButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            OutputFg("TRG 1");
        }

        private void MenuRotateBackButton_Click(object sender, RoutedEventArgs e) { RotateMenu(-1); }
        private void MenuRotateNextButton_Click(object sender, RoutedEventArgs e) { RotateMenu(1); }

        private void SerialBack_Click(object sender, RoutedEventArgs e) { SerialIncrement(-1); }
        private void SerialNext_Click(object sender, RoutedEventArgs e) { SerialIncrement(1); }
        private void SerialLockCheckBox_Checked(object sender, RoutedEventArgs e) { SerialLockToggle(); }
        private void SerialLockCheckBox_Unchecked(object sender, RoutedEventArgs e) { SerialLockToggle(); }

        private void CountBack_Click(object sender, RoutedEventArgs e) { CountIncrement(-1); }
        private void CountNext_Click(object sender, RoutedEventArgs e) { CountIncrement(1); }
        private void CountLockCheckBox_Checked(object sender, RoutedEventArgs e) { CountLockToggle(); }
        private void CountLockCheckBox_Unchecked(object sender, RoutedEventArgs e) { CountLockToggle(); }


    }
}
