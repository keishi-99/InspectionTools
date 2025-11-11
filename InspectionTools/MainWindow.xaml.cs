using InspectionTools.Common;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Interop;
using static InspectionTools.Common.Win32Wrapper;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        private bool _isHelpVisible = false;
        private string _pageName = string.Empty;

        public static IntPtr HWnd { get; set; }
        public static HwndSource? Source { get; set; }
        public static List<Hotkey> HotkeysList { get; set; } = [];

        public static DataTable VisaAddressDataTable { get; set; } = new();

        public static bool IsProcessing { get; set; } = false;

        private const int TimeOut = 3;    //タイムアウトまでの時間(sec)

        public static readonly SemaphoreSlim s_semaphore = new(1, 1); // 最大1つの接続

        private MainMenu.SubMenuUserControl? _subMenu;

        public MainWindow() {
            InitializeComponent();
            Common.HelpManager.LoadHelpFile("help.json");
            ShowMainMenu();
            // Window が完全に作られたあとにハンドルを取得
            Loaded += (s, e) => { HWnd = new WindowInteropHelper(this).Handle; };
        }

        private void ShowMainMenu() {

            _subMenu = new MainMenu.SubMenuUserControl();
            _subMenu.BackToMainRequested += (_, __) => ShowMainMenu();
            _subMenu.HelpButtonClicked += OnHelpButtonClicked;
            _subMenu.SetButtonEnabled("ProductListButton", false);
            _subMenu.SetButtonEnabled("InstListButton", true);
            SubMenuContentArea.Content = _subMenu;

            var mainMenu = new MainMenu.MainMenuUserControl();
            mainMenu.PageSelected += OnPageSelected;
            MainMenuContentArea.Content = mainMenu;

            this.Title = "Menu";
            _pageName = "MainMenu";
            if (_isHelpVisible) {
                var (keys, descriptions) = Common.HelpManager.GetHelpData(_pageName);
                HelpTextBlock1.Text = string.Join(Environment.NewLine, keys);
                HelpTextBlock2.Text = string.Join(Environment.NewLine, descriptions);
            }
            HotKeyHelpScrollViewer.Height = mainMenu.Height + (_subMenu?.Height ?? 0);
        }

        private void OnPageSelected(string pageName) {

            _subMenu?.SetButtonEnabled("ProductListButton", true);
            _subMenu?.SetButtonEnabled("InstListButton", false);

            UserControl? page = pageName switch {
                "EL0122FI" => new Product.EL0122FIUserControl(),
                "EL0122" => new Product.EL0122UserControl(),
                "EL0137" => new Product.EL0137UserControl(),
                "EL1812" => new Product.EL1812UserControl(),
                "EL3801" => new Product.EL3801UserControl(),
                "EL4001" => new Product.EL4001UserControl(),
                "EL9100" => new Product.EL9100UserControl(),
                "EL9240" => new Product.EL9240UserControl(),

                "PA14" => new Product.PA14UserControl(),
                "PA25" => new Product.PA25UserControl(),
                "PAF5amp" => new Product.PAF5ampUserControl(),
                "PAF5" => new Product.PAF5UserControl(),

                "DFPDX" => new Product.DFPDXUserControl(),
                "MassFlow" => new Product.MassFlowUserControl(),
                _ => null
            };

            if (page is not null) {

                this.Title = pageName;
                _pageName = pageName;
                MainMenuContentArea.Content = page;

                if (page is MainMenu.SubMenuUserControl.ISubMenuAware s) {
                    s.SetSubMenuControl(_subMenu);
                }

                if (_isHelpVisible) {
                    var (keys, descriptions) = Common.HelpManager.GetHelpData(_pageName);
                    HelpTextBlock1.Text = string.Join(Environment.NewLine, keys);
                    HelpTextBlock2.Text = string.Join(Environment.NewLine, descriptions);
                }
                HotKeyHelpScrollViewer.Height = page.Height + (_subMenu?.ActualHeight ?? 0);
            }
        }
        private void OnHelpButtonClicked(object? sender, EventArgs e) {
            _isHelpVisible = !_isHelpVisible;

            var margin = _isHelpVisible ? new Thickness(10) : new Thickness(0);
            HelpTextBlock1.Margin = margin;
            HelpTextBlock2.Margin = margin;

            if (_isHelpVisible) {
                var (keys, descriptions) = Common.HelpManager.GetHelpData(_pageName);
                HelpTextBlock1.Text = string.Join(Environment.NewLine, keys);
                HelpTextBlock2.Text = string.Join(Environment.NewLine, descriptions);
            }
            else {
                HelpTextBlock1.Text = string.Empty;
                HelpTextBlock2.Text = string.Empty;
            }
        }


        // ウィンドウサイズ調整
        public static void AdjustWindowSizeToUserControl(Window parentWindow) {
            if (parentWindow != null) {
                parentWindow.SizeToContent = SizeToContent.WidthAndHeight;
            }
        }

        // コンボボックス更新
        public static void UpdateComboBox(System.Windows.Controls.ComboBox comboBox, string category, List<int> signalTypes, string name) {
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
        // VisaAddress取得
        public static void GetVisaAddress(InstClass instClass, System.Windows.Controls.ComboBox comboBox) {
            instClass.ResetProperties();

            instClass.Name = comboBox.Text;
            instClass.Index = comboBox.SelectedIndex;

            if (instClass.Index <= 0) { return; }

            var dRows = VisaAddressDataTable.Select($"Name = '{instClass.Name}'");
            instClass.Category = dRows[0]["Category"] as string ?? string.Empty;
            instClass.VisaAddress = dRows[0]["VisaAddress"] as string ?? string.Empty;
            instClass.SignalType = dRows[0]["SignalType"] != DBNull.Value ? Convert.ToInt32(dRows[0]["SignalType"]) : 0;
        }

        // HotKeyの登録
        public static void SetHotKey() {

            Source = HwndSource.FromHwnd(HWnd);
            Source.AddHook(HwndHook);

            // ホットキーを登録
            foreach (var hotkey in HotkeysList) {
                RegisterHotKey(HWnd, hotkey.Id, hotkey.Modifier, (uint)hotkey.VirtualKey);
            }
        }
        public static void ClearHotKey() {
            foreach (var hotkey in HotkeysList) {
                UnregisterHotKey(HWnd, hotkey.Id);
            }
            HotkeysList.Clear();
        }
        private static IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
            if (msg == WmHotKey) {
                int id = wParam.ToInt32();

                var hotkey = HotkeysList.FirstOrDefault(h => h.Id == id);
                hotkey?.Action.Invoke(); // ホットキーに設定されたアクションを実行
                handled = true;
            }
            return IntPtr.Zero;
        }

        // デバイス接続
        public static async Task<string> ConnectDeviceAsync(InstClass instClass) {
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
        public static async Task<string> ConnectDeviceVisaAsync(InstClass instClass, bool hasInput) {
            return await Task.Run(() => {
                using var usbDev = new USBDeviceManager();
                usbDev.OpenDev(instClass.VisaAddress);
                usbDev.OutputDev(instClass.InstCommand);
                return hasInput ? usbDev.InputDev() : "";
            });
        }
        // ADC接続
        public static async Task<string> ConnectDeviceAdcAsync(InstClass instClass) {
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

        // DMM測定値取得
        public static async Task<decimal> ReadDmm(DmmInstClass dmmInstClass) {

            dmmInstClass.InstCommand = dmmInstClass.SignalType switch {
                1 => string.Empty,
                2 => "FETC?",
                _ => throw new ApplicationException(),
            };

            var result = await ConnectDeviceAsync(dmmInstClass);
            decimal.TryParse(result, NumberStyles.AllowExponent | NumberStyles.Float, CultureInfo.InvariantCulture, out var output);

            return output;
        }

        // FG切り替え
        public static async Task RotationFgAsync(FgInstClass fgInstClass) {
            await ConnectDeviceAsync(fgInstClass);
        }

        // OSC測定値取得
        public static async Task<decimal> ReadOsc(OscInstClass oscInstClass, int oscMeas) {

            oscInstClass.InstCommand = $"MEASU:MEAS{oscMeas}:VAL?";
            var result = await ConnectDeviceAsync(oscInstClass);
            decimal.TryParse(result, NumberStyles.AllowExponent | NumberStyles.Float, CultureInfo.InvariantCulture, out var output);

            return output;
        }

        // OSC切り替え
        public static async Task RotationOscAsync(OscInstClass oscInstClass) {
            await ConnectDeviceAsync(oscInstClass);
        }

    }
}