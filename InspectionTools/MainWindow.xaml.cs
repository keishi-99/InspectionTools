using InspectionTools.Common;
using MaterialDesignThemes.Wpf;
using System.Data;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using static InspectionTools.Common.Win32Wrapper;
using Button = System.Windows.Controls.Button;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        private bool _isHelpVisible = false;
        private string _pageName = string.Empty;

        public interface IMainWindowAware {
            void SetMainWindow(MainWindow mainWindow);
        }

        // タイマー
        private readonly DispatcherTimer _timer;

        public const string ProductListButtonName = "ProductListButton";
        public const string InstListButtonName = "InstListButton";

        public static IntPtr HWnd { get; set; }
        public static HwndSource? Source { get; set; }
        public static List<Hotkey> HotkeysList { get; set; } = [];

        public static DataTable VisaAddressDataTable { get; set; } = new();

        public static volatile bool IsProcessing = false;

        public MainWindow() {
            InitializeComponent();
            HelpManager.LoadHelpFile("help.json");
            LoadEvents();

            // Window が完全に作られたあとにハンドルを取得
            Loaded += (s, e) => { HWnd = new WindowInteropHelper(this).Handle; };

            _timer = new DispatcherTimer {
                Interval = TimeSpan.FromSeconds(1)
            };
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        // ----- 初期化 -----
        private void LoadEvents() {
            ShowMainMenu();
            LoadInstList();
        }
        private static void LoadInstList() {
            const string XmlFilePath = "VisaAddress.xml";
            if (!System.IO.File.Exists(XmlFilePath)) {
                MessageBox.Show($"{XmlFilePath}が見つかりません。");
                return;
            }
            using DataSet dataSet = new();
            dataSet.ReadXml(XmlFilePath);
            MainWindow.VisaAddressDataTable = dataSet.Tables[0];
        }

        // ----- ページ遷移 -----
        private void ShowMainMenu() {
            MainWindowDrawer.IsLeftDrawerOpen = false;

            SetButtonEnabled("ProductListButton", false);
            SetButtonEnabled("InstListButton", true);

            var mainMenu = new MainMenu.MainMenuUserControl();
            mainMenu.PageSelected += OnPageSelected;
            MainMenuContentArea.Content = mainMenu;

            this.Title = "Menu";
            _pageName = "MainMenu";
            UpdateHelpText();
            HotKeyHelpScrollViewer.Height = mainMenu.Height;
        }

        private void OnPageSelected(string pageName) {
            SetButtonEnabled("ProductListButton", true);
            SetButtonEnabled("InstListButton", false);

            UserControl? page = pageName switch {
                "EL0122FI" => new Product.EL0122FIUserControl(),
                "EL0122" => new Product.EL0122UserControl(),
                "EL0137" => new Product.EL0137UserControl(),
                "EL1812" => new Product.EL1812UserControl(),
                "EL3801" => new Product.EL3801UserControl(),
                "EL4001" => new Product.EL4001UserControl(),
                "EL5000" => new Product.EL5000UserControl(),
                "EL9100" => new Product.EL9100UserControl(),
                "EL9220" => new Product.EL9220UserControl(),
                "EL9230" => new Product.EL9230UserControl(),
                "EL9240" => new Product.EL9240UserControl(),
                "PA14" => new Product.PA14UserControl(),
                "PA25" => new Product.PA25UserControl(),
                "PA2B" => new Product.PA2BUserControl(),
                "PAF5amp" => new Product.PAF5ampUserControl(),
                "PAF5" => new Product.PAF5UserControl(),
                "DFPDX" => new Product.DFPDXUserControl(),
                "MassFlow" => new Product.MassFlowUserControl(),
                _ => null
            };

            if (page is null) return;

            if (page is IMainWindowAware aware) {
                aware.SetMainWindow(this);
            }

            this.Title = pageName;
            _pageName = pageName;
            MainMenuContentArea.Content = page;
            UpdateHelpText();
            HotKeyHelpScrollViewer.Height = page.Height;
        }

        // ----- UI ユーティリティ -----
        public void SetButtonEnabled(string buttonName, bool isEnabled) {
            if (FindName(buttonName) is Button button) {
                button.IsEnabled = isEnabled;
            }
        }

        public static void AdjustWindowSizeToUserControl(Window parentWindow) {
            parentWindow?.SizeToContent = SizeToContent.WidthAndHeight;
        }

        // ---- 機器リスト ----
        public static void UpdateComboBox(
            System.Windows.Controls.ComboBox comboBox,
            string category,
            List<int> signalTypes) {

            if (VisaAddressDataTable == null) return;

            var rows = VisaAddressDataTable.AsEnumerable();
            var collection = signalTypes
                .SelectMany(st => rows
                    .Where(row =>
                        (row["Category"] as string) == category &&
                        row["SignalType"] != DBNull.Value &&
                        Convert.ToInt32(row["SignalType"]) == st)
                    .Select(row => row["Name"].ToString() ?? string.Empty))
                .ToList();

            comboBox.ItemsSource = collection;
        }

        public static void GetVisaAddress(
            InstClass instClass,
            System.Windows.Controls.ComboBox comboBox) {

            instClass.ResetProperties();
            instClass.Name = comboBox.Text;
            instClass.Index = comboBox.SelectedIndex;

            if (instClass.Index == -1) return;

            var dRow = VisaAddressDataTable.AsEnumerable()
                .FirstOrDefault(row => (row["Name"] as string) == instClass.Name);
            if (dRow == null) return;

            instClass.Category = dRow["Category"] as string ?? string.Empty;
            instClass.VisaAddress = dRow["VisaAddress"] as string ?? string.Empty;
            instClass.SignalType = dRow["SignalType"] != DBNull.Value
                ? Convert.ToInt32(dRow["SignalType"])
                : 0;
        }

        // ----- ホットキー -----
        public static void SetHotKey() {

            Source = HwndSource.FromHwnd(HWnd);
            Source.AddHook(HwndHook);

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
                hotkey?.Action.Invoke();
                handled = true;
            }
            return IntPtr.Zero;
        }

        // ----- テーマ -----
        internal static void SetTheme(BaseTheme baseTheme) {
            var paletteHelper = new PaletteHelper();
            var theme = paletteHelper.GetTheme();
            theme.SetBaseTheme(baseTheme);
            paletteHelper.SetTheme(theme);
        }

        // ----- ヘルプ -----
        private void UpdateHelpText() {
            if (!_isHelpVisible) return;
            var (keys, descriptions) = Common.HelpManager.GetHelpData(_pageName);
            HelpTextBlock1.Text = string.Join(Environment.NewLine, keys);
            HelpTextBlock2.Text = string.Join(Environment.NewLine, descriptions);
        }

        private void HelpCheckBoxChecked() {
            _isHelpVisible = true;
            HelpTextBlock1.Margin = new Thickness(10);
            HelpTextBlock2.Margin = new Thickness(10);
            UpdateHelpText();
        }

        private void HelpCheckBoxUnchecked() {
            _isHelpVisible = false;
            HelpTextBlock1.Margin = new Thickness(0);
            HelpTextBlock2.Margin = new Thickness(0);
            HelpTextBlock1.Text = string.Empty;
            HelpTextBlock2.Text = string.Empty;
        }

        // ----- 機器リスト -----
        private void ShowInstList() {
            Common.InstListWindow frm1 = new() {
                Owner = this
            };
            frm1.ShowDialog();
            LoadInstList();
            // ドロワーを閉じる
            MainWindowDrawer.IsLeftDrawerOpen = false;

        }

        // ----- イベントハンドラ -----
        private void ProductListButton_Click(object sender, RoutedEventArgs e) => ShowMainMenu();
        private void InstListButton_Click(object sender, RoutedEventArgs e) => ShowInstList();
        private void HelpCheckBox_Checked(object sender, RoutedEventArgs e) => HelpCheckBoxChecked();
        private void HelpCheckBox_Unchecked(object sender, RoutedEventArgs e) => HelpCheckBoxUnchecked();
        private void TopMostCheckBox_Checked(object sender, RoutedEventArgs e) { Window.GetWindow(this).Topmost = true; }
        private void TopMostCheckBox_Unchecked(object sender, RoutedEventArgs e) { Window.GetWindow(this).Topmost = false; }
        private void ThemeToggle_Loaded(object sender, RoutedEventArgs e) {
            var paletteHelper = new PaletteHelper();
            var theme = paletteHelper.GetTheme();
            ThemeToggle.IsChecked = theme.GetBaseTheme() == BaseTheme.Dark;
        }
        private void ThemeToggle_Checked(object sender, RoutedEventArgs e) => SetTheme(BaseTheme.Dark);
        private void ThemeToggle_Unchecked(object sender, RoutedEventArgs e) => SetTheme(BaseTheme.Light);
        private void Timer_Tick(object? sender, EventArgs e) { Time.Text = DateTime.Now.ToString("HH:mm:ss"); }


    }
}