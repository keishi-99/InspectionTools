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

        // アプリ起動時の初期化処理（メインメニュー表示と計測器リスト読み込み）
        private void LoadEvents() {
            ShowMainMenu();
            LoadInstList();
        }
        // XMLファイルからVISAアドレス一覧を読み込んでDataTableに格納する
        private static void LoadInstList() {
            const string XmlFilePath = "VisaAddress.xml";
            if (!System.IO.File.Exists(XmlFilePath)) {
                MessageBox.Show($"{XmlFilePath}が見つかりません。");
                return;
            }
            using DataSet dataSet = new();
            dataSet.ReadXml(XmlFilePath);
            // DataSet から切り離してから静的フィールドへ格納する
            // （切り離さないと DataTable が DataSet への参照を保持し続け、
            //   using による DataSet Dispose 後も GC されなくなる）
            var table = dataSet.Tables[0];
            dataSet.Tables.Remove(table);
            MainWindow.VisaAddressDataTable = table;
        }

        // ドロワーを閉じてメインメニューUserControlを表示する
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

        // ページ名に対応したUserControlを生成してコンテンツエリアに切り替える
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

        // 指定名のButtonの有効/無効を切り替える
        public void SetButtonEnabled(string buttonName, bool isEnabled) {
            if (FindName(buttonName) is Button button) {
                button.IsEnabled = isEnabled;
            }
        }

        // 親ウィンドウのサイズをUserControlのコンテンツに合わせて自動調整する
        public static void AdjustWindowSizeToUserControl(Window parentWindow) {
            parentWindow?.SizeToContent = SizeToContent.WidthAndHeight;
        }

        // カテゴリと信号種別でフィルタリングした計測器名をコンボボックスに設定する
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

        // コンボボックスで選択した計測器のVISAアドレスとプロパティをInstClassに設定する
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

        // HotkeysList に登録されたホットキーをすべてWindowsに登録する
        public static void SetHotKey() {

            Source = HwndSource.FromHwnd(HWnd);
            // 重複登録を防ぐため、追加前に一度削除する
            // FromHwnd が null を返す場合に備えて ?. で安全に呼び出す
            Source?.RemoveHook(HwndHook);
            Source?.AddHook(HwndHook);

            foreach (var hotkey in HotkeysList) {
                RegisterHotKey(HWnd, hotkey.Id, hotkey.Modifier, (uint)hotkey.VirtualKey);
            }
        }
        // 登録済みホットキーをすべてWindowsから解除してリストをクリアする
        public static void ClearHotKey() {
            foreach (var hotkey in HotkeysList) {
                UnregisterHotKey(HWnd, hotkey.Id);
            }
            HotkeysList.Clear();
            // フックを削除してリークを防ぐ
            Source?.RemoveHook(HwndHook);
        }
        // Windowsメッセージをフックしてホットキーイベントを処理する
        private static IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
            if (msg == WmHotKey) {
                int id = wParam.ToInt32();
                var hotkey = HotkeysList.FirstOrDefault(h => h.Id == id);
                hotkey?.Action.Invoke();
                handled = true;
            }
            return IntPtr.Zero;
        }

        // アプリのカラーテーマ（ライト/ダーク）を切り替える
        internal static void SetTheme(BaseTheme baseTheme) {
            var paletteHelper = new PaletteHelper();
            var theme = paletteHelper.GetTheme();
            theme.SetBaseTheme(baseTheme);
            paletteHelper.SetTheme(theme);
        }

        // ヘルプが表示中であれば現在ページのキーと説明文を更新する
        private void UpdateHelpText() {
            if (!_isHelpVisible) return;
            var (keys, descriptions) = Common.HelpManager.GetHelpData(_pageName);
            HelpTextBlock1.Text = string.Join(Environment.NewLine, keys);
            HelpTextBlock2.Text = string.Join(Environment.NewLine, descriptions);
        }

        // ヘルプを表示状態にしてテキストを更新する
        private void HelpCheckBoxChecked() {
            _isHelpVisible = true;
            HelpTextBlock1.Margin = new Thickness(10);
            HelpTextBlock2.Margin = new Thickness(10);
            UpdateHelpText();
        }

        // ヘルプを非表示状態にしてテキストを空にする
        private void HelpCheckBoxUnchecked() {
            _isHelpVisible = false;
            HelpTextBlock1.Margin = new Thickness(0);
            HelpTextBlock2.Margin = new Thickness(0);
            HelpTextBlock1.Text = string.Empty;
            HelpTextBlock2.Text = string.Empty;
        }

        // 機器リストウィンドウをダイアログ表示してXMLを再読み込みする
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