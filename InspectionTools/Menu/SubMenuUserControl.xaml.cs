using InspectionTools.Common;
using MaterialDesignThemes.Wpf;
using System.Data;
using System.Windows;
using System.Windows.Threading;
using Button = System.Windows.Controls.Button;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.MainMenu {
    /// <summary>
    /// SubMenuUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class SubMenuUserControl : UserControl {

        public event EventHandler? BackToMainRequested;
        public event EventHandler? HelpCheckBoxChecked;
        public event EventHandler? HelpCheckBoxUnchecked;

        private static bool s_themeMode = false;

        // タイマー
        private readonly DispatcherTimer _timer;

        public interface ISubMenuAware {
            void SetSubMenuControl(MainMenu.SubMenuUserControl? subMenu);
        }

        public SubMenuUserControl() {
            InitializeComponent();
            LoadEvents();

            _timer = new DispatcherTimer {
                Interval = TimeSpan.FromSeconds(1)
            };
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        private void LoadEvents() {
            const string XmlFilePath = "VisaAddress.xml";
            if (!System.IO.File.Exists(XmlFilePath)) {
                MessageBox.Show($"{XmlFilePath}が見つかりません。");
                return;
            }
            using DataSet dataSet = new();
            dataSet.ReadXml(XmlFilePath);
            MainWindow.VisaAddressDataTable = dataSet.Tables[0];

            ThemeModeCheckBox.IsChecked = s_themeMode;
        }

        // MainMenu表示
        private void ShowMainMenu() {
            BackToMainRequested?.Invoke(this, EventArgs.Empty);
        }

        // 機器リスト表示
        private void ShowInstList() {
            Common.InstListWindow frm1 = new();
            frm1.ShowDialog();
            LoadEvents();
        }

        // テーマ切り替え
        internal static void OnIsDarkModeChanged(bool value) {
            var theme = ThemeHelper.GetBundledTheme();
            theme.BaseTheme = value ? BaseTheme.Dark : BaseTheme.Light;
            ThemeHelper.SetBundledTheme(theme);
        }

        // ボタン名を指定して有効/無効を切り替えるメソッド
        public void SetButtonEnabled(string buttonName, bool isEnabled) {
            if (FindName(buttonName) is Button button) {
                button.IsEnabled = isEnabled;
            }
        }

        // イベントハンドラ
        private void ProductListButton_Click(object sender, RoutedEventArgs e) {
            ShowMainMenu();
        }
        private void InstListButton_Click(object sender, RoutedEventArgs e) {
            ShowInstList();
        }
        private void HelpCheckBox_Checked(object sender, RoutedEventArgs e) {
            HelpCheckBoxChecked?.Invoke(this, EventArgs.Empty);
        }
        private void HelpCheckBox_Unchecked(object sender, RoutedEventArgs e) {
            HelpCheckBoxUnchecked?.Invoke(this, EventArgs.Empty);
        }
        private void TopMostCheckBox_Checked(object sender, RoutedEventArgs e) {
            var parentWindow = Window.GetWindow(this);
            parentWindow.Topmost = true;
        }
        private void TopMostCheckBox_Unchecked(object sender, RoutedEventArgs e) {
            var parentWindow = Window.GetWindow(this);
            parentWindow.Topmost = false;
        }
        private void ThemeModeCheckBox_Checked(object sender, RoutedEventArgs e) {
            s_themeMode = true;
            OnIsDarkModeChanged(true);
        }
        private void ThemeModeCheckBox_Unchecked(object sender, RoutedEventArgs e) {
            s_themeMode = false;
            OnIsDarkModeChanged(false);
        }
        private void Timer_Tick(object? sender, EventArgs e) { Time.Text = DateTime.Now.ToString("HH:mm:ss"); }


    }
}
