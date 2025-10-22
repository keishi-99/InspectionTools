using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace InspectionTools.MainMenu {
    /// <summary>
    /// SubMenuUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class SubMenuUserControl : UserControl {

        public event EventHandler? BackToMainRequested;
        public event EventHandler? HelpButtonClicked;

        public interface ISubMenuAware {
            void SetSubMenuControl(MainMenu.SubMenuUserControl? subMenu);
        }

        public SubMenuUserControl() {
            InitializeComponent();
        }

        // MainMenu表示
        private void ShowMainMenu() {
            BackToMainRequested?.Invoke(this, EventArgs.Empty);
        }

        // 機器リスト表示
        private static void ShowInstList() {
            const string XmlFilePath = "VisaAddress.xml";
            if (!System.IO.File.Exists(XmlFilePath)) {
                MessageBox.Show($"{XmlFilePath}が見つかりません。");
                return;
            }

            using DataSet dataSet = new();
            dataSet.ReadXml("VisaAddress.xml");
            DataTable dataTable = dataSet.Tables[0];

            Common.InstListWindow frm1 = new(dataTable);
            frm1.ShowDialog();
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
        private void HelpButton_Click(object sender, RoutedEventArgs e) {
            HelpButtonClicked?.Invoke(this, EventArgs.Empty);
        }
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
