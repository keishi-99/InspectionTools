using System.Windows;

namespace InspectionTools.Common {
    /// <summary>
    /// InstListWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class InstListWindow : Window {

        public InstListWindow() {
            InitializeComponent();
            InstListGrid.ItemsSource = MainWindow.VisaAddressDataTable.DefaultView;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e) {
            MainWindow.VisaAddressDataTable.WriteXml("VisaAddress.xml");
            Close();
        }
    }
}
