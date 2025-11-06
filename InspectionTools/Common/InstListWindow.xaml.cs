using System.Windows;
using static InspectionTools.MainMenu.SubMenuUserControl;

namespace InspectionTools.Common {
    /// <summary>
    /// InstListWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class InstListWindow : Window {

        public InstListWindow() {
            InitializeComponent();
            InstListGrid.ItemsSource = VisaAddressDataTable.DefaultView;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e) {
            VisaAddressDataTable.WriteXml("VisaAddress.xml");
            Close();
        }
    }
}
