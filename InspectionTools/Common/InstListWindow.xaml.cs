using System.Data;
using System.Windows;

namespace InspectionTools.Common {
    /// <summary>
    /// InstListWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class InstListWindow : Window {

        private readonly DataTable _dataTable;

        public InstListWindow(DataTable dataTable) {
            InitializeComponent();
            _dataTable = dataTable;
            InstListGrid.ItemsSource = _dataTable.DefaultView;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e) {
            _dataTable.WriteXml("VisaAddress.xml");
        }
    }
}
