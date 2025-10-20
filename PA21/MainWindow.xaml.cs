using System.Windows;

namespace PA21 {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        public MainWindow() {
            InitializeComponent();
        }


        private void Page1_Click(object sender, RoutedEventArgs e) {
            MainFrame.Navigate(new PA21Detector());
        }

        private void Page2_Click(object sender, RoutedEventArgs e) {
            MainFrame.Navigate(new PA21Transducer());
        }



    }
}