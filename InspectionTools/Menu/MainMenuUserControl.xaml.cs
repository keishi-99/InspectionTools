using System.Windows;
using System.Windows.Controls;

namespace InspectionTools.MainMenu {
    /// <summary>
    /// MainMenuUserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class MainMenuUserControl : UserControl {
        public event Action<string>? PageSelected;

        public MainMenuUserControl() {
            InitializeComponent();
        }

        private void DFPDX_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("DFPDX");
        }
        private void EL0122FI_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("EL0122FI");
        }
        private void EL0122_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("EL0122");
        }
        private void EL0137_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("EL0137");
        }
        private void EL1812_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("EL1812");
        }
        private void EL3801_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("EL3801");
        }
        private void EL4001_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("EL4001");
        }
        private void EL9100_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("EL9100");
        }
        private void EL9240_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("EL9240");
        }
        private void MassFlow_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("MassFlow");
        }
        private void PA14_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("PA14");
        }
        private void PAF5amp_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("PAF5amp");
        }
        private void PAF5_Click(object sender, RoutedEventArgs e) {
            PageSelected?.Invoke("PAF5");
        }

    }
}
