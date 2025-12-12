using System.IO;
using System.Windows;
using System.Windows.Threading;
using Application = System.Windows.Application;

namespace InspectionTools {
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application {

        public App() {
            // UI スレッドの未処理例外
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;

            // 非 UI スレッド（Taskなど）の例外
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e) {
            var ex = e.Exception;

            string methodName = ex.TargetSite?.DeclaringType?.FullName + "." + ex.TargetSite?.Name;

            System.Windows.MessageBox.Show($"エラー発生メソッド:\n{methodName}\n\n{ex.Message}");
            e.Handled = true; // アプリが落ちるのを防ぐ
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e) {
            if (e.ExceptionObject is Exception ex) {
                ShowError(ex);
            }
        }

        private static void ShowError(Exception ex) {
            System.Windows.MessageBox.Show(
                $"予期しないエラーが発生しました。\n\n{ex.Message}",
                "エラー",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            );

            // ログ保存例
            File.AppendAllText("error.log", $"{DateTime.Now}\n{ex}\n\n");
        }

    }

}
