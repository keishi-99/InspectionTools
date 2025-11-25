using InspectionTools.Common;
using System.Data;
using System.IO;
using System.Windows;
using Tesseract;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;
using static InspectionTools.MainWindow;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.Product {
    /// <summary>
    /// EL9100UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL9100UserControl : UserControl, IMainWindowAware {

        private MainWindow? _mainWindow;
        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DmmInstClass _instDmm01 = new();
        private readonly DmmInstClass _instDmm02 = new();

        public EL9100UserControl() {
            InitializeComponent();
        }

        // 起動時
        private void LoadEvents() {
            InstListImport();
            FormatSet();
            var parentWindow = Window.GetWindow(this);
            MainWindow.AdjustWindowSizeToUserControl(parentWindow);
        }
        private void InstListImport() {
            // デジタルマルチメータ、ファンクションジェネレータ、オシロスコープのコンボボックスを更新する
            MainWindow.UpdateComboBox(Dmm01ComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(Dmm02ComboBox, "デジタルマルチメータ", [1, 2]);
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            MainWindow.IsProcessing = isVisible;
            MainGrid.IsEnabled = !isVisible;
            ProgressRing.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }

        // 選択した機器のVisaAddressを取得
        private void SelectInst() {
            MainWindow.GetVisaAddress(_instDmm01, Dmm01ComboBox);
            MainWindow.GetVisaAddress(_instDmm02, Dmm02ComboBox);
        }
        // 機器初期設定
        private void FormatSet() {
            _instDmm01.InstCommand = _instDmm01.SignalType switch {
                1 => "*RST,F5,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 0.02;*OPC?",
                _ => string.Empty,
            };
            _instDmm02.InstCommand = _instDmm02.SignalType switch {
                1 => "*RST,F5,R7,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 0.2;*OPC?",
                _ => string.Empty,
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {
            try {
                _mainWindow?.SetButtonEnabled("ProductListButton", false);

                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                CheckDmmId();
                FormatSet();

                InstClass[] devices = [_instDmm01, _instDmm02];
                var tasks = devices.Select(device => MainWindow.ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);

                Dmm01ComboBox.IsEnabled = false;
                Dmm02ComboBox.IsEnabled = false;
                ConnectButton.IsEnabled = false;
                ReleaseButton.IsEnabled = true;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー");
            } finally {
                VisibleProgressImage(false);
            }
        }
        // DMMのIDチェック処理
        private void CheckDmmId() {
            var indices = new[] { _instDmm01.Index, _instDmm02.Index }
                .Where(i => i >= 1); // 未選択(0以下)は無視

            if (indices.Count() != indices.Distinct().Count()) {
                throw new Exception("同じ測定器が選択されています。");
            }
        }

        // 解除
        private void Release() {
            VisibleProgressImage(false);

            _instDmm01.ResetProperties();
            _instDmm02.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
            Dmm01ComboBox.IsEnabled = true;
            Dmm02ComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyChekBox.IsChecked = false;
        }

        // DMM測定値取得
        private async Task<decimal> ReadDmm(DmmInstClass dmmInstClass) {
            try {
                VisibleProgressImage(true);

                var output = await MainWindow.ReadDmm(dmmInstClass);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // DMM切り替え
        private async Task SwitchDmm2(DmmInstClass dmmInstClass, string func) {
            try {
                VisibleProgressImage(true);

                dmmInstClass.InstCommand = func switch {
                    "DCI" => dmmInstClass.SignalType switch {
                        1 => "*RST,F5,R7,*OPC?",
                        2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 0.2;*OPC?",
                        _ => throw new ApplicationException(),
                    },
                    "DCV" => dmmInstClass.SignalType switch {
                        1 => "*RST,F1,R6,*OPC?",
                        2 => "*RST;:INIT:CONT 1;:VOLT:DC:RANG 20;*OPC?",
                        _ => throw new ApplicationException(),
                    },
                    _ => throw new ApplicationException(),
                };

                await MainWindow.ConnectDeviceAsync(dmmInstClass);

            } finally {
                VisibleProgressImage(false);
            }
        }

        // OCR処理
        private void Capture() {
            var captureWindow = new ScreenCaptureWindow();
            using Bitmap? captured = captureWindow.Capture("EL9100", 150, 30);
            if (captured == null) {
                OcrResult.Text = "キャプチャがキャンセルされました。";
                return;
            }

            // 画像ファイルをOCRを実行
            string ocrResult;
            try {
                var returnOcr = PerformOCR(captured);
                ocrResult = returnOcr.Replace(" ", string.Empty);
                //ocrResult = new string([.. returnOcr.Where(c => !char.IsWhiteSpace(c))]);
            } finally {
                captured.Dispose();
            }

            // 結果を表示
            OcrResult.Text = ocrResult;
        }
        public static string PerformOCR(Bitmap image) {
            try {
                // tessdata フォルダのパス（exe と同じ階層に配置）
                string tessDataPath = Path.Combine(Environment.CurrentDirectory, "tessdata");

                // OCRエンジンの初期化（日本語＋英語）

                using var engine = new TesseractEngine(tessDataPath, "jpn+eng", EngineMode.Default);

                using var ms = new MemoryStream();
                // Bitmap → メモリストリーム（PNG形式）
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;

                // MemoryStream → Pix に変換
                using var pix = Pix.LoadFromMemory(ms.ToArray());
                using var page = engine.Process(pix);
                // 認識結果のテキストを取得
                string text = page.GetText();
                return text.Trim();
            } catch (Exception ex) {
                return $"OCRエラー: {ex.Message}";
            }
        }

        // DMM01測定値コピー
        private async void ActionHotkeyBracketR() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm01);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000000).ToString("0"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM02測定値コピー
        private async void ActionHotkeySlash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm02);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry((output * 1000).ToString("0.0"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private async void ActionHotkeyBackslash() {
            if (MainWindow.IsProcessing) { return; }

            try {
                var output = await ReadDmm(_instDmm02);

                var sim = new InputSimulator();
                sim.Keyboard.TextEntry(output.ToString("0.00"));
                await Task.Delay(100);
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM2切替(DCI)
        private async void ActionHotkeyAtsign() {
            if (MainWindow.IsProcessing) { return; }

            try {
                await SwitchDmm2(_instDmm02, "DCI");
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM2切替(DCV)
        private async void ActionHotkeyBracketL() {
            if (MainWindow.IsProcessing) { return; }

            try {
                await SwitchDmm2(_instDmm02, "DCV");
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // HotKeyの登録
        private void SetHotKey() {
            MainWindow.HotkeysList.Clear();

            if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                    new(ModNone, HotkeyBracketL, ActionHotkeyBracketL ),
                ]);
            }

            MainWindow.SetHotKey();
        }
        private static void ClearHotKey() {
            MainWindow.ClearHotKey();
        }

        // イベントハンドラ
        private void UserControl_Loaded(object sender, RoutedEventArgs e) { LoadEvents(); }
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void OcrButton_Click(object sender, RoutedEventArgs e) { Capture(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }


    }
}