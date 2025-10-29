using InspectionTools.Common;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Interop;
using Tesseract;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;
using static InspectionTools.MainMenu.SubMenuUserControl;
using ComboBox = System.Windows.Controls.ComboBox;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.Product {
    /// <summary>
    /// EL9100UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL9100UserControl : UserControl, ISubMenuAware {

        private MainMenu.SubMenuUserControl? _subMenu;
        public void SetSubMenuControl(MainMenu.SubMenuUserControl? subMenu) {
            _subMenu = subMenu;
        }

        private IntPtr _hWnd = IntPtr.Zero;

        private readonly InstClass _instDmm01;
        private readonly InstClass _instDmm02;

        public ObservableCollection<string> Dmm1List { get; } = [];
        public ObservableCollection<string> Dmm2List { get; } = [];

        public EL9100UserControl() {
            InitializeComponent();
            _instDmm01 = new();
            _instDmm02 = new();
            LoadEvents();
            // 親ウィンドウを取得
            var parentWindow = Window.GetWindow(this);
            if (parentWindow != null) {
                _hWnd = new WindowInteropHelper(parentWindow).Handle;
            }

            Loaded += (s, e) => AdjustWindowSizeToUserControl();
        }
        private void AdjustWindowSizeToUserControl() {
            var parentWindow = Window.GetWindow(this);
            if (parentWindow != null) {
                parentWindow.SizeToContent = SizeToContent.WidthAndHeight;
            }
        }

        public class InstClass {
            internal USBDeviceManager UsbDev { get; set; } = new();
            public string Category { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public string VisaAddress { get; set; } = string.Empty;
            public int SignalType { get; set; } = 0;
            public int Index { get; set; } = 0;
            public string InstCommand { get; set; } = string.Empty;
            public int SettingNumber { get; set; } = 0;

            public void ResetProperties() {
                UsbDev = new();
                Name = string.Empty;
                Category = string.Empty;
                VisaAddress = string.Empty;
                SignalType = 0;
                Index = 0;
                InstCommand = string.Empty;
                SettingNumber = 0;
            }
            public void Dispose() {
                // UsbDevの解放処理
                UsbDev?.Dispose();
            }
        }

        private const int TimeOut = 3;    //タイムアウトまでの時間(sec)

        internal DataTable _dataTable = new();

        private readonly List<Hotkey> _hotkeys = [];
        private HwndSource? _source;

        private volatile bool _isProcessing = false;

        private static readonly SemaphoreSlim s_semaphore = new(1, 1); // 最大1つの接続

        // 起動時
        private void LoadEvents() {
            InstListImport();
            FormatSet();
        }
        private void InstListImport() {
            const string XmlFilePath = "VisaAddress.xml";
            if (!System.IO.File.Exists(XmlFilePath)) {
                MessageBox.Show($"{XmlFilePath}が見つかりません。");
                return;
            }

            using DataSet dataSet = new();
            dataSet.ReadXml("VisaAddress.xml");
            _dataTable = dataSet.Tables[0];

            // デジタルマルチメータ、ファンクションジェネレータ、オシロスコープのコンボボックスを更新する
            UpdateComboBox(Dmm01ComboBox, Dmm1List, "デジタルマルチメータ", [1, 2], "[DMM1]");
            UpdateComboBox(Dmm02ComboBox, Dmm2List, "デジタルマルチメータ", [1, 2], "[DMM2]");
        }
        private void UpdateComboBox(ComboBox comboBox, ObservableCollection<string> collection, string category, List<int> signalTypes, string name) {
            if (_dataTable == null) {
                return;
            }

            collection.Clear();
            collection.Add(name);

            foreach (var signalType in signalTypes) {
                var rows = _dataTable.Select($"Category = '{category}' AND SignalType = {signalType}");
                foreach (var d in rows) {
                    collection.Add(d["Name"].ToString() ?? string.Empty);
                }
            }
            comboBox.ItemsSource = collection;
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            _isProcessing = isVisible;
            MainGrid.IsEnabled = !isVisible;
            ProgressRing.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }

        // 選択した機器のVisaAddressを取得
        private void SelectInst() {
            GetVisaAddress(_instDmm01, Dmm01ComboBox);
            GetVisaAddress(_instDmm02, Dmm02ComboBox);
        }
        private void GetVisaAddress(InstClass instClass, ComboBox comboBox) {
            instClass.ResetProperties();

            instClass.Name = comboBox.Text;
            instClass.Index = comboBox.SelectedIndex;

            if (instClass.Index <= 0) { return; }

            var dRows = _dataTable.Select($"Name = '{instClass.Name}'");
            instClass.Category = dRows[0]["Category"] as string ?? string.Empty;
            instClass.VisaAddress = dRows[0]["VisaAddress"] as string ?? string.Empty;
            instClass.SignalType = dRows[0]["SignalType"] != DBNull.Value ? Convert.ToInt32(dRows[0]["SignalType"]) : 0;
        }
        // 機器初期設定
        private void FormatSet() {
            _instDmm01.InstCommand = _instDmm01.SignalType switch {
                1 => "*RST,F5,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 0.02;*OPC?",
                _ => string.Empty,
            };
            _instDmm02.InstCommand = _instDmm02.SignalType switch {
                1 => "*RST,F5,R8,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 10;*OPC?",
                _ => string.Empty,
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {
            try {
                _subMenu?.SetButtonEnabled("ProductListButton", false);
                _subMenu?.SetButtonEnabled("InstListButton", false);

                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                CheckDmmId();
                FormatSet();

                var devices = new[] { _instDmm01, _instDmm02 };
                var tasks = devices.Select(device => ConnectDeviceAsync(device));
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
        // デバイス接続
        private static async Task<string> ConnectDeviceAsync(InstClass instClass) {
            return instClass.Index < 1
                ? ""
                : instClass.SignalType switch {
                    1 => await ConnectDeviceAdcAsync(instClass),
                    2 or 4 => await ConnectDeviceVisaAsync(instClass, true),
                    3 => await ConnectDeviceVisaAsync(instClass, false),
                    _ => throw new ApplicationException(),
                };
        }
        // Visa接続
        private static async Task<string> ConnectDeviceVisaAsync(InstClass instClass, bool hasInput) {
            return await Task.Run(() => {
                using var usbDev = new USBDeviceManager();
                usbDev.OpenDev(instClass.VisaAddress);
                usbDev.OutputDev(instClass.InstCommand);
                return hasInput ? usbDev.InputDev() : "";
            });
        }
        // ADC接続
        private static async Task<string> ConnectDeviceAdcAsync(InstClass instClass) {
            await s_semaphore.WaitAsync();
            try {
                uint hDev = 0;
                var rcvDt = "";
                uint rcvLen = 50;
                var id = uint.Parse(instClass.VisaAddress);
                try {
                    if (AusbWrapper.Start(TimeOut) != 0 || AusbWrapper.Open(ref hDev, id) != 0) { throw new Exception("開始できません"); }
                    if (!string.IsNullOrEmpty(instClass.InstCommand)) {
                        if (AusbWrapper.Write(hDev, instClass.InstCommand) != 0) { throw new Exception("コマンドの送信に失敗しました"); }
                    }
                    if (AusbWrapper.Read(hDev, ref rcvDt, ref rcvLen) != 0) { throw new Exception("メッセージの受信に失敗しました"); }
                } finally {
                    _ = AusbWrapper.Close(hDev);
                    _ = AusbWrapper.End();
                }
                return rcvDt;
            } finally {
                s_semaphore.Release();
            }
        }

        // 解除
        private void Release() {
            VisibleProgressImage(false);

            _instDmm01.ResetProperties();
            _instDmm02.ResetProperties();

            _subMenu?.SetButtonEnabled("ProductListButton", true);
            _subMenu?.SetButtonEnabled("InstListButton", true);
            Dmm01ComboBox.IsEnabled = true;
            Dmm02ComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyChekBox.IsChecked = false;
        }

        // DMM測定値取得
        private async Task<decimal> ReadDmm(InstClass instClass) {
            try {
                VisibleProgressImage(true);

                instClass.InstCommand = instClass.SignalType switch {
                    1 => string.Empty,
                    2 => "FETC?",
                    _ => throw new ApplicationException(),
                };

                var result = await ConnectDeviceAsync(instClass);
                decimal.TryParse(result, NumberStyles.AllowExponent | NumberStyles.Float, CultureInfo.InvariantCulture, out var output);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // DMM切り替え
        private async Task SwitchDmm2(InstClass instClass, string func) {
            try {
                VisibleProgressImage(true);

                instClass.InstCommand = func switch {
                    "DCI" => instClass.SignalType switch {
                        1 => "*RST,F5,R8,*OPC?",
                        2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;:CURR:DC:RANG 10;*OPC?",
                        _ => throw new ApplicationException(),
                    },
                    "DCV" => instClass.SignalType switch {
                        1 => "*RST,F1,R6,*OPC?",
                        2 => "*RST;:INIT:CONT 1;:VOLT:DC:RANG 20;*OPC?",
                        _ => throw new ApplicationException(),
                    },
                    _ => throw new ApplicationException(),
                };

                await ConnectDeviceAsync(instClass);

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

                using (var engine = new TesseractEngine(tessDataPath, "jpn+eng", EngineMode.Default))

                using (var ms = new MemoryStream()) {
                    // Bitmap → メモリストリーム（PNG形式）
                    image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    ms.Position = 0;

                    // MemoryStream → Pix に変換
                    using (var pix = Pix.LoadFromMemory(ms.ToArray()))
                    using (var page = engine.Process(pix)) {
                        // 認識結果のテキストを取得
                        string text = page.GetText();
                        return text.Trim();
                    }
                }
            } catch (Exception ex) {
                return $"OCRエラー: {ex.Message}";
            }
        }


        //private async void Capture() {
        //    var captureWindow = new ScreenCaptureWindow();
        //    using Bitmap? captured = captureWindow.Capture();
        //    if (captured == null) {
        //        OcrResult.Text = "キャプチャがキャンセルされました。";
        //        return;
        //    }

        //    // 画像ファイルをOCRを実行
        //    string ocrResult;
        //    try {
        //        ocrResult = await PerformOCR(captured);
        //    } finally {
        //        captured.Dispose();
        //    }

        //    // 結果を表示
        //    OcrResult.Text = ocrResult;
        //}
        //private async static Task<string> PerformOCR(Bitmap image) {
        //    try {
        //        // 画像ファイルをバイト配列として読み込む
        //        byte[] imageBytes;
        //        using (var ms = new MemoryStream()) {
        //            image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        //            imageBytes = ms.ToArray();
        //        }

        //        // バイト配列をIBufferに変換
        //        IBuffer buffer = imageBytes.AsBuffer();

        //        // IBufferからSoftwareBitmapに変換
        //        SoftwareBitmap softwareBitmap;
        //        using (var stream = new InMemoryRandomAccessStream()) {
        //            await stream.WriteAsync(buffer);
        //            stream.Seek(0);
        //            var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream);
        //            softwareBitmap = await decoder.GetSoftwareBitmapAsync();
        //        }

        //        // OCRエンジンを初期化
        //        OcrEngine ocrEngine = OcrEngine.TryCreateFromUserProfileLanguages();

        //        // 画像をOCRにかけ、結果を取得
        //        OcrResult ocrResult = await ocrEngine.RecognizeAsync(softwareBitmap);

        //        // 認識結果をテキストに変換
        //        string recognizedText = ocrResult.Lines.Select(line => line.Text).Aggregate((current, next) => current + Environment.NewLine + next);

        //        return recognizedText;

        //    } catch (Exception ex) {
        //        return "OCRエラー: " + ex.Message;
        //    }
        //}

        // DMM01測定値コピー
        private async void ActionHotkeyBracketR() {
            if (_isProcessing) { return; }

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
            if (_isProcessing) { return; }

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
            if (_isProcessing) { return; }

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
            if (_isProcessing) { return; }

            try {
                await SwitchDmm2(_instDmm02, "DCI");
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // DMM2切替(DCV)
        private async void ActionHotkeyBracketL() {
            if (_isProcessing) { return; }

            try {
                await SwitchDmm2(_instDmm02, "DCV");
            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // HotkKeyの登録
        private void SetHotKey() {
            _hotkeys.Clear();
            if (!string.IsNullOrEmpty(_instDmm01.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm02.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                    new(ModNone, HotkeyBracketL, ActionHotkeyBracketL ),
                ]);
            }

            // 親ウィンドウを取得
            var parentWindow = Window.GetWindow(this);
            if (parentWindow != null) {
                _hWnd = new WindowInteropHelper(parentWindow).Handle;
                _source = HwndSource.FromHwnd(_hWnd);
                _source.AddHook(HwndHook);
            }

            // ホットキーを登録
            foreach (var hotkey in _hotkeys) {
                RegisterHotKey(_hWnd, hotkey.Id, hotkey.Modifier, (uint)hotkey.VirtualKey);
            }
        }
        private void ClearHotKey() {
            foreach (var hotkey in _hotkeys) {
                UnregisterHotKey(_hWnd, hotkey.Id);
            }
            _hotkeys.Clear();
        }
        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
            if (msg == WmHotKey) {
                int id = wParam.ToInt32();

                var hotkey = _hotkeys.FirstOrDefault(h => h.Id == id);
                hotkey?.Action.Invoke(); // ホットキーに設定されたアクションを実行
                handled = true;
            }
            return IntPtr.Zero;
        }

        // イベントハンドラ
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void OcrButton_Click(object sender, RoutedEventArgs e) { Capture(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }


    }
}
