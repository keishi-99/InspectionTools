using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using Tools.Common;
using Tools.Common.InstList;
using static Tools.Common.Win32Wrapper;

namespace InstrumentCheck {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 

    public class MeasurementRow : INotifyPropertyChanged {
        private string? _dmm1;
        private string? _dmm2;
        private string? _dmm3;

        public string Condition { get; set; } = string.Empty;

        public string? Dmm1 {
            get => _dmm1;
            set { _dmm1 = value; OnPropertyChanged(nameof(Dmm1)); }
        }

        public string? Dmm2 {
            get => _dmm2;
            set { _dmm2 = value; OnPropertyChanged(nameof(Dmm2)); }
        }

        public string? Dmm3 {
            get => _dmm3;
            set { _dmm3 = value; OnPropertyChanged(nameof(Dmm3)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public partial class MainWindow : Window {

        private IntPtr _hWnd = IntPtr.Zero;

        private readonly InstClass _instDcs;
        private readonly InstClass _instDmm01;
        private readonly InstClass _instDmm02;
        private readonly InstClass _instDmm03;

        public ObservableCollection<string> DcsList { get; } = [];
        public ObservableCollection<string> Dmm1List { get; } = [];
        public ObservableCollection<string> Dmm2List { get; } = [];
        public ObservableCollection<string> Dmm3List { get; } = [];

        public MainWindow() {
            InitializeComponent();
            _instDcs = new();
            _instDmm01 = new();
            _instDmm02 = new();
            _instDmm03 = new();
            LoadEvents();
            // Window が完全に作られたあとにハンドルを取得
            Loaded += (s, e) => { _hWnd = new WindowInteropHelper(this).Handle; };
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

        private Dictionary<int, (string cmd, string text)> _dicSwitchDcs = [];

        internal DataTable _dataTable = new();

        private readonly List<Hotkey> _hotkeys = [];
        private HwndSource? _source;

        private volatile bool _isProcessing = false;

        private static readonly SemaphoreSlim s_semaphore = new(1, 1); // 最大1つの接続

        private List<MeasurementRow> _data = [];

        // プロパティ名を配列で管理
        private readonly string[] _propertyNames = ["Dmm1", "Dmm2", "Dmm3"];
        private readonly string[] _dmmDataCollection = ["4mA", "12mA", "20mA", "1V", "3V", "5V"];

        // 起動時
        private void LoadEvents() {
            InstListImport();
            FormatSet();
            RegDictionary();

            _data = [.. _dmmDataCollection.Select(c => new MeasurementRow { Condition = c })];
            ResultDataGrid.ItemsSource = _data;

            // RowHeader に Condition を表示
            ResultDataGrid.LoadingRow += (s, e) => {
                if (e.Row.Item is MeasurementRow row)
                    e.Row.Header = row.Condition;
            };

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
            UpdateComboBox(DcsComboBox, DcsList, "電流電圧発生器", [2], "[DCS]");
            UpdateComboBox(Dmm01ComboBox, Dmm1List, "デジタルマルチメータ", [1, 2], "[DMM-1]");
            UpdateComboBox(Dmm02ComboBox, Dmm2List, "デジタルマルチメータ", [1, 2], "[DMM-2]");
            UpdateComboBox(Dmm03ComboBox, Dmm3List, "デジタルマルチメータ", [1, 2], "[DMM-3]");
        }
        private void UpdateComboBox(ComboBox comboBox, ObservableCollection<string> collection, string category, List<int> signalTypes, string name) {
            if (_dataTable == null) return;
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
            GetVisaAddress(_instDcs, DcsComboBox);
            GetVisaAddress(_instDmm01, Dmm01ComboBox);
            GetVisaAddress(_instDmm02, Dmm02ComboBox);
            GetVisaAddress(_instDmm03, Dmm03ComboBox);
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
        // 機器リスト表示
        private void ShowInstList() {
            InstListWindow frm1 = new(_dataTable);
            frm1.ShowDialog();
            InstListImport();
        }
        // 機器設定辞書登録
        private void RegDictionary() {
            _dicSwitchDcs = new Dictionary<int, (string cmd, string text)>
            {
                { 0, ("SIR2,SOI+0MA,SBY","STBY") },
                { 1, ("SIR2,SOI+4MA,OPR", "4mA") },
                { 2, ("SIR2,SOI+12MA,OPR", "12mA") },
                { 3, ("SIR2,SOI+20MA,OPR", "20mA") },
                { 4, ("SVR5,SOV+0,SBY", "STBY") },
                { 5, ("SVR5,SOV+1,OPR", "1V") },
                { 6, ("SVR5,SOV+3,OPR", "3V") },
                { 7, ("SVR5,SOV+5,OPR", "5V") },
            };
        }
        // 機器初期設定
        private void FormatSet() {
            _instDcs.InstCommand = _instDcs.SignalType switch {
                2 => "*RST,SVR5,SOV+0,SBY,*OPC?",
                _ => string.Empty,
            };
            _instDmm01.InstCommand = _instDmm01.SignalType switch {
                1 => "*RST,F5",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?",
                _ => string.Empty,
            };
            _instDmm02.InstCommand = _instDmm02.SignalType switch {
                1 => "*RST,F5",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?",
                _ => string.Empty,
            };
            _instDmm03.InstCommand = _instDmm03.SignalType switch {
                1 => "*RST,F5",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?",
                _ => string.Empty,
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {
            try {
                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                CheckDmmId();
                FormatSet();

                var devices = new[] { _instDcs, _instDmm01, _instDmm02, _instDmm03 };
                var tasks = devices.Select(device => ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);


                DcsComboBox.IsEnabled = false;
                Dmm01ComboBox.IsEnabled = false;
                Dmm02ComboBox.IsEnabled = false;
                Dmm03ComboBox.IsEnabled = false;
                ConnectButton.IsEnabled = false;
                ReleaseButton.IsEnabled = true;
                InstListButton.IsEnabled = false;

                if (string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                    Dcs4mAButton.IsEnabled = true;
                    Dcs12mAButton.IsEnabled = true;
                    Dcs20mAButton.IsEnabled = true;
                    Dcs1VButton.IsEnabled = true;
                    Dcs3VButton.IsEnabled = true;
                    Dcs5VButton.IsEnabled = true;
                    DcsStanbyButton.IsEnabled = true;
                    AutoButton.IsEnabled = true;
                }

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー");
            } finally {
                VisibleProgressImage(false);
            }
        }
        // DMMのIDチェック処理
        private void CheckDmmId() {
            var indices = new[] { _instDmm01.Index, _instDmm02.Index, _instDmm03.Index }
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

            _instDcs.ResetProperties();
            _instDmm01.ResetProperties();
            _instDmm02.ResetProperties();
            _instDmm03.ResetProperties();

            DcsComboBox.IsEnabled = true;
            Dmm01ComboBox.IsEnabled = true;
            Dmm02ComboBox.IsEnabled = true;
            Dmm03ComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            InstListButton.IsEnabled = true;
            HotKeyChekBox.IsChecked = false;

            Dcs4mAButton.IsEnabled = false;
            Dcs12mAButton.IsEnabled = false;
            Dcs20mAButton.IsEnabled = false;
            Dcs1VButton.IsEnabled = false;
            Dcs3VButton.IsEnabled = false;
            Dcs5VButton.IsEnabled = false;
            DcsStanbyButton.IsEnabled = false;
            AutoButton.IsEnabled = false;
        }

        // DMM測定値取得
        private static async Task<decimal> ReadDmm(InstClass instClass) {
            instClass.InstCommand = instClass.SignalType switch {
                1 => string.Empty,
                2 => "FETC?",
                _ => throw new ApplicationException(),
            };

            var result = await ConnectDeviceAsync(instClass);
            decimal.TryParse(result, NumberStyles.AllowExponent | NumberStyles.Float, CultureInfo.InvariantCulture, out var output);

            return output;
        }

        // DCS切り替え
        private async Task SwitchDcsAsync(InstClass instClass, int count) {
            try {
                if (string.IsNullOrEmpty(instClass.VisaAddress)) { return; }

                if (count == -1) {
                    instClass.SettingNumber += 1;
                    if (instClass.SettingNumber > 7) { instClass.SettingNumber = 0; }
                    ;
                }
                else { _instDcs.SettingNumber = count; }

                instClass.InstCommand = _dicSwitchDcs[instClass.SettingNumber].cmd;

                if (string.IsNullOrEmpty(instClass.InstCommand)) { return; }
                await ConnectDeviceAsync(instClass);
                DcsRangeTextBox.Text = _dicSwitchDcs[instClass.SettingNumber].text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // DMM切り替え
        private async Task SwitchDmmAsync(InstClass instClass, string func) {
            try {
                instClass.InstCommand = func switch {
                    "DCI" => instClass.SignalType switch {
                        1 => "*RST,F5,*OPC?",
                        2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?",
                        _ => throw new ApplicationException(),
                    },
                    "DCV" => instClass.SignalType switch {
                        1 => "*RST,*OPC?",
                        2 => "*RST;:INIT:CONT 1;*OPC?",
                        _ => throw new ApplicationException(),
                    },
                    _ => throw new ApplicationException(),
                };
                await ConnectDeviceAsync(instClass);

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 自動計測を行うメインメソッド
        private async Task AutoInstrumentation() {
            try {
                // フォームを無効化する
                VisibleProgressImage(true);
                ProgressBar.Value = 0;

                // 待機時間を取得し、必要であれば最小値を設定する
                var waitTime = int.Parse(WaitTimeTextBox.Text);

                // 各DMMの設定を配列に格納する
                InstClass[] instDmm = [_instDmm01, _instDmm02, _instDmm03];

                // ヘッダーテキストを設定する
                SetColumnHeaders(instDmm);

                // 進捗バー用: 測定回数の合計（各DMMで6回ずつ）
                int totalSteps = instDmm.Count(d => !string.IsNullOrEmpty(d.VisaAddress)) * 6;
                int currentStep = 0;

                // 各DMMに対して計測を行う
                for (var i = 0; i < ResultDataGrid.Columns.Count && i < instDmm.Length; i++) {
                    if (!string.IsNullOrEmpty(instDmm[i].VisaAddress)) {
                        int step = await MeasureDmmAsync(instDmm[i], i, waitTime, totalSteps, currentStep);
                        currentStep += step;
                    }
                }
                // 計測終了処理を行う
                await FinalizeMeasurement();

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
                ProgressBar.Value = 100;
            }
        }
        // ヘッダーテキストを設定する
        private void SetColumnHeaders(InstClass[] instDmm) {
            for (var i = 0; i < ResultDataGrid.Columns.Count && i < instDmm.Length; i++) {
                ResultDataGrid.Columns[i].Header = instDmm[i].Name;
            }
        }
        // 各DMMに対して計測を行う
        private async Task<int> MeasureDmmAsync(InstClass instClassDmm, int i, int waitTime, int totalSteps, int currentStep) {
            int stepCount = 0;
            await SwitchDmmAsync(instClassDmm, "DCI");
            await SwitchDcsAsync(_instDcs, 0);

            Activate();
            var result = MessageBox.Show(this, $"{instClassDmm.Name} DCI測定します。", "確認", MessageBoxButton.OKCancel);
            if (result == MessageBoxResult.Cancel) { return stepCount; }

            for (var j = 1; j <= 3; j++) {
                await SwitchDcsAsync(_instDcs, j);
                await Task.Delay(waitTime);
                await MeasureAndSetCellValue(i, j - 1, instClassDmm);
                stepCount++;
                ProgressBar.Value = Math.Min(100, (int)((double)(currentStep + stepCount) / totalSteps * 100));
            }

            await SwitchDcsAsync(_instDcs, 4);
            await SwitchDmmAsync(instClassDmm, "DCV");

            Activate();
            result = MessageBox.Show(this, $"{instClassDmm.Name} DCV測定します。", "確認", MessageBoxButton.OKCancel);
            if (result == MessageBoxResult.Cancel) { return stepCount; }

            for (var j = 5; j <= 7; j++) {
                await SwitchDcsAsync(_instDcs, j);
                await Task.Delay(waitTime);
                await MeasureAndSetCellValue(i, j - 2, instClassDmm);
                stepCount++;
                ProgressBar.Value = Math.Min(100, (int)((double)(currentStep + stepCount) / totalSteps * 100));
            }
            return stepCount;
        }
        // DMMの計測値を取得し、セルに設定する
        private async Task MeasureAndSetCellValue(int i, int j, InstClass instClass) {
            var mesValue = await ReadDmm(instClass);
            var value = j switch {
                0 or 1 or 2 => (mesValue * 1000).ToString("0.000"),
                3 => mesValue.ToString("0.00000"),
                4 or 5 => mesValue.ToString("0.0000"),
                _ => string.Empty,
            };

            var row = _data.FirstOrDefault(r => r.Condition == _dmmDataCollection[j]);
            if (row == null) return;

            switch (_propertyNames[i]) {
                case "Dmm1": row.Dmm1 = value; break;
                case "Dmm2": row.Dmm2 = value; break;
                case "Dmm3": row.Dmm3 = value; break;
            }
        }
        // 計測終了処理を行う
        private async Task FinalizeMeasurement() {
            await SwitchDcsAsync(_instDcs, 0);
            Activate();
            MessageBox.Show("終了");
        }

        // HotkKeyの登録
        private void SetHotKey() {
            _hotkeys.Clear();

            var helper = new WindowInteropHelper(this);
            _source = HwndSource.FromHwnd(helper.Handle);
            _source.AddHook(HwndHook);

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

        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void InstListButton_Click(object sender, RoutedEventArgs e) { ShowInstList(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }
        private void TopMostCheckBox_Checked(object sender, RoutedEventArgs e) { Topmost = true; }
        private void TopMostCheckBox_Unchecked(object sender, RoutedEventArgs e) { Topmost = false; }

        private async void AutoButton_Click(object sender, RoutedEventArgs e) { await AutoInstrumentation(); }
        private async void DcsStanbyButton_Click(object sender, RoutedEventArgs e) { await SwitchDcsAsync(_instDcs, 0); }
        private async void Dcs4mAButton_Click(object sender, RoutedEventArgs e) { await SwitchDcsAsync(_instDcs, 1); }
        private async void Dcs12mAButton_Click(object sender, RoutedEventArgs e) { await SwitchDcsAsync(_instDcs, 2); }
        private async void Dcs20mAButton_Click(object sender, RoutedEventArgs e) { await SwitchDcsAsync(_instDcs, 3); }
        private async void Dcs1VButton_Click(object sender, RoutedEventArgs e) { await SwitchDcsAsync(_instDcs, 5); }
        private async void Dcs3VButton_Click(object sender, RoutedEventArgs e) { await SwitchDcsAsync(_instDcs, 6); }
        private async void Dcs5VButton_Click(object sender, RoutedEventArgs e) { await SwitchDcsAsync(_instDcs, 7); }
    }
}