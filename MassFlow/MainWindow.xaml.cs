using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using Tools.Common;
using Tools.Common.InstList;
using static Tools.Common.Win32Wrapper;

namespace MassFlow {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        private IntPtr _hWnd = IntPtr.Zero;

        private readonly InstClass _instDcs;
        private readonly InstClass _instDmm;
        private readonly InstClass _instFg01;
        private readonly InstClass _instFg02_1;
        private readonly InstClass _instFg02_2;
        private readonly InstClass _instOsc;

        public ObservableCollection<string> DcsList { get; } = [];
        public ObservableCollection<string> DmmList { get; } = [];
        public ObservableCollection<string> Fg01List { get; } = [];
        public ObservableCollection<string> Fg02_1List { get; } = [];
        public ObservableCollection<string> Fg02_2List { get; } = [];
        public ObservableCollection<string> OscList { get; } = [];

        public MainWindow() {
            InitializeComponent();
            _instDcs = new();
            _instDmm = new();
            _instDmm = new();
            _instFg01 = new();
            _instFg02_1 = new();
            _instFg02_2 = new();
            _instOsc = new();
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

        internal DataTable _dataTable = new();

        private Dictionary<int, string> _dicSwitchDcs = [];
        private Dictionary<int, (string fg01, string fg02, string fg03_1, string fg03_2)> _dicSwitchFg = [];
        private Dictionary<int, (string fg01, string fg02, string fg03_1, string fg03_2)> _dicSwitchRFg = [];
        private Dictionary<int, string> _dicSwitchOsc = [];
        private Dictionary<int, string> _dicSwitchROsc = [];
        private readonly List<Hotkey> _hotkeys = [];
        private HwndSource? _source;

        private volatile bool _isProcessing = false;

        private static readonly SemaphoreSlim s_semaphore = new(1, 1); // 最大1つの接続

        // FMRemoteのメニューアイテムID
        private const int MenuItemIdA1L = 32806;
        private const int MenuItemIdA1H = 32807;
        private const int MenuItemIdA2L = 32808;
        private const int MenuItemIdA2H = 32809;
        private const int MenuItemIdAnalogTrim = 32815;

        // 起動時
        private void LoadEvents() {
            InstListImport();
            FormatSet();
            RegDictionary();
        }
        private void InstListImport() {
            const string XmlFilePath = "VisaAddress.xml";
            if (!File.Exists(XmlFilePath)) {
                MessageBox.Show($"{XmlFilePath}が見つかりません。");
                return;
            }

            using DataSet dataSet = new();
            dataSet.ReadXml("VisaAddress.xml");
            _dataTable = dataSet.Tables[0];

            // デジタルマルチメータ、ファンクションジェネレータ、オシロスコープのコンボボックスを更新する
            UpdateComboBox(DcsComboBox, DcsList, "電流電圧発生器", [2, 3], "[DCS]");
            UpdateComboBox(DmmComboBox, DmmList, "デジタルマルチメータ", [1, 2], "[DMM1]");
            UpdateComboBox(Fg01ComboBox, Fg01List, "ファンクションジェネレータ", [3], "[FG]");
            UpdateComboBox(Fg02_1ComboBox, Fg02_1List, "ファンクションジェネレータ", [2], "[FG]");
            UpdateComboBox(Fg02_2ComboBox, Fg02_2List, "ファンクションジェネレータ", [2], "[FG]");
            UpdateComboBox(OscComboBox, OscList, "オシロスコープ", [2], "[OSC]");
        }
        private void UpdateComboBox(ComboBox comboBox, ObservableCollection<string> collection, string category, List<int> signalTypes, string name) {
            if (_dataTable == null) return;
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
            GetVisaAddress(_instDmm, DmmComboBox);
            GetVisaAddress(_instFg01, Fg01ComboBox);
            GetVisaAddress(_instFg02_1, Fg02_1ComboBox);
            GetVisaAddress(_instFg02_2, Fg02_2ComboBox);
            GetVisaAddress(_instOsc, OscComboBox);
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
            // FG
            _dicSwitchFg = new Dictionary<int, (string fg01, string fg02, string fg03_1, string fg03_2)>() {
                {
                    0,(
                    ":CHAN 2;:OUTP:STAT OFF;:CHAN 1;:OUTP:STAT OFF;:VOLT:OFFS 3.0;*WAI;*STB?\r\n",
                    ":OUTP1:STAT OFF;:SOUR1:VOLT:OFFS 3.0;:OUTP2:STAT OFF;*OPC?",
                    ":OUTP OFF;:VOLT:OFFS 3.0;*OPC?",
                    ":OUTP OFF;*OPC?"
                    )
                },
                {
                    1,(
                    ":CHAN 2;:OUTP:STAT ON;:CHAN 1;:OUTP:STAT ON;*WAI;*STB?\r\n",
                    ":OUTP1:STAT ON;:OUTP2:STAT ON;*OPC?",
                    ":OUTP ON;*OPC?",
                    ":OUTP ON;*OPC?"
                    )
                },
                {
                    2,(
                    ":CHAN 1;:OUTP:STAT OFF;:VOLT 5.0;:OUTP:STAT ON;*WAI;*STB?\r\n",
                    ":SOUR1:VOLT:LEV:IMM:AMPL 5.0VPP;*OPC?",
                    ":VOLT 5.0VPP;*OPC?",
                    string.Empty
                    )
                },
                {
                    3,(
                    ":CHAN 1;:OUTP:STAT OFF;:VOLT 0.0;:VOLT:OFFS 4.0;:OUTP:STAT ON;*WAI;*STB?\r\n",
                    ":SOUR1:VOLT:LEV:IMM:AMPL 0.0VPP;:SOUR1:VOLT:OFFS 4.0;*OPC?",
                    ":VOLT 0.0VPP;:VOLT:OFFS 4.0;*OPC?",
                    string.Empty
                    )
                },
                {
                    4,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
                {
                    5,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
                {
                    6,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
            };
            _dicSwitchRFg = new Dictionary<int, (string fg01, string fg02, string fg03_1, string fg03_2)>() {
                {
                    0,(
                    ":CHAN 2;:OUTP:STAT OFF;:CHAN 1;:OUTP:STAT OFF;*WAI;*STB?\r\n",
                    ":OUTP1:STAT OFF;:OUTP2:STAT OFF;*OPC?",
                    ":OUTP OFF;*OPC?",
                    ":OUTP OFF;*OPC?"
                    )
                },
                {
                    1,(
                    ":CHAN 1;:OUTP:STAT OFF;:VOLT 0.0;:OUTP:STAT ON;*WAI;*STB?\r\n",
                    ":SOUR1:VOLT:LEV:IMM:AMPL 0.0VPP;*OPC?",
                    ":VOLT 0.0VPP;*OPC?",
                    string.Empty
                    )
                },
                {
                    2,(
                    ":CHAN 1;:OUTP:STAT OFF;:VOLT 5.0;:VOLT:OFFS 3.0;:OUTP:STAT ON;*WAI;*STB?\r\n",
                    ":SOUR1:VOLT:LEV:IMM:AMPL 5.0VPP;:SOUR1:VOLT:OFFS 3.0;*OPC?",
                    ":VOLT 5.0VPP;:VOLT:OFFS 3.0;*OPC?",
                    string.Empty
                    )
                },
                {
                    3,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
                {
                    4,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
                {
                    5,(
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty
                    )
                },
                {
                    6,(
                    ":CHAN 2;:OUTP:STAT ON;:CHAN 1;:OUTP:STAT OFF;:VOLT:OFFS 4.0;:OUTP:STAT ON;*WAI;*STB?\r\n",
                    ":OUTP1:STAT ON;:OUTP2:STAT ON;:SOUR1:VOLT:OFFS 4.0;*OPC?",
                    ":OUTP ON;:VOLT:OFFS 4.0;*OPC?",
                    ":OUTP ON;*OPC?"
                    )
                },
            };

            // OSC
            _dicSwitchOsc = new Dictionary<int, string> {
                { 0, string.Empty },
                {
                    1,
                    """
                    :CH2:SCALE 5.0E0;
                    :CH3:SCALE 5.0E0;
                    :CH4:SCALE 5.0E0;
                    :HORIZONTAL:MAIN:SCALE 2.5E-3;
                    :TRIGGER:MAIN:LEVEL 2.0E0;
                    *OPC?
                    """
                },
                {
                    2,
                    """
                    :SELECT:CH1 1;CH2 0;CH4 0;
                    :MEASUREMENT:MEAS1:TYPE FREQUENCY;SOURCE CH1;
                    :MEASUREMENT:MEAS2:TYPE PK2PK;SOURCE CH1;
                    *OPC?
                    """
                },
                {
                    3,
                    """
                    :SELECT:CH1 0;CH3 1;
                    :MEASUREMENT:MEAS1:TYPE NONE;SOURCE CH3;
                    :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH3;
                    :HORIZONTAL:MAIN:SCALE 5.0E-2;
                    :TRIGGER:MAIN:EDGE:SOURCE CH3;
                    :TRIGGER:MAIN:LEVEL 1.0E1;
                    *OPC?
                    """
                },
                {
                    4,
                    """
                    :MEASUREMENT:MEAS2:TYPE MINIMUM;SOURCE CH3;
                    :ACQUIRE:MODE AVERAGE;NUMAVG 4;
                    :CH3:SCALE 5.0E-2;
                    :HORIZONTAL:MAIN:SCALE 2.5E-3;
                    :TRIGGER:MAIN:LEVEL 5.0E-2;
                    *OPC?
                    """
                },
                {
                    5,
                    """
                    :MEASUREMENT:MEAS2:TYPE FREQUENCY;SOURCE CH3;
                    :ACQUIRE:MODE SAMPLE;NUMAVG 16;
                    :CH3:SCALE 5.0E0;
                    :HORIZONTAL:MAIN:SCALE 5.0E-2;
                    :TRIGGER:MAIN:LEVEL 1.0E1;
                    *OPC?
                    """
                },
                {
                    6,
                    """
                    :SELECT:CH2 1;CH3 0;CH4 1;
                    :MEASUREMENT:MEAS1:TYPE MAXIMUM;SOURCE CH2;
                    :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH4;
                    :ACQUIRE:MODE SAMPLE;NUMAVG 16;
                    :CH2:SCALE 1.0E0;
                    :CH4:SCALE 1.0E0;
                    :HORIZONTAL:MAIN:SCALE 1.0E-2;
                    :TRIGGER:MAIN:EDGE:SOURCE CH1;
                    :TRIGGER:MAIN:LEVEL 0.0E0;
                    *OPC?
                    """
                }
            };
            _dicSwitchROsc = new Dictionary<int, string> {
                { 0, string.Empty },
                {
                    1,
                    """
                    :SELECT:CH1 0;CH2 1;CH3 0;CH4 1;
                    :MEASUREMENT:MEAS1:TYPE MAXIMUM;SOURCE CH2;
                    :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH4;
                    *OPC?
                    """
                },
                {
                    2,
                    """
                    :SELECT:CH1 1;CH3 0;CH2 0;CH4 0;
                    :MEASUREMENT:MEAS1:TYPE FREQUENCY;SOURCE CH1;
                    :MEASUREMENT:MEAS2:TYPE PK2PK;SOURCE CH1;
                    :CH2:SCALE 5.0E0;
                    :CH4:SCALE 5.0E0;
                    :HORIZONTAL:MAIN:SCALE 2.5E-3;
                    :TRIGGER:MAIN:EDGE:SOURCE CH1;
                    :TRIGGER:MAIN:LEVEL 2.0E0;
                    *OPC?
                    """
                },
                {
                    3,
                    """
                    :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH3;
                    :ACQUIRE:MODE SAMPLE;NUMAVG 16;
                    :CH3:SCALE 5.0E0;
                    :HORIZONTAL:MAIN:SCALE 5.0E-2;
                    :TRIGGER:MAIN:LEVEL 1.0E1;
                    *OPC?
                    """
                },
                {
                    4,
                    """
                    :MEASUREMENT:MEAS2:TYPE MINIMUM;SOURCE CH3;
                    :ACQUIRE:MODE AVERAGE;NUMAVG 4;
                    :CH3:SCALE 5.0E-2;
                    :HORIZONTAL:MAIN:SCALE 2.5E-3;
                    :TRIGGER:MAIN:LEVEL 5.0E-2;
                    *OPC?
                    """
                },
                {
                    5,
                    """
                    :SELECT:CH1 0;CH2 0;CH3 1;CH4 0;
                    :MEASUREMENT:MEAS1:TYPE NONE;SOURCE CH3;
                    :MEASUREMENT:MEAS2:TYPE FREQUENCY;SOURCE CH3;
                    :CH3:SCALE 5.0E0;:HORIZONTAL:MAIN:SCALE 5.0E-2;
                    :TRIGGER:MAIN:EDGE:SOURCE CH3;
                    :TRIGGER:MAIN:LEVEL 1.0E1;
                    *OPC?
                    """
                },
                {
                    6,
                    """
                    :CH2:SCALE 1.0E0;
                    :CH3:SCALE 5.0E-2;
                    :CH4:SCALE 1.0E0;
                    :HORIZONTAL:MAIN:SCALE 1.0E-2;
                    :TRIGGER:MAIN:LEVEL 0.0E0;
                    *OPC?
                    """
                }
            };
            //_dicSwitchDcs = new Dictionary<int, (string cmd2, string cmd3, string text)>
            //{
            //    { 0, ("SOI+0MA,SBY", "F5R6S0EO0EOC", "OFF") },
            //    { 1, ("SOI+4MA,OPR", "F5R6S4.0E-3O1EOC", "4.0mA") },
            //    { 2, ("SOI+20MA,OPR", "F5R6S20.0E-3O1EOC", "20mA") },
            //    { 3, ("SOI+4MA,OPR", "F5R6S4.0E-3O1EOC", "4.0mA") },
            //    { 4, ("SOI+20MA,OPR", "F5R6S20.0E-3O1EOC", "20mA") },
            //    { 5, ("SOI+0MA,SBY", "F5R6S0EO0EOC", "OFF") },
            //    { 6, ("SOI+22MA,OPR", "F5R6S22.0E-3O1EOC", "22mA") },
            //    { 7, ("SOI+20MA,OPR", "F5R6S20.0E-3O1EOC", "20mA") },
            //    { 8, ("SOI+12MA,OPR", "F5R6S12.0E-3O1EOC", "12mA") },
            //    { 9, ("SOI+4MA,OPR", "F5R6S4.0E-3O1EOC", "4.0mA") },
            //    { 10, ("SOI+3.2MA,OPR", "F5R6S3.2E-3O1EOC", "3.2mA") },
            //    { 11, ("SOI+22MA,OPR", "F5R6S22.0E-3O1EOC", "22mA") },
            //    { 12, ("SOI+20MA,OPR", "F5R6S20.0E-3O1EOC", "20mA") },
            //    { 13, ("SOI+12MA,OPR", "F5R6S12.0E-3O1EOC", "12mA") },
            //    { 14, ("SOI+4MA,OPR", "F5R6S4.0E-3O1EOC", "4.0mA") },
            //    { 15, ("SOI+3.2MA,OPR", "F5R6S3.2E-3O1EOC", "3.2mA") },
            //};
        }
        // 機器初期設定
        private void FormatSet() {
            // DCS
            _instDcs.InstCommand = _instDcs.SignalType switch {
                2 => "SVR5,SOV+0,SBY,*OPC?",
                _ => string.Empty,
            };
            // DMM
            _instDmm.InstCommand = _instDmm.SignalType switch {
                1 => "*RST,F5,R6,*OPC?",
                2 => "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?",
                _ => string.Empty,
            };
            // FG
            _instFg01.InstCommand = _instFg01.SignalType switch {
                3 => "*RST;:CHAN 2;:MODE NORM;:FUNC:SHAP FSQU;:FREQ 100;:VOLT 0.0;:VOLT:OFFS 2.0;:OUTP:STAT OFF;:CHAN 1;:CHAN:MODE IND;:MODE NORM;:FUNC:SHAP FSQU;:FREQ 100;:VOLT 0.0;:VOLT:OFFS 3.0;:OUTP:STAT OFF;*WAI;*STB?\r\n",
                4 => "*RST;:CHAN:MODE IND;:SOUR1:FUNC:SHAP SQU;SQU:DCYC 50PCT;:SOUR1:VOLT:LEV:IMM:AMPL 0.0VPP;:SOUR1:FREQ:CW 100HZ;:SOUR1:VOLT:OFFS 3.0;:SOUR2:FUNC:SHAP SQU;SQU:DCYC 50PCT;:SOUR2:VOLT:LEV:IMM:AMPL 0.0VPP;:SOUR2:FREQ:CW 100HZ;:SOUR2:VOLT:OFFS 2.0;:OUTP1:STAT OFF;:OUTP2:STAT OFF;*OPC?",
                _ => string.Empty,
            };
            _instFg02_1.InstCommand = _instFg02_1.SignalType switch {
                2 => "*RST;:FUNC SQU;:FREQ 100;:VOLT 0.0VPP;:VOLT:OFFS 3.0;*OPC?",
                _ => string.Empty,
            };
            _instFg02_2.InstCommand = _instFg02_2.SignalType switch {
                2 => "*RST;:FUNC SQU;:FREQ 100;:VOLT 0.0VPP;:VOLT:OFFS 2.0;*OPC?",
                _ => string.Empty,
            };
            // OSC
            _instOsc.InstCommand = _instOsc.SignalType switch {
                2 =>
                """
                *RST;
                :HEADER 0;
                :ACQUIRE:STATE 1;STOPAFTER RUNSTOP;
                :SELECT:CH1 0;CH2 1;CH4 1;
                :CH1:PROBE 1.0E1;CURRENTPROBE 1.0E1;SCALE 1.0E0;POSITION -2.0E0;
                :CH2:PROBE 1.0E1;CURRENTPROBE 1.0E1;SCALE 5.0E0;POSITION -2.0E0;
                :CH3:PROBE 1.0E1;CURRENTPROBE 1.0E1;SCALE 5.0E0;POSITION -3.0E0;
                :CH4:PROBE 1.0E1;CURRENTPROBE 1.0E1;SCALE 5.0E0;POSITION -3.0E0;
                :HORIZONTAL:MAIN:SCALE 2.5E-3;
                :HORIZONTAL:DELAY:SCALE 5.0E-9;POSITION 5.0E-4;
                :HORIZONTAL:SCALE 2.5E-3;
                :TRIGGER:MAIN:HOLDOFF:VALUE 5.0E-7;
                :TRIGGER:MAIN:VIDEO:SYNC FIELD;LINE 1;
                :TRIGGER:MAIN:LEVEL 2.0E0;
                :MEASUREMENT:MEAS1:TYPE MAXIMUM;SOURCE CH2;
                :MEASUREMENT:MEAS2:TYPE MAXIMUM;SOURCE CH4;
                :MEASUREMENT:MEAS3:TYPE NONE;SOURCE MATH;
                :MEASUREMENT:MEAS4:TYPE NONE;SOURCE MATH;
                :MEASUREMENT:MEAS5:TYPE NONE;SOURCE MATH;
                *OPC?
                """,
                _ => string.Empty,
            };
        }

        // 機器接続
        private async void ConnectInstAsync() {
            try {
                HotKeyChekBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();
                CheckFgId();
                FormatSet();

                var value = int.Parse(FgNumberComboBox.Text);

                InstClass[] devices = value switch {
                    1 => [_instDcs, _instDmm, _instFg01, _instOsc],
                    2 => [_instDcs, _instDmm, _instFg02_1, _instFg02_2, _instOsc],
                    _ => [_instDcs, _instDmm, _instOsc] // 1, 2 以外の値の場合のデフォルト
                };

                var tasks = devices.Select(device => ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);

                //if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                //    DcsNumberLabel.Text = "00";
                //    DcsRangeLabel.Text = "OFF";
                //}
                //if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) { OscRangeLabel.Text = "50m"; }

                DcsComboBox.IsEnabled = false;
                DmmComboBox.IsEnabled = false;
                Fg01ComboBox.IsEnabled = false;
                Fg02_1ComboBox.IsEnabled = false;
                Fg02_2ComboBox.IsEnabled = false;
                OscComboBox.IsEnabled = false;
                ConnectButton.IsEnabled = false;
                ReleaseButton.IsEnabled = true;
                InstListButton.IsEnabled = false;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー");
            } finally {
                VisibleProgressImage(false);
            }
        }
        // DMMのIDチェック処理
        private void CheckFgId() {
            var indices = new[] { _instFg02_1.Index, _instFg02_2.Index }
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
                    2 or 3 or 4 => await ConnectDeviceVisaAsync(instClass),
                    _ => throw new ApplicationException(),
                };
        }
        // Visa接続
        private static async Task<string> ConnectDeviceVisaAsync(InstClass instClass) {
            return await Task.Run(() => {
                using var usbDev = new USBDeviceManager();
                usbDev.OpenDev(instClass.VisaAddress);
                usbDev.OutputDev(instClass.InstCommand);
                return usbDev.InputDev();
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
            _instDmm.ResetProperties();
            _instFg01.ResetProperties();
            _instFg02_1.ResetProperties();
            _instFg02_2.ResetProperties();
            _instOsc.ResetProperties();

            DcsComboBox.IsEnabled = true;
            DmmComboBox.IsEnabled = true;
            Fg01ComboBox.IsEnabled = true;
            Fg02_1ComboBox.IsEnabled = true;
            Fg02_2ComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            InstListButton.IsEnabled = true;
            HotKeyChekBox.IsChecked = false;
        }





















        // HotkKeyの登録
        private void SetHotKey() {
            _hotkeys.Clear();
            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                _hotkeys.AddRange([
                    //new(ModNone, HotkeyColon, ActionHotkeyColon),
                    //new(ModNone, HotkeyNumMultiply, ActionHotkeyNumMultiply),
                    ]);
            }
            if (!string.IsNullOrEmpty(_instFg01.VisaAddress)) {
                _hotkeys.AddRange([
                    //new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                    //    new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
                    ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                _hotkeys.AddRange([
                    //new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                    //    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    //    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    ]);
            }
            if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                _hotkeys.AddRange([
                    //new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                    //    new(ModShift, HotkeyAtsign, ActionHotkeyShiftAtsign),
                    //    new(ModNone, HotkeyNumDivide, ActionHotkeyNumDivide),
                    ]);
            }


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

        // 特定の親ウィンドウに属するすべての子ウィンドウのハンドルを取得するメソッド
        private static List<IntPtr> FindChildWindows(IntPtr hWndParent) {
            List<IntPtr> childWindows = [];
            EnumChildProc callback = new((hwnd, param) => {
                childWindows.Add(hwnd);
                return true;
            });

            EnumChildWindows(hWndParent, callback, IntPtr.Zero);
            return childWindows;
        }





        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void InstListButton_Click(object sender, RoutedEventArgs e) { ShowInstList(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }
        private void TopMostCheckBox_Checked(object sender, RoutedEventArgs e) { Topmost = true; }
        private void TopMostCheckBox_Unchecked(object sender, RoutedEventArgs e) { Topmost = false; }


    }
}