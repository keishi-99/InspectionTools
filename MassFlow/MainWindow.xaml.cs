using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using Tools.Common;
using Tools.Common.InstList;
using WindowsInput;
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
            public string Tag { get; set; } = string.Empty;
            public string VisaAddress { get; set; } = string.Empty;
            public int SignalType { get; set; } = 0;
            public int Index { get; set; } = 0;
            public string InstCommand { get; set; } = string.Empty;
            public int SettingNumber { get; set; } = 0;

            public void ResetProperties() {
                UsbDev = new();
                Name = string.Empty;
                Tag = string.Empty;
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

        // ラジオボタンリスト
        private List<RadioButton> DcsRadioButtonsList => [
            DcsOffRadioButton, Dcs2VRadioButton, Dcs8VRadioButton, Dcs1VRadioButton, Dcs7VRadioButton
        ];
        private List<RadioButton> FgOscRadioButtonsList => [
            FgOscRange0RadioButton, FgOscRange1RadioButton, FgOscRange2RadioButton, FgOscRange3RadioButton,
            FgOscRange4RadioButton, FgOscRange5RadioButton,FgOscRange6RadioButton
        ];

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
            UpdateComboBox(DmmComboBox, DmmList, "デジタルマルチメータ", [1, 2], "[DMM]");
            UpdateComboBox(Fg01ComboBox, Fg01List, "ファンクションジェネレータ", [3, 4], "[FG]");
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

                if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                    DcsOffRadioButton.IsChecked = true;
                }
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress) || !string.IsNullOrEmpty(_instFg01.VisaAddress) || !string.IsNullOrEmpty(_instFg02_1.VisaAddress) || !string.IsNullOrEmpty(_instFg02_2.VisaAddress)) {
                    FgOscRotationButton.IsEnabled = true;
                    FgOscRange0RadioButton.IsChecked = true;
                }

                DcsComboBox.IsEnabled = false;
                DmmComboBox.IsEnabled = false;
                FgNumberComboBox.IsEnabled = false;
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
        // FGのIDチェック処理
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
            FgNumberComboBox.IsEnabled = true;
            Fg01ComboBox.IsEnabled = true;
            Fg02_1ComboBox.IsEnabled = true;
            Fg02_2ComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            InstListButton.IsEnabled = true;
            HotKeyChekBox.IsChecked = false;
        }

        // DCS切り替え
        private async void SwitchDcs(int i) {
            VisibleProgressImage(true);
            try {
                if (_instFg01.SettingNumber != 0 || _instFg02_1.SettingNumber != 0 || _instOsc.SettingNumber != 0) {
                    MessageBox.Show("FG&OSCがONになっています。");
                    return;
                }

                var hWnd = IntPtr.Zero;
                if (i != 0) {
                    hWnd = FindWindow(null, "マルチ流量計渦 [V01.08]");
                    SetForegroundWindow(hWnd);
                }

                await SwitchDcsAsync(_instDcs, i);

                // 非同期処理完了後、UIスレッドでPostMessageを送信

                var menuItemID = i switch {
                    0 => 0,
                    1 => MenuItemIdA1L,
                    2 => MenuItemIdA1H,
                    3 => MenuItemIdA2L,
                    4 => MenuItemIdA2H,
                    _ => throw new ArgumentOutOfRangeException(nameof(i), "無効なメニューアイテムID")
                };
                if (hWnd != IntPtr.Zero) {
                    _ = PostMessage(hWnd, WmCommand, menuItemID, 0);
                }
            } catch (Exception ex) {
                Release();
                Activate();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        private async Task SwitchDcsAsync(InstClass instClass, int i) {
            if (string.IsNullOrEmpty(instClass.VisaAddress)) { return; }

            instClass.SettingNumber = i;
            instClass.InstCommand = instClass.SettingNumber switch {
                0 => "SVR5,SOV+0,SBY",
                1 => "SVR5,SOV+2,OPR",
                2 => "SVR5,SOV+8,OPR",
                3 => "SVR5,SOV+1,OPR",
                4 => "SVR5,SOV+7,OPR",
                _ => string.Empty
            };

            if (string.IsNullOrEmpty(instClass.InstCommand) || instClass.UsbDev is null) { return; }
            await ConnectDeviceAsync(instClass);

            // 対応するラジオボタンを選択
            DcsRadioButtonsList[instClass.SettingNumber].IsChecked = true;

        }

        // FG&OSCローテーション
        private async void RotationFgOsc(bool isNext) {
            try {
                VisibleProgressImage(true);

                if (_instDcs.SettingNumber != 0) {
                    MessageBox.Show("DCSがONになっています。", "", MessageBoxButton.OK, MessageBoxImage.Error);
                    Activate();
                    return;
                }

                var value = int.Parse(FgNumberComboBox.Text);
                await (value switch {
                    1 => Task.Run(() => RotationFgOscAsync(_instFg01, _instOsc, isNext)),
                    2 => Task.Run(() => RotationFgOscAsync(_instFg02_1, _instFg02_2, _instOsc, isNext)),
                    _ => Task.CompletedTask // どれにも該当しない場合はすぐに完了するタスク
                });

                // 対応するラジオボタンを選択
                FgOscRadioButtonsList[_instOsc.SettingNumber].IsChecked = true;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                Activate();
            } finally {
                VisibleProgressImage(false);
            }
        }
        private async Task RotationFgOscAsync(InstClass instFg, InstClass instOsc, bool isNext) {
            var fgMaxSettingNumber = _dicSwitchFg.Count;
            var oscMaxSettingNumber = _dicSwitchOsc.Count;

            instFg.SettingNumber = (instFg.SettingNumber + (isNext ? 1 : -1) + fgMaxSettingNumber) % fgMaxSettingNumber;
            instOsc.SettingNumber = (instOsc.SettingNumber + (isNext ? 1 : -1) + oscMaxSettingNumber) % oscMaxSettingNumber;

            await ConnectAndSendCommand(instFg, GetFgCommand(instFg, isNext));
            await ConnectAndSendCommand(instOsc, GetOscCommand(instOsc, isNext));
        }
        private async Task RotationFgOscAsync(InstClass instFg2_1, InstClass instFg2_2, InstClass instOsc, bool isNext) {
            var fgMaxSettingNumber = _dicSwitchFg.Count;
            var oscMaxSettingNumber = _dicSwitchOsc.Count;

            instFg2_1.SettingNumber = (instFg2_1.SettingNumber + (isNext ? 1 : -1) + fgMaxSettingNumber) % fgMaxSettingNumber;
            instOsc.SettingNumber = (instOsc.SettingNumber + (isNext ? 1 : -1) + oscMaxSettingNumber) % oscMaxSettingNumber;

            await ConnectAndSendCommand(instFg2_1, GetFgCommand(instFg2_1, isNext));
            await ConnectAndSendCommand(instFg2_2, GetFgCommand(instFg2_2, isNext));
            await ConnectAndSendCommand(instOsc, GetOscCommand(instOsc, isNext));
        }
        private string GetFgCommand(InstClass instFg, bool isNext) {
            var dic = isNext ? _dicSwitchFg : _dicSwitchRFg;
            var (fg01, fg02, fg03_1, fg03_2) = dic[instFg.SettingNumber];

            return instFg.SignalType switch {
                2 => instFg.Tag switch {
                    "FG-1" => fg03_1,
                    "FG-2" => fg03_2,
                    _ => string.Empty
                },
                3 => fg01,
                4 => fg02,
                _ => string.Empty
            };
        }
        private string GetOscCommand(InstClass instOsc, bool isNext) {
            var dic = isNext ? _dicSwitchOsc : _dicSwitchROsc;
            return dic[instOsc.SettingNumber];
        }
        private static async Task ConnectAndSendCommand(InstClass instClass, string command) {
            if (!string.IsNullOrEmpty(instClass.VisaAddress) && !string.IsNullOrEmpty(command) && instClass.UsbDev is not null) {
                instClass.InstCommand = command;
                await ConnectDeviceAsync(instClass);
            }
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

        // OSC測定値取得
        private async Task<decimal> ReadOsc(InstClass instClass, int oscMeas) {
            try {
                VisibleProgressImage(true);

                instClass.InstCommand = $"MEASU:MEAS{oscMeas}:VAL?";
                var result = await ConnectDeviceAsync(instClass);
                decimal.TryParse(result, NumberStyles.AllowExponent | NumberStyles.Float, CultureInfo.InvariantCulture, out var output);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }

        // Serialロック
        private void SerialLockToggle() { SerialTextBox.IsEnabled = !SerialLockCheckBox.IsChecked ?? false; }
        // Serialインクリメント
        private void SerialIncrement(int i) {
            try {
                var serialText = SerialTextBox.Text;
                if (serialText.Length != 8) {
                    throw new Exception("シリアル文字数が一致しません。");
                }
                var middleDigits = serialText.Substring(4, 3);
                if (int.TryParse(middleDigits, out var currentValue)) {
                    currentValue = (currentValue += i) % 1000;
                    var newValue = currentValue.ToString("000");
                    var sb = new System.Text.StringBuilder();
                    sb.Append(serialText.AsSpan(0, 4));
                    sb.Append(newValue);
                    sb.Append(serialText.AsSpan(7));
                    SerialTextBox.Text = sb.ToString();
                }
                else {
                    throw new Exception("シリアルの中間値が数値に変換できません。");
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // キャリブレーション値取得
        private static (string a1l, string a1h, string a2l, string a2h) GetCalibration() {
            var hWnd = FindWindow(null, "マルチ流量計渦 [V01.08]");
            if (hWnd == IntPtr.Zero) { return (string.Empty, string.Empty, string.Empty, string.Empty); }

            var childHandles = FindChildWindows(hWnd);

            var a1l = GetMessageText(childHandles[332], WmGetText, 256);
            var a1h = GetMessageText(childHandles[342], WmGetText, 256);
            var a2l = GetMessageText(childHandles[352], WmGetText, 256);
            var a2h = GetMessageText(childHandles[362], WmGetText, 256);

            return (a1l, a1h, a2l, a2h);
        }
        // 受信ボタン実行
        private static void ReceiveData(IntPtr hWnd) {
            var childHandles = FindChildWindows(hWnd);
            _ = PostMessage(childHandles[9], BmClick, 0, 0);    // 受信ボタンクリック

            var hWnd3 = IntPtr.Zero;
            var timeOut = 0;
            while (hWnd3 == IntPtr.Zero) {
                hWnd3 = FindWindow(null, "FMRemote2014");
                timeOut += 1;
                Thread.Sleep(100);
                if (timeOut > 30) {
                    MessageBox.Show("タイムアウトしました");
                    return;
                }
            }
            var childHandles2 = FindChildWindows(hWnd3);
            _ = PostMessage(childHandles2[0], BmClick, 0, 0);    // メッセージウィンドウ「はい」クリック
        }
        // アナログトリムオープン
        private void AnalogTrimOpen(IntPtr hWnd) {
            if (_isProcessing) { return; }
            VisibleProgressImage(true);
            //Application.DoEvents();

            _ = PostMessage(hWnd, WmCommand, MenuItemIdAnalogTrim, 0);

            var hWnd3 = IntPtr.Zero;
            var timeOut = 0;
            while (hWnd3 == IntPtr.Zero) {
                hWnd3 = FindWindow(null, "FMRemote2014");
                timeOut += 1;
                Thread.Sleep(100);
                if (timeOut > 30) {
                    MessageBox.Show("タイムアウトしました");
                    VisibleProgressImage(false);
                    return;
                }
            }

            var childHandles = FindChildWindows(hWnd3);
            _ = PostMessage(childHandles[0], BmClick, 0, 0);    // メッセージウィンドウ「はい」クリック

            var hWnd2 = GetForegroundWindow();
            if (hWnd2 == IntPtr.Zero) { return; }

            (hWnd2, var windowText) = GetActiveWindow;

            while (windowText != "ANLOG TRIM") {
                Task.Delay(1000).Wait();
                (hWnd2, windowText) = GetActiveWindow;
            }

            Task.Delay(1000).Wait();
            Activate();

            ActivateAndBringToFront(hWnd);

            var childHandles2 = FindChildWindows(hWnd2);
            _ = PostMessage(childHandles2[5], WmLButtonDown, 0, 0);    // 入力欄フォーカス

            //Application.DoEvents();
            VisibleProgressImage(false);
        }
        // アナログトリムスタート
        private void AnalogTrimStart(IntPtr hWnd) {
            if (_isProcessing) { return; }
            VisibleProgressImage(true);
            //Application.DoEvents();

            var childHandles = FindChildWindows(hWnd);
            _ = PostMessage(childHandles[7], BmClick, 0, 0);    // STARTクリック

            Task.Delay(1000).Wait();

            ActivateAndBringToFront(hWnd);

            _ = PostMessage(childHandles[5], WmLButtonDown, 0, 0);  // 入力欄フォーカス
            _ = PostMessage(childHandles[5], EmSetSel, 0, -1);      // 入力欄全選択

            //Application.DoEvents();
            VisibleProgressImage(false);
        }
        // アナログトリムクローズ
        private void AnalogTrimClose(IntPtr hWnd) {
            if (_isProcessing) { return; }
            VisibleProgressImage(true);
            //Application.DoEvents();

            var childHandles = FindChildWindows(hWnd);
            _ = PostMessage(childHandles[0], BmClick, 0, 0);    // CLOSEクリック

            //Application.DoEvents();
            VisibleProgressImage(false);
        }
        // アナログトリム4mA
        private void AnalogTrim4mA(IntPtr hWnd) {
            if (_isProcessing) { return; }
            VisibleProgressImage(true);
            //Application.DoEvents();

            var childHandles = FindChildWindows(hWnd);
            _ = PostMessage(childHandles[2], BmClick, 0, 0);    // 4mAクリック

            Task.Delay(1000).Wait();

            ActivateAndBringToFront(hWnd);

            _ = PostMessage(childHandles[5], WmLButtonDown, 0, 0);  // 入力欄フォーカス
            _ = PostMessage(childHandles[5], EmSetSel, 0, -1);      // 入力欄全選択

            //Application.DoEvents();
            VisibleProgressImage(false);
        }
        // アナログトリム20mA
        private void AnalogTrim20mA(IntPtr hWnd) {
            if (_isProcessing) { return; }
            VisibleProgressImage(true);
            //Application.DoEvents();

            var childHandles = FindChildWindows(hWnd);
            _ = PostMessage(childHandles[3], BmClick, 0, 0);    // 20mAクリック

            Task.Delay(1000).Wait();

            ActivateAndBringToFront(hWnd);

            _ = PostMessage(childHandles[5], WmLButtonDown, 0, 0);  // 入力欄フォーカス
            _ = PostMessage(childHandles[5], EmSetSel, 0, -1);      // 入力欄全選択

            //Application.DoEvents();
            VisibleProgressImage(false);
        }
        // アクティブウィンドウ取得
        private static (IntPtr hWnd, string windowText) GetActiveWindow {
            get {
                var hWnd = GetForegroundWindow();
                var windowText = GetWindowText(hWnd);

                return (hWnd, windowText);
            }
        }


        [GeneratedRegex("^OP-650-10.*xlsm")]
        private static partial Regex MassFlowRegex();
        // アナログトリムスタート または OP-650-102をアクティブ
        private void ActionHotkeyNumAdd() {
            var (hWnd, windowText) = GetActiveWindow;
            switch (windowText.ToString()) {
                case "ANLOG TRIM": {
                        AnalogTrimStart(hWnd);
                        break;
                    }
                default: {
                        //すべてのプロセスを列挙する
                        foreach (var p in Process.GetProcesses()) {// メインウィンドウハンドルが存在し、かつタイトルが空でないことを確認
                            if (p.MainWindowHandle != IntPtr.Zero && !string.IsNullOrEmpty(p.MainWindowTitle)) {
                                //"OP-650-102"がメインウィンドウのタイトルに含まれているか調べる
                                if (MassFlowRegex().IsMatch(p.MainWindowTitle)) {
                                    //ウィンドウをアクティブにする
                                    var excelTitle = p.MainWindowTitle;
                                    var hWnd2 = FindWindow(null, excelTitle);
                                    ActivateAndBringToFront(hWnd2);
                                    break;
                                }
                            }
                        }
                        break;
                    }
            }
        }
        // OP-650-102 をアクティブ
        private static void ActionHotkeyMinus() {
            //すべてのプロセスを列挙する
            foreach (var p in Process.GetProcesses()) {
                // メインウィンドウハンドルが存在し、かつタイトルが空でないことを確認
                if (p.MainWindowHandle != IntPtr.Zero && !string.IsNullOrEmpty(p.MainWindowTitle)) {
                    //"OP-650-102"がメインウィンドウのタイトルに含まれているか調べる
                    if (MassFlowRegex().IsMatch(p.MainWindowTitle)) {
                        //ウィンドウをアクティブにする
                        var excelTitle = p.MainWindowTitle;
                        var hWnd = FindWindow(null, excelTitle);
                        ActivateAndBringToFront(hWnd);
                        break;
                    }
                }
            }
        }
        // マルチ流量計渦 [V01.08] をアクティブ
        private static void ActionHotkeyTilde() {
            var hWnd = FindWindow(null, "マルチ流量計渦 [V01.08]");
            ActivateAndBringToFront(hWnd);
        }
        // アナログトリムクローズ または マルチ流量計渦 [V01.08]をアクティブ
        private void ActionHotkeyNumSubtract() {
            var (hWnd, windowText) = GetActiveWindow;
            switch (windowText.ToString()) {
                case "ANLOG TRIM": {
                        AnalogTrimClose(hWnd);
                        break;
                    }
                default: {
                        string[] windowTitles = ["FMRemote2014", "マルチ流量計渦 [V01.08]"];

                        foreach (string title in windowTitles) {
                            IntPtr hWnd2 = FindWindow(null, title);
                            if (hWnd2 != IntPtr.Zero) {
                                ActivateAndBringToFront(hWnd2);
                                break;
                            }
                        }
                        break;
                    }
            }
        }
        // アナログトリム4mA
        private void ActionHotkeyNumDivide() {
            var (hWnd, windowText) = GetActiveWindow;
            if (windowText.ToString() == "ANLOG TRIM") {
                AnalogTrim4mA(hWnd);
            }
        }
        // アナログトリム20mA
        private void ActionHotkeyNumMultiply() {
            var (hWnd, windowText) = GetActiveWindow;
            if (windowText.ToString() == "ANLOG TRIM") {
                AnalogTrim20mA(hWnd);
            }
        }
        // 受信ボタン実行
        private static void ActionHotkeyAtsign() {
            var (hWnd, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                ReceiveData(hWnd);
            }
        }
        // Serialインクリメント＆コピー
        private void ActionHotkeyBracketL() {
            var (_, windowText) = GetActiveWindow;
            if (windowText.ToString() != "名前を付けて保存") {
                return;
            }
            if (string.IsNullOrEmpty(SerialTextBox.Text)) {
                return;
            }

            var sim = new InputSimulator();
            Clipboard.SetText(SerialTextBox.Text);
            sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.INSERT);
            SerialIncrement(1);
        }
        // キャリブレーション値コピー
        private async void ActionHotkeySemiColon() {
            try {
                VisibleProgressImage(true);

                var (hWnd, windowText) = GetActiveWindow;
                if (!MassFlowRegex().IsMatch(windowText)) {
                    return;
                }
                (var a1l, var a1h, var a2l, var a2h) = GetCalibration();
                if (string.IsNullOrEmpty(a1l)) { a1l = "error"; }
                if (string.IsNullOrEmpty(a1h)) { a1h = "error"; }
                if (string.IsNullOrEmpty(a2l)) { a2l = "error"; }
                if (string.IsNullOrEmpty(a2h)) { a2h = "error"; }

                // 各クリップボード操作とキー送信のブロック
                await PerformClipboardAndSendKeys(a1l);
                await PerformClipboardAndSendKeys(a1h);
                await PerformClipboardAndSendKeys(a2l);
                await PerformClipboardAndSendKeys(a2h);

                static async Task PerformClipboardAndSendKeys(string textToSet) {
                    await Task.Delay(500);
                    Clipboard.SetText(textToSet);
                    var sim = new InputSimulator();
                    sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.INSERT);
                    sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                }
            } finally {
                VisibleProgressImage(false);
            }
        }
        // FG&OSCローテーション
        private void ActionHotkeyColon() {
            if (_isProcessing) { return; }
            RotationFgOsc(false);
        }
        private void ActionHotkeyBracketR() {
            if (_isProcessing) { return; }
            RotationFgOsc(true);
        }
        // DMM値コピー
        private async void ActionHotkeyPeriod() {
            if (_isProcessing) { return; }

            var output = await ReadDmm(_instDmm);
            Clipboard.SetText((output * 1000).ToString("0.0000"));
            var sim = new InputSimulator();
            sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.INSERT);
        }
        // OSC mes1値コピー
        private async void ActionHotkeySlash() {
            if (_isProcessing) { return; }

            var output = await ReadOsc(_instOsc, 1);
            Clipboard.SetText(output.ToString("0.0000"));
            var sim = new InputSimulator();
            sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.INSERT);
            sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }
        // OSC mes2値コピー
        private async void ActionHotkeyBackslash() {
            if (_isProcessing) { return; }

            var output = await ReadOsc(_instOsc, 2);
            Clipboard.SetText(output.ToString("0.0000"));
            var sim = new InputSimulator();
            sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.INSERT);
            sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }
        // DCS切り替え
        private async void ActionHotkeyNum0() {
            var (_, windowText) = GetActiveWindow;

            var sim = new InputSimulator();

            switch (windowText.ToString()) {
                case "マルチ流量計渦 [V01.08]": {
                        if (!string.IsNullOrEmpty(_instDcs.VisaAddress)) {
                            await SwitchDcsAsync(_instDcs, 0);
                        }
                        break;
                    }
                case "FMRemote2014": {
                        sim.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
                        return;
                    }
                default: {
                        sim.Keyboard.TextEntry("0");
                        break;
                    }
            }
        }
        private void ActionHotkeyNum1() {
            var (_, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                if (_isProcessing) { return; }
                SwitchDcs(1);
                return;
            }
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry("1");
        }
        private void ActionHotkeyNum2() {
            var (_, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                if (_isProcessing) { return; }
                SwitchDcs(2);
                return;
            }
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry("2");
        }
        private void ActionHotkeyNum3() {
            var (_, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                if (_isProcessing) { return; }
                SwitchDcs(3);
                return;
            }
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry("3");
        }
        private void ActionHotkeyNum4() {
            var (_, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                if (_isProcessing) { return; }
                SwitchDcs(4);
                return;
            }
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry("4");
        }
        private void ActionHotkeyNum7() {
            var (hWnd, windowText) = GetActiveWindow;
            if (windowText.ToString() == "マルチ流量計渦 [V01.08]") {
                AnalogTrimOpen(hWnd);
                return;
            }
            var sim = new InputSimulator();
            sim.Keyboard.TextEntry("7");
        }


        // HotkKeyの登録
        private void SetHotKey() {
            _hotkeys.Clear();

            _hotkeys.AddRange([
                new(ModNone,    HotkeyNumAdd,       ActionHotkeyNumAdd),
                new(ModNone,    HotkeyNumSubtract,  ActionHotkeyNumSubtract),
                new(ModNone,    HotkeyNumDivide,    ActionHotkeyNumDivide),
                new(ModNone,    HotkeyNumMultiply,  ActionHotkeyNumMultiply),
                new(ModNone,    HotkeyMinus,        ActionHotkeyMinus),
                new(ModNone,    HotkeyTilde,        ActionHotkeyTilde),
                new(ModNone,    HotkeyAtsign,       ActionHotkeyAtsign),
                new(ModNone,    HotkeyBracketL,     ActionHotkeyBracketL),
                new(ModNone,    HotkeySemiColon,    ActionHotkeySemiColon),
                new(ModNone,    HotkeyNum0,         ActionHotkeyNum0),
                new(ModNone,    HotkeyNum7,         ActionHotkeyNum7),
                new(ModNone,    HotkeyNum1,         ActionHotkeyNum1),
                new(ModNone,    HotkeyNum2,         ActionHotkeyNum2),
                new(ModNone,    HotkeyNum3,         ActionHotkeyNum3),
                new(ModNone,    HotkeyNum4,         ActionHotkeyNum4),
            ]);
            if (!string.IsNullOrEmpty(_instFg01.VisaAddress) || !string.IsNullOrEmpty(_instFg02_1.VisaAddress) || !string.IsNullOrEmpty(_instFg02_2.VisaAddress) || !string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                _hotkeys.Add(new(ModNone, HotkeyPeriod, ActionHotkeyPeriod));
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                _hotkeys.AddRange([
                    new(ModNone, HotkeySlash, ActionHotkeySlash),
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
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

        // ウィンドウをアクティブにして最前面に持ってくる
        private static void ActivateAndBringToFront(IntPtr hWnd) {
            if (hWnd == IntPtr.Zero) { return; }

            // ウィンドウが最小化されているか確認し、復元する
            if (IsIconic(hWnd)) {
                // SW_RESTORE は最小化されたウィンドウを元のサイズに戻し、アクティブにします。
                ShowWindow(hWnd, SwRestore);
                // 少し待機することで、ウィンドウが表示されるのを確実にする（必要な場合のみ）
                System.Threading.Thread.Sleep(50);
            }
            SetForegroundWindow(hWnd);
        }

        // FGパネルの表示切替
        private void FgPanelVisible() {
            var i = FgNumberComboBox.SelectedIndex;

            (FgNumberGrid1.Visibility, FgNumberGrid2.Visibility) = i switch {
                0 => (Visibility.Visible, Visibility.Collapsed),
                1 => (Visibility.Collapsed, Visibility.Visible),
                _ => (Visibility.Collapsed, Visibility.Collapsed) // 0と1以外の場合
            };
        }

        // イベントハンドラ
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void InstListButton_Click(object sender, RoutedEventArgs e) { ShowInstList(); }
        private void HotKeyChekBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyChekBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }
        private void TopMostCheckBox_Checked(object sender, RoutedEventArgs e) { Topmost = true; }
        private void TopMostCheckBox_Unchecked(object sender, RoutedEventArgs e) { Topmost = false; }

        private void DcsOffButton_Click(object sender, RoutedEventArgs e) { SwitchDcs(0); }
        private void Dcs2VButton_Click(object sender, RoutedEventArgs e) { SwitchDcs(1); }
        private void Dcs8VButton_Click(object sender, RoutedEventArgs e) { SwitchDcs(2); }
        private void Dcs1VButton_Click(object sender, RoutedEventArgs e) { SwitchDcs(3); }
        private void Dcs7VButton_Click(object sender, RoutedEventArgs e) { SwitchDcs(4); }

        private void FgOscRotationButton_Click(object sender, RoutedEventArgs e) {
            if (_isProcessing) { return; }
            RotationFgOsc(true);
        }

        private void DcsRadioButton_Checked(object sender, RoutedEventArgs e) {
            if (sender is RadioButton checkedRadioButton) {
                checkedRadioButton.FontWeight = FontWeights.Bold;
                checkedRadioButton.Foreground = Brushes.Black;
            }
        }
        private void DcsRadioButton_Unchecked(object sender, RoutedEventArgs e) {
            if (sender is RadioButton checkedRadioButton) {
                checkedRadioButton.FontWeight = FontWeights.Regular;
                checkedRadioButton.Foreground = Brushes.Gray;
            }
        }
        private void FgOscRadioButton_Checked(object sender, RoutedEventArgs e) {
            if (sender is RadioButton checkedRadioButton) {
                checkedRadioButton.FontWeight = FontWeights.Bold;
                checkedRadioButton.Foreground = Brushes.Black;
            }
        }
        private void FgOscRadioButton_Unchecked(object sender, RoutedEventArgs e) {
            if (sender is RadioButton checkedRadioButton) {
                checkedRadioButton.FontWeight = FontWeights.Regular;
                checkedRadioButton.Foreground = Brushes.Gray;
            }
        }

        private void SerialLockCheckBox_Checked(object sender, RoutedEventArgs e) { SerialLockToggle(); }
        private void SerialLockCheckBox_Unchecked(object sender, RoutedEventArgs e) { SerialLockToggle(); }
        private void SerialBack_Click(object sender, RoutedEventArgs e) { SerialIncrement(-1); }
        private void SerialNext_Click(object sender, RoutedEventArgs e) { SerialIncrement(1); }

        private void FgNumberComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) { FgPanelVisible(); }
    }
}