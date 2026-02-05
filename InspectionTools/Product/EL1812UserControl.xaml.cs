using InspectionTools.Common;
using System.Data;
using System.Windows;
using VisaComLib;
using WindowsInput;
using static InspectionTools.Common.Win32Wrapper;
using static InspectionTools.MainWindow;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.Product {
    /// <summary>
    /// EL1812UserControl.xaml の相互作用ロジック
    /// </summary>
    public partial class EL1812UserControl : UserControl, IMainWindowAware, IDisposable {

        private MainWindow? _mainWindow;
        private bool _disposed = false;

        public void SetMainWindow(MainWindow mainWindow) {
            _mainWindow = mainWindow;
        }

        private readonly DmmInstClass _instDmm = new();
        private readonly FgInstClass _instFg = new();
        private readonly OscInstClass _instOsc = new();

        private record SwitchCommand {
            public string Text { get; init; } = string.Empty;
            public string Adc { get; init; } = string.Empty;
            public string Visa { get; init; } = string.Empty;
            public string Gpib { get; init; } = string.Empty;
            public bool Query { get; init; } = false;
        }
        private readonly Dictionary<InstClass, (SwitchCommand Init, List<SwitchCommand> Settings)> _dicCommands = [];
        private readonly Dictionary<InstClass, (SwitchCommand Init, List<SwitchCommand> Settings)> _dicReverseCommands = [];

        public EL1812UserControl() {
            InitializeComponent();
        }

        #region IDisposable Implementation

        /// <summary>
        /// リソースの解放（IDisposableパターン）
        /// </summary>
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// リソースの解放処理
        /// </summary>
        /// <param name="disposing">マネージドリソースも解放する場合はtrue</param>
        protected virtual void Dispose(bool disposing) {
            if (_disposed) {
                return;
            }

            if (disposing) {
                // マネージドリソースの解放
                try {
                    // ホットキーのクリア
                    ClearHotKey();

                    // 計測器の解放
                    DisposeInstrument(_instDmm);
                    DisposeInstrument(_instFg);
                    DisposeInstrument(_instOsc);

                    // 辞書のクリア
                    _dicCommands.Clear();
                } catch (Exception ex) {
                    // Dispose中のエラーはログに記録するのみ
                    System.Diagnostics.Debug.WriteLine($"Dispose error: {ex.Message}");
                }
            }

            // アンマネージドリソースの解放（必要に応じて）
            // ...

            _disposed = true;
        }

        /// <summary>
        /// 個別の計測器インスタンスを解放
        /// </summary>
        private static void DisposeInstrument(InstClass instrument) {
            if (instrument == null) return;

            try {
                // 計測器がIDisposableを実装している場合
                if (instrument is IDisposable disposable) {
                    disposable.Dispose();
                }
                else {
                    // ResetPropertiesで状態をリセット
                    instrument.ResetProperties();
                }
            } catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine($"Instrument dispose error: {ex.Message}");
            }
        }

        /// <summary>
        /// オブジェクトが破棄済みかチェック
        /// </summary>
        private void ThrowIfDisposed() {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        /// <summary>
        /// ファイナライザ
        /// </summary>
        ~EL1812UserControl() {
            Dispose(false);
        }

        #endregion

        private int _rotateMenuNumber = 0;
        private List<string> _listRotateMenuTitle = [];

        // 起動時
        private void LoadEvents() {
            ThrowIfDisposed();
            InstListImport();
            RegMenuTitle();
            var parentWindow = Window.GetWindow(this);
            MainWindow.AdjustWindowSizeToUserControl(parentWindow);
        }
        private void InstListImport() {
            // デジタルマルチメータ、ファンクションジェネレータ、オシロスコープのコンボボックスを更新する
            MainWindow.UpdateComboBox(DmmComboBox, "デジタルマルチメータ", [1, 2]);
            MainWindow.UpdateComboBox(FgComboBox, "ファンクションジェネレータ", [3]);
            MainWindow.UpdateComboBox(OscComboBox, "オシロスコープ", [2]);
        }
        // 処理中の画像を表示/非表示にします。
        private void VisibleProgressImage(bool isVisible) {
            MainWindow.IsProcessing = isVisible;
            MainGrid.IsEnabled = !isVisible;
            ProgressRing.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }

        // 選択した機器のVisaAddressを取得
        private void SelectInst() {
            MainWindow.GetVisaAddress(_instDmm, DmmComboBox);
            MainWindow.GetVisaAddress(_instFg, FgComboBox);
            MainWindow.GetVisaAddress(_instOsc, OscComboBox);
        }
        // 機器設定辞書登録
        private void RegDictionary() {

            _dicCommands[_instDmm] =
                (
                    Init: new() { Adc = "*RST,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;*OPC?", Query = true },
                    Settings: []
                );

            _dicCommands[_instFg] =
                (
                    Init: new() { Text = "00", Gpib = "*RST;OMO 1;BES 0;BTY 1;FNC 3;TRS 1;TRE 0;BSS 1;BSV -100.0;FRQ 2E+02;HIV 4.5;LOV 2.0;MRK 79;SIG 0;" },
                    Settings: [
                        new(){ Text="00", Gpib = "SIG 0;OMO 1;BTY 1;FRQ 2E+02;"},
                        new(){ Text="01", Gpib = "SIG 1;MRK 79;TRG 1;" },
                        new(){ Text="02", Gpib = "MRK 1;TRG 1;" },
                        new(){ Text="03", Gpib = "MRK 839;TRG 1;" },
                        new(){ Text="04", Gpib = "MRK 1;TRG 1;" },
                        new(){ Text="05", Gpib = "MRK 78;TRG 1;" },
                        new(){ Text="06", Gpib = "MRK 2;TRG 1;" },
                        new(){ Text="07", Gpib = "FRQ 3E+3;MRK 1E+5;TRG 1;" },
                        new(){ Text="08", Gpib = "MRK 10;TRG 1;" },
                        new(){ Text="09", Gpib = "MRK 1002;TRG 1;" },
                        new(){ Text="10", Gpib = "OMO 0;MRK 1;FRQ 10;" },
                        new(){ Text="11" ,Gpib = "SIG 0;" },
                        new(){ Text="12", Gpib = "SIG 1;" },
                        new(){ Text="13", Gpib = "SIG 0;" },
                        new(){ Text="14", Gpib = "SIG 1;" }
                    ]
                );

            _dicCommands[_instOsc] =
                (
                    Init: new() {
                        Visa =
                        """
                        *RST;:HEADER 0;
                        :CH1:SCALE 5.0E0;POSITION -2.0E0;COUPLING DC;
                        :HORIZONTAL:MAIN:SCALE 5.0E-5;
                        :HORIZONTAL:MAIN:POSITION 10.0E-5;
                        :TRIGGER:MAIN:LEVEL 1.5E1;
                        :TRIGGER:MAIN:EDGE:SLOPE FALL;
                        :MEASUREMENT:MEAS1:TYPE NWIDTH;SOURCE CH1;
                        *OPC?
                        """
                        ,
                        Query = true
                    },
                    Settings: [
                        new(){
                            Text = "150u",
                            Visa=
                                 """
                                :HORIZONTAL:MAIN:SCALE 5.0E-5;
                                :HORIZONTAL:MAIN:POSITION 10.0E-5;
                                *OPC?
                                """,
                            Query = true
                        },
                        new(){
                            Text = "1m",
                            Visa=
                                """
                                :HORIZONTAL:MAIN:SCALE 25.0E-5;
                                *OPC?
                                """,
                            Query = true
                        },
                        new(){
                            Text = "50m",
                            Visa=
                                """
                                :HORIZONTAL:MAIN:SCALE 10.0E-3;
                                :HORIZONTAL:MAIN:POSITION 20.0E-3;
                                *OPC?
                                """,
                            Query = true
                        },
                    ]
                );

            _dicReverseCommands[_instFg] =
                (
                    Init: new(),
                    Settings: [
                        new(){Text= "00",Gpib="SIG 0;HIV 4.5;LOV 2.0;FRQ 2E+02;"},
                        new(){Text= "01",Gpib="HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 79;"},
                        new(){Text= "02",Gpib="HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 1;"},
                        new(){Text= "03",Gpib="HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 839;"},
                        new(){Text= "04",Gpib="HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 1;"},
                        new(){Text= "05",Gpib="HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 78;"},
                        new(){Text= "06",Gpib="HIV 4.5;LOV 2.0;FRQ 2E+02;MRK 2;"},
                        new(){Text= "07",Gpib="FRQ 3E+3;MRK 1E+5;"},
                        new(){Text= "08",Gpib="FRQ 3E+3;MRK 10;"},
                        new(){Text= "09",Gpib="BTY 1;FRQ 3E+3;MRK 1002;"},
                        new(){Text= "10",Gpib="SIG 1;FRQ 10;"},
                        new(){Text= "11",Gpib="SIG 0;FRQ 10;"},
                        new(){Text= "12",Gpib="SIG 1;FRQ 10;"},
                        new(){Text= "13",Gpib="SIG 0;FRQ 10;"},
                        new(){Text= "14",Gpib="SIG 1;MRK 1;BTY 0;FRQ 10;"},
                    ]
                );
        }
        // 機器初期設定
        private void FormatSet() {
            (_instDmm.InstCommand, _instDmm.Query) = ResolveCommand(_dicCommands[_instDmm].Init, _instDmm.SignalType);
            (_instFg.InstCommand, _instFg.Query) = ResolveCommand(_dicCommands[_instFg].Init, _instFg.SignalType);
            (_instOsc.InstCommand, _instOsc.Query) = ResolveCommand(_dicCommands[_instOsc].Init, _instOsc.SignalType);
        }
        private static (string Cmd, bool Query) ResolveCommand(SwitchCommand sw, int signalType) {
            return signalType switch {
                1 => (sw.Adc, sw.Query),
                2 => (sw.Visa, sw.Query),
                3 => (sw.Gpib, sw.Query),
                _ => (string.Empty, false),
            };
        }
        // 検査項目名登録
        private void RegMenuTitle() {
            _listRotateMenuTitle = [
                "1.絶縁抵抗",
                "2.耐電圧",
                "3.検査準備",
                "4.外観構造",
                "5.発信器電源電圧",
                "6.バッチ動作",
                "7.バッチ量設定\r\nスケーラ-値積算\r\nパルス出力",
                "8.パルス未登録\r\nリーク警報\r\n過充填警報",
                "9.チャンネル選択",
                "10.外部入力",
                "11.パルス出力幅",
                "12.電源ON/OFF時\r\n設定値",
                "13.電圧、電流パルス(12V)",
                "14.バッチカウンタ初期化",
                "15.結果保存",
                "検査情報"
            ];
        }

        // 機器接続
        private async void ConnectInstAsync() {
            ThrowIfDisposed();

            try {
                _mainWindow?.SetButtonEnabled("ProductListButton", false);

                HotKeyCheckBox.IsChecked = false;
                VisibleProgressImage(true);

                SelectInst();

                InstClass[] devices = [_instDmm, _instFg, _instOsc];
                RegDictionary();
                FormatSet();
                var tasks = devices.Select(device => MainWindow.ConnectDeviceAsync(device));
                await Task.WhenAll(tasks);

                if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                    FgRotateBackButton.IsEnabled = true;
                    FgRotateNextButton.IsEnabled = true;
                    FgOutputOnButton.IsEnabled = true;
                    FgOutputOffButton.IsEnabled = true;
                    FgTriggerButton.IsEnabled = true;
                    FgRotateRangeTextBox.Text = "00";
                }
                if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                    OscRotationButton.IsEnabled = true;
                    OscRotateRangeTextBox.Text = "150u";
                }

                DmmComboBox.IsEnabled = false;
                FgComboBox.IsEnabled = false;
                OscComboBox.IsEnabled = false;
                ConnectButton.IsEnabled = false;
                ReleaseButton.IsEnabled = true;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー");
            } finally {
                VisibleProgressImage(false);
            }
        }

        // 解除
        private void Release() {
            VisibleProgressImage(false);

            _instDmm.ResetProperties();
            _instFg.ResetProperties();
            _instOsc.ResetProperties();

            _mainWindow?.SetButtonEnabled("ProductListButton", true);
            DmmComboBox.IsEnabled = true;
            FgComboBox.IsEnabled = true;
            OscComboBox.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            ReleaseButton.IsEnabled = false;
            HotKeyCheckBox.IsChecked = false;
            FgRotateBackButton.IsEnabled = false;
            FgRotateNextButton.IsEnabled = false;
            FgOutputOnButton.IsEnabled = false;
            FgOutputOffButton.IsEnabled = false;
            FgTriggerButton.IsEnabled = false;
            OscRotationButton.IsEnabled = false;
        }

        // FG切り替え
        private async void RotationFg(FgInstClass fgInstClass, bool isNext) {
            ThrowIfDisposed();

            try {
                if (string.IsNullOrEmpty(fgInstClass.VisaAddress)) { return; }
                VisibleProgressImage(true);

                var settings = _dicCommands[fgInstClass].Settings;
                fgInstClass.SettingNumber = (fgInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[fgInstClass.SettingNumber];
                fgInstClass.InstCommand = fgInstClass.SignalType switch {
                    1 => sw.Adc,
                    2 => sw.Visa,
                    3 => sw.Gpib,
                    _ => string.Empty,
                };
                fgInstClass.Query = sw.Query;

                if (fgInstClass.InstCommand == string.Empty) { return; }

                await MainWindow.RotationFgAsync(fgInstClass);

                FgRotateRangeTextBox.Text = sw.Text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        // FG 出力ON/OFF
        private async void OutputFg(string cmd) {
            ThrowIfDisposed();

            try {
                if (string.IsNullOrEmpty(_instFg.VisaAddress)) { return; }
                VisibleProgressImage(true);

                await OutputFgAsync(_instFg, cmd); ;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        private static async Task OutputFgAsync(FgInstClass fgInstClass, string cmd) {
            fgInstClass.InstCommand = $"{cmd};";
            await MainWindow.ConnectDeviceAsync(fgInstClass);
        }
        // OSC切り替え
        private async void RotationOsc(OscInstClass oscInstClass, bool isNext) {
            ThrowIfDisposed();

            try {
                if (string.IsNullOrEmpty(oscInstClass.VisaAddress)) { return; }
                VisibleProgressImage(true);

                var dic = isNext ? _dicCommands : _dicReverseCommands;
                var settings = dic[_instOsc].Settings;
                oscInstClass.SettingNumber = (oscInstClass.SettingNumber + (isNext ? 1 : -1) + settings.Count) % settings.Count;

                var sw = settings[oscInstClass.SettingNumber];
                oscInstClass.InstCommand = oscInstClass.SignalType switch {
                    1 => sw.Adc,
                    2 => sw.Visa,
                    3 => sw.Gpib,
                    _ => string.Empty,
                };
                oscInstClass.Query = sw.Query;

                if (oscInstClass.InstCommand == string.Empty) { return; }

                await MainWindow.RotationOscAsync(oscInstClass);

                OscRotateRangeTextBox.Text = sw.Text;

            } catch (Exception ex) {
                Release();
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally {
                VisibleProgressImage(false);
            }
        }
        // DMM測定値取得
        private async Task<decimal> ReadDmm(DmmInstClass dmmInstClass) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var output = await MainWindow.ReadDmm(dmmInstClass);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // OSC測定値取得
        private async Task<decimal> ReadOsc(OscInstClass oscInstClass, int meas) {
            ThrowIfDisposed();

            try {
                VisibleProgressImage(true);

                var output = await MainWindow.ReadOsc(oscInstClass, meas);

                return output;

            } finally {
                VisibleProgressImage(false);
            }
        }
        // 検査項目切り替え
        private void RotateMenu(int i) {

            var hWnd = FindWindow(null, "工程間検査");
            var hWnd2 = FindWindow(null, "検査情報画面");
            var hWnd3 = FindWindow(null, "確認");

            var windowToActivate = IntPtr.Zero;
            // ウィンドウが最小化されているか確認し、復元する
            if (hWnd3 != IntPtr.Zero) {
                windowToActivate = hWnd3;
            }
            else if (hWnd2 != IntPtr.Zero) {
                windowToActivate = hWnd2;
            }
            else if (hWnd != IntPtr.Zero) {
                windowToActivate = hWnd;
            }

            if (windowToActivate != IntPtr.Zero) {
                if (IsIconic(windowToActivate)) {
                    ShowWindow(windowToActivate, SwRestore);
                    System.Threading.Thread.Sleep(50);
                }
                SetForegroundWindow(windowToActivate);
            }
            else {
                return;
            }

            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);

            nint parentWindow;
            string itemName;

            switch (windowText.ToString()) {
                case "確認": {
                        return;
                    }
                case "検査情報画面": {
                        parentWindow = hWnd2;
                        itemName = "登録";
                        break;
                    }
                default: {
                        var n = _listRotateMenuTitle.Count;
                        _rotateMenuNumber = (n + ((_rotateMenuNumber + i) % n)) % n;
                        MenuRotateRangeTextBox.Text = (_rotateMenuNumber + 1).ToString("00");
                        var windowName = (_rotateMenuNumber != 15) ? "検査実施" : "検査情報";
                        parentWindow = FindWindowEx(hWnd, IntPtr.Zero, null, windowName);
                        itemName = _listRotateMenuTitle[_rotateMenuNumber];
                        break;
                    }
            }
            if (parentWindow == IntPtr.Zero) { return; }

            var targetWindow = FindWindowEx(parentWindow, IntPtr.Zero, null, itemName);
            if (targetWindow == IntPtr.Zero) { return; }

            _ = PostMessage(targetWindow, BmClick, 0, 0);

            nint hWnd4;
            List<nint> childHandles;
            switch (itemName) {
                case "5.発信器電源電圧":
                    hWnd4 = FindWindow(null, "発信器電源電圧");
                    //SetForegroundWindow(hWnd4);
                    childHandles = FindChildWindows(hWnd4);
                    _ = PostMessage(childHandles[5], EmSetSel, 0, -1);     // 入力欄全選択
                    _ = PostMessage(childHandles[5], WmClear, 0, -1);      // 入力欄消去
                    _ = PostMessage(childHandles[11], EmSetSel, 0, -1);     // 入力欄全選択
                    _ = PostMessage(childHandles[11], WmClear, 0, -1);      // 入力欄消去
                    _ = PostMessage(childHandles[11], WmLButtonDown, 0, 0);  // フォーカス
                    _ = PostMessage(childHandles[11], WmLButtonUp, 0, 0);  // フォーカス
                    break;
                case "10.外部入力":
                    hWnd4 = FindWindow(null, "外部入力");
                    //SetForegroundWindow(hWnd4);
                    childHandles = FindChildWindows(hWnd4);
                    _ = PostMessage(childHandles[11], WmLButtonDown, 0, 0);  // フォーカス
                    _ = PostMessage(hWnd4, WmLButtonUp, 0, 0);  // フォーカス
                    break;
                case "11.パルス出力幅":
                    hWnd4 = FindWindow(null, "パルス出力幅");
                    childHandles = FindChildWindows(hWnd4);
                    _ = PostMessage(childHandles[6], EmSetSel, 0, -1);     // 入力欄全選択
                    _ = PostMessage(childHandles[6], WmClear, 0, -1);      // 入力欄消去
                    _ = PostMessage(childHandles[12], EmSetSel, 0, -1);     // 入力欄全選択
                    _ = PostMessage(childHandles[12], WmClear, 0, -1);      // 入力欄消去
                    _ = PostMessage(childHandles[18], EmSetSel, 0, -1);     // 入力欄全選択
                    _ = PostMessage(childHandles[18], WmClear, 0, -1);      // 入力欄消去
                    _ = PostMessage(childHandles[23], WmLButtonDown, 0, 0);  // フォーカス
                    _ = PostMessage(hWnd4, WmLButtonUp, 0, 0);  // フォーカス
                    break;
                case "13.電圧、電流パルス(12V)":
                    break;
                case "14.バッチカウンタ初期化":
                    hWnd4 = FindWindow(null, "バッチカウンタ初期化");
                    childHandles = FindChildWindows(hWnd4);
                    _ = PostMessage(childHandles[15], WmLButtonDown, 0, 0);  // フォーカス
                    _ = PostMessage(hWnd4, WmLButtonUp, 0, 0);  // フォーカス
                    break;
            }
        }
        // Serialロック
        private void SerialLockToggle() { SerialTextBox.IsEnabled = !SerialLockCheckBox.IsChecked ?? false; }
        // Serialインクリメント
        private void SerialIncrement(int i) {
            try {
                var serialText = SerialTextBox.Text;
                if (serialText.Length == 0) {
                    return;
                }
                if (serialText.Length != 8) {
                    throw new Exception("シリアル文字数が一致しません。");
                }
                var middleDigits = serialText.Substring(4, 3);
                if (int.TryParse(middleDigits, out var currentValue)) {
                    // 0〜999 の範囲で循環
                    currentValue = (currentValue + i + 1000) % 1000;

                    // 3桁に0埋め
                    var newValue = currentValue.ToString("000");

                    // 元のシリアル番号を再構築
                    var sb = new System.Text.StringBuilder();
                    sb.Append(serialText.AsSpan(0, 4));  // 先頭4文字
                    sb.Append(newValue);                  // 中間3桁
                    sb.Append(serialText.AsSpan(7));     // 7文字目以降

                    SerialTextBox.Text = sb.ToString();
                }
                else {
                    throw new Exception("シリアルの中間値が数値に変換できません。");
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Countロック
        private void CountLockToggle() { CountTextBox.IsEnabled = !CountLockCheckBox.IsChecked ?? false; }
        // Countインクリメント
        private void CountIncrement(int i) {
            if (!int.TryParse(CountTextBox.Text, out var intCount)) { return; }
            intCount += i;
            CountTextBox.Text = intCount.ToString();
        }

        // OSC切り替え
        private void ActionSwitchOscRange(bool isNext) {
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);
            if (windowText != "パルス出力幅") { return; }

            // 11.パルス出力幅でのみ有効
            if (MainWindow.IsProcessing) { return; }

            RotationOsc(_instOsc, isNext);
        }
        // OSC測定値コピー
        private async void ActionCopyOscValue() {
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);
            if (windowText != "パルス出力幅") { return; }

            // 11.パルス出力幅でのみ有効
            if (MainWindow.IsProcessing) { return; }

            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);

            var output = await ReadOsc(_instOsc, 1);
            (var value, var format) = _instOsc.SettingNumber switch {
                0 => (output * 1000000, "0.0"),
                1 or 2 => (output * 1000, "0.00"),
                _ => (output, "0.00"),
            };
            sim.Keyboard.TextEntry(value.ToString(format));
        }
        // DMM測定値コピー
        private async void ActionCopyDmmValue() {
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);
            if (windowText != "発信器電源電圧") { return; }

            // 5.発信器電源電圧でのみ有効
            if (MainWindow.IsProcessing) { return; }

            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
            sim.Keyboard.KeyPress(VirtualKeyCode.BACK);

            var output = await ReadDmm(_instDmm);
            sim.Keyboard.TextEntry(output.ToString("0.0"));
            Thread.Sleep(200);
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
        }
        // Serial貼り付け
        private void ActionPasteSerial() {
            if (string.IsNullOrEmpty(SerialTextBox.Text)) { return; }

            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);
            if (windowText != "検査情報画面") { return; }

            // 検査情報画面でのみ有効
            var childHandles = FindChildWindows(foregroundWindow);
            _ = PostMessage(childHandles[15], WmLButtonDown, 0, 0);  // フォーカス
            _ = PostMessage(childHandles[15], WmLButtonUp, 0, 0);  // フォーカス

            var sim = new InputSimulator();
            sim.Keyboard.TextEntry(SerialTextBox.Text);
            SerialIncrement(1);
        }
        // Count貼り付け
        private void ActionPasteCount() {
            if (string.IsNullOrEmpty(CountTextBox.Text)) { return; }

            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            var windowText = GetWindowText(foregroundWindow);
            if (windowText != "検査情報画面") { return; }

            // 検査情報画面でのみ有効
            var childHandles = FindChildWindows(foregroundWindow);
            _ = PostMessage(childHandles[6], WmLButtonDown, 0, 0);  // フォーカス
            _ = PostMessage(childHandles[6], WmLButtonUp, 0, 0);  // フォーカス

            var sim = new InputSimulator();
            sim.Keyboard.TextEntry(CountTextBox.Text);
            CountIncrement(1);
        }
        // FG切り替え
        private void ActionSwitchFg(bool isNext) {
            if (MainWindow.IsProcessing) { return; }
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            // FGを使用する項目でのみ有効
            var windowText = GetWindowText(foregroundWindow);
            if (windowText is
                not "バッチ動作" and
                not "スケーラ値" and
                not "警報" and
                not "パルス出力幅"
                ) { return; }
            RotationFg(_instFg, isNext);
        }
        // OSC+FG切り替え
        private void ActionSwitchFgOsc(bool isNext) {
            if (MainWindow.IsProcessing) { return; }
            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            // パルス出力幅 でのみ有効
            var windowText = GetWindowText(foregroundWindow);
            if (windowText is
                not "パルス出力幅"
                ) { return; }
            RotationFg(_instFg, isNext);
            RotationOsc(_instOsc, isNext);
        }

        // メニュー切り替え
        private void ActionHotkeyBracketL() { RotateMenu(1); }
        private void ActionHotkeyNum7() { RotateMenu(1); }
        // Tabキー
        private static void ActionHotkeyAtsign() {
            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
        }
        private static void ActionHotkeyNum8() {
            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
        }
        // OSC切り替え
        private void ActionHotkeyBracketR() { ActionSwitchOscRange(true); }
        private void ActionHotkeyNum1() { ActionSwitchOscRange(true); }
        // OSC測定値コピー
        private void ActionHotkeyColon() { ActionCopyOscValue(); }
        private void ActionHotkeyNum6() { ActionCopyOscValue(); }
        // DMM測定値コピー
        private void ActionHotkeySemiColon() { ActionCopyDmmValue(); }
        private void ActionHotkeyNum9() { ActionCopyDmmValue(); }
        // Serial貼り付け
        private void ActionHotkeyPeriod() { ActionPasteSerial(); }
        private void ActionHotkeyNumSubtract() { ActionPasteSerial(); }
        // Count貼り付け
        private void ActionHotkeyComma() { ActionPasteCount(); }
        private void ActionHotkeyNumAdd() { ActionPasteCount(); }
        // FG切り替え
        private void ActionHotkeyBackslash() { ActionSwitchFg(true); }
        private void ActionHotkeyNum5() { ActionSwitchFg(true); }
        private void ActionHotkeyShiftBackslash() { ActionSwitchFg(false); }
        private void ActionHotkeyShiftNum5() { ActionSwitchFg(false); }
        private void ActionHotkeyNum2() {
            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            ActionSwitchFg(true);
        }
        private void ActionHotkeyNum3() {
            if (MainWindow.IsProcessing) { return; }

            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);

            var foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero) { return; }

            // FGを使用する項目でのみ有効
            var windowText = GetWindowText(foregroundWindow);
            if (windowText is
                not "バッチ動作" and
                not "スケーラ値" and
                not "警報" and
                not "パルス出力幅"
                ) { return; }
            ActionSwitchFgOsc(true);
        }

        // HotKeyの登録
        private void SetHotKey() {
            ThrowIfDisposed();

            MainWindow.HotkeysList.Clear();

            MainWindow.HotkeysList.AddRange([
                new(ModNone, HotkeyBracketL, ActionHotkeyBracketL),
                new(ModNone, HotkeyAtsign, ActionHotkeyAtsign),
                new(ModNone, HotkeyPeriod, ActionHotkeyPeriod),
                new(ModNone, HotkeyComma, ActionHotkeyComma),
                new(ModNone, HotkeyNum7, ActionHotkeyNum7),
                new(ModNone, HotkeyNum8, ActionHotkeyNum8),
                new(ModNone, HotkeyNumAdd, ActionHotkeyNumAdd),
                new(ModNone, HotkeyNumSubtract, ActionHotkeyNumSubtract),
            ]);
            if (!string.IsNullOrEmpty(_instFg.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBackslash, ActionHotkeyBackslash),
                    new(ModShift, HotkeyBackslash, ActionHotkeyShiftBackslash),
                    new(ModNone, HotkeyNum5, ActionHotkeyNum5),
                    new(ModShift, HotkeyNum5, ActionHotkeyShiftNum5),
                ]);
            }
            if (!string.IsNullOrEmpty(_instDmm.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeySemiColon, ActionHotkeySemiColon),
                    new(ModNone, HotkeyNum9,ActionHotkeyNum9),
                ]);
            }
            if (!string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyBracketR, ActionHotkeyBracketR),
                    new(ModNone, HotkeyColon, ActionHotkeyColon),
                    new(ModNone, HotkeyNum1, ActionHotkeyNum1),
                    new(ModNone, HotkeyNum6, ActionHotkeyNum6),
                ]);
            }
            if (!string.IsNullOrEmpty(_instFg.VisaAddress) && !string.IsNullOrEmpty(_instOsc.VisaAddress)) {
                MainWindow.HotkeysList.AddRange([
                    new(ModNone, HotkeyNum2, ActionHotkeyNum2),
                    new(ModNone, HotkeyNum3, ActionHotkeyNum3),
                ]);
            }

            MainWindow.SetHotKey();
        }
        private static void ClearHotKey() {
            MainWindow.ClearHotKey();
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

        // イベントハンドラ
        private void UserControl_Loaded(object sender, RoutedEventArgs e) { LoadEvents(); }
        private void ConnectButton_Click(object sender, RoutedEventArgs e) { ConnectInstAsync(); }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e) { Release(); }
        private void HotKeyCheckBox_Checked(object sender, RoutedEventArgs e) { SetHotKey(); }
        private void HotKeyCheckBox_Unchecked(object sender, RoutedEventArgs e) { ClearHotKey(); }

        private void OscRotationButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationOsc(_instOsc, true);
        }
        private void FgRotateBackButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, false);
        }
        private void FgRotateNextButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            RotationFg(_instFg, true);
        }
        private void FgOutputOnButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            OutputFg("SIG 1");
        }
        private void FgOutputOffButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            OutputFg("SIG 0");
        }
        private void FgTriggerButton_Click(object sender, RoutedEventArgs e) {
            if (MainWindow.IsProcessing) { return; }
            OutputFg("TRG 1");
        }

        private void MenuRotateBackButton_Click(object sender, RoutedEventArgs e) { RotateMenu(-1); }
        private void MenuRotateNextButton_Click(object sender, RoutedEventArgs e) { RotateMenu(1); }

        private void SerialBack_Click(object sender, RoutedEventArgs e) { SerialIncrement(-1); }
        private void SerialNext_Click(object sender, RoutedEventArgs e) { SerialIncrement(1); }
        private void SerialLockCheckBox_Checked(object sender, RoutedEventArgs e) { SerialLockToggle(); }
        private void SerialLockCheckBox_Unchecked(object sender, RoutedEventArgs e) { SerialLockToggle(); }

        private void CountBack_Click(object sender, RoutedEventArgs e) { CountIncrement(-1); }
        private void CountNext_Click(object sender, RoutedEventArgs e) { CountIncrement(1); }
        private void CountLockCheckBox_Checked(object sender, RoutedEventArgs e) { CountLockToggle(); }
        private void CountLockCheckBox_Unchecked(object sender, RoutedEventArgs e) { CountLockToggle(); }
        private void UserControl_Unloaded(object sender, RoutedEventArgs e) { Dispose(); }

    }
}
