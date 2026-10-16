using InspectionTools.Common;

namespace InspectionTools.Products;

/// <summary>
/// PAF5 の測定器の手順表。コマンドは旧アプリ（PAF5UserControl）のものをそのまま使う。
/// PAF5 はタイプ（分離形、一体形2線式など）ごとに製品名が分かれるが、手順表は共通。
/// 項目と動作 ID の対応は、products/PAF5-*.json の actionId で決める
/// </summary>
public static class Paf5Procedure {

    /// <summary>PAF5 の製品名の先頭（「PAF5-分離形」のように、タイプを後ろにつなげる）</summary>
    public const string ProductPrefix = "PAF5-";

    /// <summary>PAF5 の手順表を作る。product はタイプごとの製品名（例: PAF5-分離形）</summary>
    public static ProductProcedure Create(string product) => new(product, CreateProfiles(), CreateActions(), CreateIdleStates(), CreateRequirements());

    // 周波数発生器は、使わない項目では OFF に戻す（出力を出しっぱなしにしないため）。測定する機器は、項目が指定しない限り触らない
    private static Dictionary<InstrumentRole, string> CreateIdleStates() => new() {
        [InstrumentRole.Fg] = "OFF",
    };

    // 役割ごとに選べる測定器の条件（旧アプリの選択欄の絞り込みと同じ。通信方式は 1=ADC、2=VISA USB、3=GPIB）
    private static Dictionary<InstrumentRole, RoleRequirement> CreateRequirements() => new() {
        [InstrumentRole.Dmm01] = new("デジタルマルチメータ", [1, 2]),
        [InstrumentRole.Dmm02] = new("デジタルマルチメータ", [1, 2]),
        [InstrumentRole.Fg] = new("ファンクションジェネレータ", [2]),
        [InstrumentRole.Osc] = new("オシロスコープ", [2]),
    };

    // ===== 動作（actionId ごと） =====

    private static Dictionary<string, ItemAction> CreateActions() {
        var actions = new Dictionary<string, ItemAction> {
            // 測定器の値を取得する項目。DMM01 は V、DMM02 は μA、OSC は mV（旧アプリのホットキーと同じ換算と桁）
            ["dmm01-dcv"] = new([new(InstrumentRole.Dmm01, "DCV")], new Acquisition(InstrumentRole.Dmm01, Scale: 1m, Decimals: 2)),
            ["dmm02-dci"] = new([new(InstrumentRole.Dmm02, "DCI")], new Acquisition(InstrumentRole.Dmm02, Scale: 1000000m, Decimals: 2)),
            ["osc-noise"] = new([new(InstrumentRole.Osc, "1ms")], new Acquisition(InstrumentRole.Osc, Scale: 1000m, Decimals: 3)),

            // 利得特性: 入力の周波数とオシロスコープの時間軸を合わせる（ATT は人が調整して入力するので、値は取得しない）
            ["gain-20hz"] = new([new(InstrumentRole.Fg, "20Hz"), new(InstrumentRole.Osc, "10ms")]),
            ["gain-200hz"] = new([new(InstrumentRole.Fg, "200Hz"), new(InstrumentRole.Osc, "1ms")]),
            ["gain-2khz"] = new([new(InstrumentRole.Fg, "2000Hz"), new(InstrumentRole.Osc, "100us")]),

            // ノッチ: 最小周波数は人が入力する。その周波数での振幅と、5000Hz での振幅を、オシロスコープで測る
            ["notch-frequency"] = new([new(InstrumentRole.Fg, "6250Hz"), new(InstrumentRole.Osc, "100us")]),
            ["notch-amplitude-min"] = new([new(InstrumentRole.Fg, "6250Hz"), new(InstrumentRole.Osc, "100us")], new Acquisition(InstrumentRole.Osc, Scale: 1000m, Decimals: 3)),
            ["notch-amplitude-5000hz"] = new([new(InstrumentRole.Fg, "5000Hz"), new(InstrumentRole.Osc, "100us")], new Acquisition(InstrumentRole.Osc, Scale: 1000m, Decimals: 3)),
        };

        // 周波数発生器だけを使う項目（トリガレベル 1000Hz、分周機能 400Hz、ローカット機能の各周波数）
        foreach (var frequency in new[] { "1000", "400", "0.51", "1", "3.5", "5.7", "33", "46", "8" }) {
            actions[$"fg-{frequency}hz"] = new([new(InstrumentRole.Fg, $"{frequency}Hz")]);
        }
        return actions;
    }

    // ===== 役割ごとのコマンド表 =====

    private static Dictionary<InstrumentRole, InstrumentProfile> CreateProfiles() {
        var dmm01Dcv = new SwitchCommand { Adc = "*RST,R7,*OPC?", Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG 200;*OPC?", Query = true };
        var dmm02Dci = new SwitchCommand { Adc = "*RST,F5,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?", Query = true };

        return new Dictionary<InstrumentRole, InstrumentProfile> {
            [InstrumentRole.Dmm01] = new(
                Init: dmm01Dcv,
                InitialState: "DCV",
                States: new Dictionary<string, SwitchCommand> { ["DCV"] = dmm01Dcv }),

            [InstrumentRole.Dmm02] = new(
                Init: dmm02Dci,
                InitialState: "DCI",
                States: new Dictionary<string, SwitchCommand> { ["DCI"] = dmm02Dci }),

            [InstrumentRole.Fg] = new(
                Init: new SwitchCommand { Visa = "*RST;:FREQ 20;:VOLT 0.44VPP;*OPC?", Query = true },
                InitialState: "OFF",
                States: CreateFgStates()),

            [InstrumentRole.Osc] = new(
                Init: new SwitchCommand {
                    Visa =
                        """
                        *RST;:HEADER 0;
                        :ACQUIRE:MODE AVERAGE;
                        :CH1:SCALE 1.0E-1;COUPLING DC;
                        :CURSOR:FUNCTION HBARS;SELECT:SOURCE CH1;:CURSOR:HBARS:POSITION1 2.0E-1;POSITION2 -2.0E-1;
                        :HORIZONTAL:MAIN:SCALE 1.0E-3;
                        :MEASUREMENT:MEAS1:TYPE PK2PK;SOURCE CH1;
                        *OPC?
                        """,
                    Query = true,
                },
                InitialState: "1ms",
                States: new Dictionary<string, SwitchCommand> {
                    ["1ms"] = new() { Visa = ":HORIZONTAL:MAIN:SCALE 1.0E-3;*OPC?", Query = true },
                    ["10ms"] = new() { Visa = ":HORIZONTAL:MAIN:SCALE 1.0E-2;*OPC?", Query = true },
                    ["100us"] = new() { Visa = ":HORIZONTAL:MAIN:SCALE 1.0E-4;*OPC?", Query = true },
                }),
        };
    }

    // 周波数発生器の状態(旧アプリの順番表のコマンド)。切り替えのたびに、いったん出力を止めてから周波数を変える
    private static Dictionary<string, SwitchCommand> CreateFgStates() {
        static SwitchCommand Command(string visa) => new() { Visa = visa, Query = true };

        return new Dictionary<string, SwitchCommand> {
            ["OFF"] = Command(":FREQ 20;:OUTPUT OFF;*OPC?"),
            ["20Hz"] = Command(":FREQ 20;:OUTPUT ON;*OPC?"),
            ["200Hz"] = Command(":OUTPUT OFF;:FREQ 200;:OUTPUT ON;*OPC?"),
            ["2000Hz"] = Command(":OUTPUT OFF;:FREQ 2000;:OUTPUT ON;*OPC?"),
            ["6250Hz"] = Command(":OUTPUT OFF;:FREQ 6250;:OUTPUT ON;*OPC?"),
            ["5000Hz"] = Command(":OUTPUT OFF;:FREQ 5000;:OUTPUT ON;*OPC?"),
            ["1000Hz"] = Command(":OUTPUT OFF;:FREQ 1000;:OUTPUT ON;*OPC?"),
            ["400Hz"] = Command(":OUTPUT OFF;:FREQ 400;:OUTPUT ON;*OPC?"),
            ["0.51Hz"] = Command(":OUTPUT OFF;:FREQ 0.51;:OUTPUT ON;*OPC?"),
            ["1Hz"] = Command(":OUTPUT OFF;:FREQ 1.0;:OUTPUT ON;*OPC?"),
            ["3.5Hz"] = Command(":OUTPUT OFF;:FREQ 3.5;:OUTPUT ON;*OPC?"),
            ["5.7Hz"] = Command(":OUTPUT OFF;:FREQ 5.7;:OUTPUT ON;*OPC?"),
            ["33Hz"] = Command(":OUTPUT OFF;:FREQ 33;:OUTPUT ON;*OPC?"),
            ["46Hz"] = Command(":OUTPUT OFF;:FREQ 46;:OUTPUT ON;*OPC?"),
            ["8Hz"] = Command(":OUTPUT OFF;:FREQ 8.0;:OUTPUT ON;*OPC?"),
        };
    }
}
