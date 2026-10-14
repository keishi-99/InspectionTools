using InspectionTools.Common;

namespace InspectionTools.Products;

/// <summary>
/// EL5000AC の測定器の手順表。コマンドは旧アプリ（EL5000ACUserControl）のものをそのまま使う。
/// 項目と動作 ID の対応は、products/EL5000AC.json の actionId で決める
/// </summary>
public static class El5000AcProcedure {

    /// <summary>製品名</summary>
    public const string Product = "EL5000AC";

    /// <summary>EL5000AC の手順表を作る</summary>
    public static ProductProcedure Create() => new(Product, CreateProfiles(), CreateActions(), CreateIdleStates(), CreateRequirements());

    // 出力する機器は、使わない項目では OFF に戻す（出力を出しっぱなしにしないため）。測定する機器は、項目が指定しない限り触らない
    private static Dictionary<InstrumentRole, string> CreateIdleStates() => new() {
        [InstrumentRole.Dcs] = "OFF",
        [InstrumentRole.Fg] = "OFF",
    };

    // 役割ごとに選べる測定器の条件（旧アプリの選択欄の絞り込みと同じ。通信方式は 1=ADC、2=VISA USB、3=GPIB）
    private static Dictionary<InstrumentRole, RoleRequirement> CreateRequirements() => new() {
        [InstrumentRole.Dmm01] = new("デジタルマルチメータ", [1, 2]),
        [InstrumentRole.Dmm02] = new("デジタルマルチメータ", [1, 2]),
        [InstrumentRole.Fg] = new("ファンクションジェネレータ", [2]),
        [InstrumentRole.Dcs] = new("電流電圧発生器", [2]),
        [InstrumentRole.Cnt] = new("ユニバーサルカウンタ", [3]),
    };

    // ===== 動作（actionId ごと） =====

    private static Dictionary<string, ItemAction> CreateActions() {
        var actions = new Dictionary<string, ItemAction> {
            // 測定器の値を取得する項目
            ["dmm01-aci"] = new([new(InstrumentRole.Dmm01, "ACI")], new Acquisition(InstrumentRole.Dmm01, Scale: 1000m, Decimals: 3)),
            ["dmm02-dcv"] = new([new(InstrumentRole.Dmm02, "DCV")], new Acquisition(InstrumentRole.Dmm02, Scale: 1m, Decimals: 4)),
            ["dmm02-dci"] = new([new(InstrumentRole.Dmm02, "DCI")], new Acquisition(InstrumentRole.Dmm02, Scale: 1000m, Decimals: 4)),
            ["cnt-pulse-width"] = new([], new Acquisition(InstrumentRole.Cnt, Scale: 1000m, Decimals: 3)),

            // 周波数発生器を出力する項目
            ["fg-on"] = new([new(InstrumentRole.Fg, "ON")]),
        };

        // 電圧電流発生器の出力（アナログ校正の作業と、アナログ入力の確認で使う）
        foreach (var state in new[] { "4mA", "12mA", "20mA", "1V", "3V", "5V" }) {
            actions[$"dcs-{state}"] = new([new(InstrumentRole.Dcs, state)]);
        }
        return actions;
    }

    // ===== 役割ごとのコマンド表 =====

    private static Dictionary<InstrumentRole, InstrumentProfile> CreateProfiles() {
        var dmm01Aci = new SwitchCommand { Adc = "*RST,F6,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:AC;*OPC?", Query = true };
        var dmm02Dcv = new SwitchCommand { Adc = "*RST,F1,R0,*OPC?", Visa = "*RST;:INIT:CONT 1;:VOLT:DC:RANG:AUTO ON;*OPC?", Query = true };
        var dmm02Dci = new SwitchCommand { Adc = "*RST,F5,R6,*OPC?", Visa = "*RST;:INIT:CONT 1;:CONF:CURR:DC;*OPC?", Query = true };

        return new Dictionary<InstrumentRole, InstrumentProfile> {
            [InstrumentRole.Dmm01] = new(
                Init: dmm01Aci,
                InitialState: "ACI",
                States: new Dictionary<string, SwitchCommand> { ["ACI"] = dmm01Aci }),

            [InstrumentRole.Dmm02] = new(
                Init: dmm02Dcv,
                InitialState: "DCV",
                States: new Dictionary<string, SwitchCommand> { ["DCV"] = dmm02Dcv, ["DCI"] = dmm02Dci }),

            [InstrumentRole.Fg] = new(
                Init: new SwitchCommand { Visa = "*RST;:FUNC SQU;:FREQ 7E3;:VOLT 3.0VPP;:VOLT:OFFS 1.5;*OPC?", Query = true },
                InitialState: "OFF",
                States: new Dictionary<string, SwitchCommand> {
                    // 旧アプリのとおり、*OPC? を付けていても返事は読まない（Query = false）
                    ["OFF"] = new() { Visa = ":OUTP OFF;*OPC?" },
                    ["ON"] = new() { Visa = ":OUTP ON;*OPC?" },
                }),

            [InstrumentRole.Dcs] = new(
                Init: new SwitchCommand { Visa = "SIR3,SOI+0,SBY", Gpib = "RCF5R6S0EO0E" },
                InitialState: "OFF",
                States: new Dictionary<string, SwitchCommand> {
                    ["OFF"] = new() { Visa = "SIR2,SOI+0MA,SBY" },
                    ["4mA"] = new() { Visa = "SIR2,SOI+4MA,OPR" },
                    ["12mA"] = new() { Visa = "SIR2,SOI+12MA,OPR" },
                    ["20mA"] = new() { Visa = "SIR2,SOI+20MA,OPR" },
                    ["1V"] = new() { Visa = "SVR5,SOV+1V,OPR" },
                    ["3V"] = new() { Visa = "SVR5,SOV+3V,OPR" },
                    ["5V"] = new() { Visa = "SVR5,SOV+5V,OPR" },
                }),

            [InstrumentRole.Cnt] = new(
                Init: new SwitchCommand { Gpib = "*RST;:FUNC PWID;:INPA:COUP DC;:INPA:LPF ON;:INPA:SLOP NEG;:FRUN ON;" },
                InitialState: "PWID",
                States: new Dictionary<string, SwitchCommand>(),
                // 値を読み取ったあと、次の測定のために測定を再開させる
                AfterRead: new SwitchCommand { Gpib = ":FRUN ON" }),
        };
    }
}
