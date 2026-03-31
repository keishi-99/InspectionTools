using System.Globalization;

namespace InspectionTools.Common {
    /// <summary>
    /// 測定器の種類ごとの高レベル操作（測定値取得・切り替えなど）を担当するクラス
    /// </summary>
    public static class InstrumentService {

        /// <summary>
        /// DMM 測定値取得
        /// </summary>
        public static async Task<decimal> ReadDmmAsync(DmmInstClass dmmInstClass) {
            dmmInstClass.InstCommand = dmmInstClass.SignalType switch {
                1 => string.Empty,
                2 => "FETC?",
                _ => throw new ApplicationException($"未対応の SignalType: {dmmInstClass.SignalType}"),
            };

            var result = await DeviceController.ConnectAsync(dmmInstClass);
            return ParseDecimal(result);
        }

        /// <summary>
        /// CNT 測定値取得
        /// </summary>
        public static async Task<decimal> ReadCntAsync(CntInstClass cntInstClass) {
            (cntInstClass.InstCommand, cntInstClass.Query) = cntInstClass.SignalType switch {
                3 => (":MEAS?XNOW;:FRUN ON", true),
                _ => throw new ApplicationException($"未対応の SignalType: {cntInstClass.SignalType}"),
            };

            var result = await DeviceController.ConnectAsync(cntInstClass);
            return ParseDecimal(result);
        }

        /// <summary>
        /// OSC 測定値取得
        /// </summary>
        public static async Task<decimal> ReadOscAsync(OscInstClass oscInstClass, int measIndex) {
            oscInstClass.InstCommand = $"MEASU:MEAS{measIndex}:VAL?";
            var result = await DeviceController.ConnectAsync(oscInstClass);
            return ParseDecimal(result);
        }

        /// <summary>
        /// FG 切り替え
        /// </summary>
        public static async Task RotateFgAsync(FgInstClass fgInstClass) {
            await DeviceController.ConnectAsync(fgInstClass);
        }

        /// <summary>
        /// OSC 切り替え
        /// </summary>
        public static async Task RotateOscAsync(OscInstClass oscInstClass) {
            await DeviceController.ConnectAsync(oscInstClass);
        }

        /// <summary>
        /// 測定値文字列を decimal に変換する共通処理
        /// </summary>
        private static decimal ParseDecimal(string value) {
            decimal.TryParse(
                value,
                NumberStyles.AllowExponent | NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var output);
            return output;
        }
    }
}