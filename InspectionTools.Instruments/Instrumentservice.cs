using System.Globalization;

namespace InspectionTools.Common {
    /// <summary>
    /// 測定器の種類ごとの高レベル操作（測定値取得・切り替えなど）を担当するクラス
    /// </summary>
    public static class InstrumentService {

        // 既存の呼び出し（接続を指定しない版）は本番用の接続を使う（呼び出し側のコードを変えないため）
        private static readonly IInstrumentConnection _defaultConnection = new DeviceControllerConnection();

        /// <summary>
        /// DCS 測定値取得
        /// </summary>
        public static Task<decimal> ReadDcsAsync(DcsInstClass dcsInstClass) => ReadDcsAsync(dcsInstClass, _defaultConnection);

        /// <summary>
        /// DCS 測定値取得（接続を指定する版。テストでダミーの接続に差し替えるために使う）
        /// </summary>
        public static async Task<decimal> ReadDcsAsync(DcsInstClass dcsInstClass, IInstrumentConnection connection) {
            dcsInstClass.InstCommand = dcsInstClass.SignalType switch {
                2 => "MEAS:CURR?",
                _ => throw new InvalidOperationException($"未対応の SignalType: {dcsInstClass.SignalType}"),
            };

            var result = await connection.ConnectAsync(dcsInstClass);
            return ParseDecimal(result);
        }

        /// <summary>
        /// DMM 測定値取得
        /// </summary>
        public static Task<decimal> ReadDmmAsync(DmmInstClass dmmInstClass) => ReadDmmAsync(dmmInstClass, _defaultConnection);

        /// <summary>
        /// DMM 測定値取得（接続を指定する版）
        /// </summary>
        public static async Task<decimal> ReadDmmAsync(DmmInstClass dmmInstClass, IInstrumentConnection connection) {
            dmmInstClass.InstCommand = dmmInstClass.SignalType switch {
                1 => string.Empty,
                2 => "FETC?",
                _ => throw new InvalidOperationException($"未対応の SignalType: {dmmInstClass.SignalType}"),
            };

            var result = await connection.ConnectAsync(dmmInstClass);
            return ParseDecimal(result);
        }

        /// <summary>
        /// CNT 測定値取得
        /// </summary>
        public static Task<decimal> ReadCntAsync(CntInstClass cntInstClass) => ReadCntAsync(cntInstClass, _defaultConnection);

        /// <summary>
        /// CNT 測定値取得（接続を指定する版）
        /// </summary>
        public static async Task<decimal> ReadCntAsync(CntInstClass cntInstClass, IInstrumentConnection connection) {
            (cntInstClass.InstCommand, cntInstClass.Query) = cntInstClass.SignalType switch {
                3 => (":MEAS?XNOW", true),
                _ => throw new InvalidOperationException($"未対応の SignalType: {cntInstClass.SignalType}"),
            };

            var result = await connection.ConnectAsync(cntInstClass);
            return ParseDecimal(result);
        }

        /// <summary>
        /// OSC 測定値取得
        /// </summary>
        public static Task<decimal> ReadOscAsync(OscInstClass oscInstClass, int measIndex) => ReadOscAsync(oscInstClass, measIndex, _defaultConnection);

        /// <summary>
        /// OSC 測定値取得（接続を指定する版）
        /// </summary>
        public static async Task<decimal> ReadOscAsync(OscInstClass oscInstClass, int measIndex, IInstrumentConnection connection) {
            oscInstClass.InstCommand = $"MEASU:MEAS{measIndex}:VAL?";
            var result = await connection.ConnectAsync(oscInstClass);
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