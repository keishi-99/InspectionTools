namespace InspectionTools.Common {
    internal static class InstrumentHelper {
        /// <summary>
        /// 計測器インスタンスを安全に解放します。
        /// </summary>
        public static void SafeDispose(InstClass? instrument) {
            try {
                instrument?.Dispose();
            } catch (Exception ex) {
                var label = instrument switch {
                    null => "Unknown",
                    { Name: var name } when !string.IsNullOrEmpty(name) => name,
                    _ => instrument.GetType().Name,
                };
                System.Diagnostics.Debug.WriteLine($"Instrument dispose error [{label}]: {ex.Message}");
            }
        }

        /// <summary>
        /// 信号種別に応じたコマンド文字列とクエリフラグを返します。
        /// </summary>
        public static (string Cmd, bool Query) ResolveCommand(SwitchCommand sw, int signalType) => signalType switch {
            1 => (sw.Adc, sw.Query),
            2 or 4 => (sw.Visa, sw.Query),
            3 => (sw.Gpib, sw.Query),
            _ => (string.Empty, false),
        };

        /// <summary>
        /// DMM選択の重複チェックを行います。同一機器が選択された場合は例外をスローします。
        /// </summary>
        public static void ValidateDmmSelection(params int[] indices) {
            var valid = indices.Where(i => i >= 1).ToList();
            if (valid.Count != valid.Distinct().Count())
                throw new InvalidOperationException("同じ測定器が選択されています。");
        }
    }
}
