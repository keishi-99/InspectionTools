namespace InspectionTools.Common {
    internal static class InstrumentHelper {
        /// <summary>
        /// 計測器インスタンスを安全に解放します。
        /// </summary>
        public static void SafeDispose(InstClass? instrument) {
            try {
                instrument?.Dispose();
            } catch (Exception ex) {
                var label = string.IsNullOrEmpty(instrument?.Name)
                    ? instrument?.GetType().Name ?? "Unknown"
                    : instrument.Name;
                System.Diagnostics.Debug.WriteLine($"Instrument dispose error [{label}]: {ex.Message}");
            }
        }
    }
}
