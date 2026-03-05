namespace InspectionTools.Common {
    internal static class InstrumentHelper {
        /// <summary>
        /// 計測器インスタンスを安全に解放します。
        /// </summary>
        public static void SafeDispose(InstClass? instrument) {
            try {
                instrument?.Dispose();
            } catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine($"Instrument dispose error: {ex.Message}");
            }
        }
    }
}
