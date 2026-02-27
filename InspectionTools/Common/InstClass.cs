namespace InspectionTools.Common {

    public class InstClass : IDisposable {
        internal USBDeviceManager UsbDev { get; set; } = new();
        public string Category { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string VisaAddress { get; set; } = string.Empty;
        public int SignalType { get; set; } = 0;
        public int Index { get; set; } = 0;
        public string Tag { get; set; } = string.Empty;
        public string InstCommand { get; set; } = string.Empty;
        public bool Query { get; set; } = false;
        public int SettingNumber { get; set; } = 0;

        private bool _disposed = false;

        public void ResetProperties() {
            if (_disposed) throw new ObjectDisposedException(nameof(InstClass));
            UsbDev?.Dispose();
            UsbDev = new();
            Name = string.Empty;
            Category = string.Empty;
            VisaAddress = string.Empty;
            SignalType = 0;
            Index = 0;
            Tag = string.Empty;
            InstCommand = string.Empty;
            Query = false;
            SettingNumber = 0;
        }
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing) {
            if (_disposed) {
                return;
            }

            if (disposing) {
                // マネージドリソースの解放
                UsbDev?.Dispose();
            }

            // アンマネージドリソースを解放する処理があるならここに追加

            _disposed = true;
        }
    }

    // CNT用クラス
    public class CntInstClass : InstClass { }

    // DCS用クラス
    public class DcsInstClass : InstClass { }

    // DMM用クラス
    public class DmmInstClass : InstClass {
        public DmmMode CurrentMode { get; set; } = DmmMode.None;
    }
    public enum DmmMode {
        None,
        DCV,
        DCI,
        RES
    }

    // FG用クラス
    public class FgInstClass : InstClass { }

    // OSC用クラス
    public class OscInstClass : InstClass { }

}
