namespace InspectionTools.Common {
    /// <summary>
    /// 測定器への低レベル接続処理（VISA・ADC）を担当するクラス
    /// </summary>
    public static class DeviceController {

        private const int TimeOut = 3; // タイムアウトまでの時間(sec)
        private static readonly SemaphoreSlim _visaLock = new(1, 1);
        private static readonly SemaphoreSlim _adcLock = new(1, 1);

        /// <summary>
        /// デバイスの SignalType に応じて接続メソッドを振り分ける
        /// </summary>
        public static async Task<string> ConnectAsync(InstClass instClass) {
            return instClass.Index == -1
                ? ""
                : instClass.SignalType switch {
                    1 => await ConnectAdcAsync(instClass),
                    2 or 3 or 4 => await ConnectVisaAsync(instClass),
                    _ => throw new ApplicationException($"未対応の SignalType: {instClass.SignalType}"),
                };
        }

        /// <summary>
        /// VISA接続
        /// </summary>
        public static async Task<string> ConnectVisaAsync(InstClass instClass) {
            await _visaLock.WaitAsync();
            try {
                return await Task.Run(() => {
                    using var usbDev = new USBDeviceManager();
                    usbDev.OpenDev(instClass.VisaAddress);
                    usbDev.OutputDev(instClass.InstCommand);
                    return instClass.Query ? usbDev.InputDev() : string.Empty;
                });
            } finally {
                _visaLock.Release();
            }
        }

        /// <summary>
        /// ADC接続
        /// </summary>
        public static async Task<string> ConnectAdcAsync(InstClass instClass) {
            await _adcLock.WaitAsync();
            try {
                uint hDev = 0;
                var rcvDt = "";
                uint rcvLen = 50;
                if (!uint.TryParse(instClass.VisaAddress, out var id))
                    throw new Exception($"VisaAddressが不正な形式です: '{instClass.VisaAddress}'");
                try {
                    if (AusbWrapper.Start(TimeOut) != 0 || AusbWrapper.Open(ref hDev, id) != 0) {
                        throw new Exception("開始できません");
                    }
                    if (!string.IsNullOrEmpty(instClass.InstCommand)) {
                        if (AusbWrapper.Write(hDev, instClass.InstCommand) != 0) {
                            throw new Exception("コマンドの送信に失敗しました");
                        }
                    }
                    if (AusbWrapper.Read(hDev, ref rcvDt, ref rcvLen) != 0) {
                        throw new Exception("メッセージの受信に失敗しました");
                    }
                } finally {
                    _ = AusbWrapper.Close(hDev);
                    _ = AusbWrapper.End();
                }
                return rcvDt;
            } finally {
                _adcLock.Release();
            }
        }
    }
}