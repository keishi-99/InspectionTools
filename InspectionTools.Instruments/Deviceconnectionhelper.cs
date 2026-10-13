namespace InspectionTools.Common {
    /// <summary>
    /// 複数デバイスの並列接続を担当するクラス
    /// </summary>
    public static class DeviceConnectionHelper {

        /// <summary>
        /// 複数デバイスを並列接続し、全エラーを AggregateException にまとめてスローする
        /// </summary>
        public static async Task ConnectInParallelAsync(IEnumerable<InstClass> devices) {
            var tasks = devices.Select(async device => {
                try {
                    await DeviceController.ConnectAsync(device);
                } catch (Exception ex) {
                    throw new Exception($"[{device.Name}] 接続失敗: {ex.Message}", ex);
                }
            }).ToList();

            var whenAllTask = Task.WhenAll(tasks);
            try {
                await whenAllTask;
            } catch {
                // タスクが Faulted 状態のとき Exception は必ず非 null
                // 全デバイスのエラーを AggregateException にまとめてスロー
                throw new AggregateException("複数デバイスの接続に失敗しました", whenAllTask.Exception!.InnerExceptions);
            }
        }

        /// <summary>
        /// 並列接続し、各デバイスの結果文字列を返す
        /// </summary>
        public static async Task<IReadOnlyList<(InstClass Device, string Result)>> ConnectInParallelWithResultAsync(
            IEnumerable<InstClass> devices) {

            var deviceList = devices.ToList();
            var tasks = deviceList.Select(async device => {
                try {
                    var result = await DeviceController.ConnectAsync(device);
                    return (device, result);
                } catch (Exception ex) {
                    throw new Exception($"[{device.Name}] 接続失敗: {ex.Message}", ex);
                }
            }).ToList();

            var whenAllTask = Task.WhenAll(tasks);
            try {
                return await whenAllTask;
            } catch {
                // タスクが Faulted 状態のとき Exception は必ず非 null
                throw new AggregateException("複数デバイスの接続に失敗しました", whenAllTask.Exception!.InnerExceptions);
            }
        }
    }
}