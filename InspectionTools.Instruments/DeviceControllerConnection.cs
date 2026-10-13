namespace InspectionTools.Common;

/// <summary>本番用の接続（既存の DeviceController を変更せずに interface に載せるため、処理を渡すだけにしている）</summary>
public sealed class DeviceControllerConnection : IInstrumentConnection {

    public Task<string> ConnectAsync(InstClass instClass) => DeviceController.ConnectAsync(instClass);
}
