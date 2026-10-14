using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using InspectionTools.Products;

namespace InspectionTools.App.ViewModels;

/// <summary>測定器の選択と接続の画面の状態と操作。通信は InstrumentConnection に任せ、ここは画面に出す文字と有効・無効だけを決める</summary>
public sealed partial class InstrumentSelectionViewModel : ObservableObject {

    private readonly ProductProcedure _procedure;
    private readonly InstrumentConnection _connection;
    private readonly string _catalogPath;

    /// <param name="procedure">製品の手順表（役割ごとの候補の条件を持つ）</param>
    /// <param name="connection">測定器への接続</param>
    /// <param name="catalogPath">測定器の一覧（VisaAddress.xml）のパス</param>
    public InstrumentSelectionViewModel(ProductProcedure procedure, InstrumentConnection connection, string catalogPath) {
        _procedure = procedure;
        _connection = connection;
        _catalogPath = catalogPath;
        Roles = Enum.GetValues<InstrumentRole>()
            .Where(role => procedure.Requirements.ContainsKey(role))
            .Select(role => new RoleSelectionViewModel(role, CreateLabel(role)))
            .ToList();
    }

    /// <summary>役割ごとの選択（役割の順）</summary>
    public IReadOnlyList<RoleSelectionViewModel> Roles { get; }

    /// <summary>接続中のコントローラー（接続していないときは null）</summary>
    public InstrumentController? Controller => _connection.Controller;

    /// <summary>接続しているか</summary>
    public bool IsConnected => _connection.IsConnected;

    /// <summary>処理中か（接続や解除の間は、操作を止める）</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanConnect), nameof(CanDisconnect), nameof(CanEditSelection))]
    private bool _isBusy;

    /// <summary>測定器の一覧に関する案内（読み込めないときの理由、読み飛ばした項目）。なければ空</summary>
    [ObservableProperty]
    private string _catalogMessage = "";

    /// <summary>操作の結果の案内（接続できなかった理由など）。なければ空</summary>
    [ObservableProperty]
    private string _message = "";

    /// <summary>接続の状態の表示（検査画面の上部にも出す）</summary>
    public string StatusText => IsConnected
        ? $"測定器: 接続済み（{_connection.ConnectedRoles.Count}台）"
        : "測定器: 未接続";

    /// <summary>接続ボタンを押せるか</summary>
    public bool CanConnect => !IsConnected && !IsBusy;

    /// <summary>解除ボタンを押せるか</summary>
    public bool CanDisconnect => IsConnected && !IsBusy;

    /// <summary>選択欄を変えられるか（接続中や処理中は変えられない）</summary>
    public bool CanEditSelection => !IsConnected && !IsBusy;

    /// <summary>画面を開くときに呼ぶ。測定器の一覧を読み直して、役割ごとの候補を更新する（接続中は、選択を変えないので何もしない）</summary>
    public void Refresh() {
        if (IsConnected) {
            return;
        }

        InstrumentCatalog catalog;
        try {
            catalog = InstrumentCatalog.LoadFile(_catalogPath);
        } catch (InstrumentCatalogException ex) {
            foreach (var role in Roles) {
                role.ReplaceCandidates([]);
            }
            CatalogMessage = $"{ex.Message}\n実行ファイルのフォルダに VisaAddress.xml を置いてください（VisaAddress.template.xml が見本です）。";
            return;
        }

        foreach (var role in Roles) {
            role.ReplaceCandidates(catalog.FindCandidates(_procedure.Requirements[role.Role]));
        }
        CatalogMessage = catalog.Warnings.Count == 0
            ? ""
            : "測定器の一覧の一部を読み飛ばしました。\n" + string.Join("\n", catalog.Warnings.Select(warning => $"・{warning}"));
    }

    /// <summary>選んだ測定器に接続する。1台でも失敗したら、すべて解除して、失敗した測定器を役割ごとに示す</summary>
    [RelayCommand]
    public async Task ConnectAsync() {
        if (!CanConnect) {
            return;
        }

        var selections = Roles
            .Where(role => role.SelectedEntry is not null)
            .ToDictionary(role => role.Role, role => role.SelectedEntry!);
        if (selections.Count == 0) {
            Message = "接続する測定器を選んでください。";
            return;
        }

        IsBusy = true;
        Message = "";
        foreach (var role in Roles) {
            role.ResultText = "";
        }

        try {
            await _connection.ConnectAsync(selections);
            foreach (var role in Roles) {
                role.ResultText = selections.ContainsKey(role.Role) ? "接続しました" : "使いません";
            }
            Message = "接続しました。";
        } catch (InstrumentConnectionException ex) {
            ShowConnectionFailures(selections.Keys, ex);
        } catch (ArgumentException ex) {
            // 同じ測定器が複数の役割に選ばれているときなど、通信を始める前に分かる誤り
            Message = ex.Message;
        } finally {
            IsBusy = false;
            NotifyConnectionChanged();
        }
    }

    /// <summary>すべて解除する。測定器には何も送らない（出力が ON のままでも、そのまま残す）</summary>
    [RelayCommand]
    public void Disconnect() {
        if (!CanDisconnect) {
            return;
        }

        _connection.Disconnect();
        foreach (var role in Roles) {
            role.ResultText = "";
        }
        Message = "解除しました。";
        NotifyConnectionChanged();
    }

    // 接続に失敗した役割には理由を、成功した役割には解除したことを示す
    private void ShowConnectionFailures(IEnumerable<InstrumentRole> selectedRoles, InstrumentConnectionException exception) {
        var failedMessages = exception.Failures.ToDictionary(failure => failure.Role, failure => failure.Message);
        foreach (var role in Roles) {
            role.ResultText = !selectedRoles.Contains(role.Role)
                ? "使いません"
                : failedMessages.TryGetValue(role.Role, out var failure)
                    ? failure
                    : "接続できましたが、ほかの測定器の失敗のため解除しました";
        }
        Message = "接続できない測定器があったため、すべて解除しました。";
    }

    // 接続の状態に合わせて、画面の表示と有効・無効を更新する
    private void NotifyConnectionChanged() {
        OnPropertyChanged(nameof(IsConnected));
        OnPropertyChanged(nameof(Controller));
        OnPropertyChanged(nameof(StatusText));
        OnPropertyChanged(nameof(CanConnect));
        OnPropertyChanged(nameof(CanDisconnect));
        OnPropertyChanged(nameof(CanEditSelection));
    }

    // 画面に出す役割の名前
    private static string CreateLabel(InstrumentRole role) => role switch {
        InstrumentRole.Dmm01 => "DMM01（マルチメーター1）",
        InstrumentRole.Dmm02 => "DMM02（マルチメーター2）",
        InstrumentRole.Fg => "FG（周波数発生器）",
        InstrumentRole.Dcs => "DCS（電流電圧発生器）",
        InstrumentRole.Cnt => "CNT（カウンタ）",
        _ => role.ToString(),
    };
}
