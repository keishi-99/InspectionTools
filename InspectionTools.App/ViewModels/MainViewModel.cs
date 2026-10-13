using System.IO;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using InspectionTools.Inspection;

namespace InspectionTools.App.ViewModels;

/// <summary>検査画面の状態と操作。処理は InspectionSession と SessionStore を呼ぶだけにして、ここには画面の部品を持ち込まない</summary>
public sealed partial class MainViewModel : ObservableObject {

    private readonly ProductDefinition _definition;
    private readonly SessionStore _store;
    private readonly IDialogService _dialog;
    private InspectionSession _session;

    public MainViewModel(ProductDefinition definition, SessionStore store, IDialogService dialog) {
        _definition = definition;
        _store = store;
        _dialog = dialog;
        _session = new InspectionSession(definition.Items, unitCount: 1);
        Rows = CreateRows();
    }

    /// <summary>製品名</summary>
    public string Product => _definition.Product;

    /// <summary>表の行（検査項目の順）。台数が変わると作り直される</summary>
    [ObservableProperty]
    private IReadOnlyList<RowViewModel> _rows;

    /// <summary>台数</summary>
    public int UnitCount => _session.UnitCount;

    /// <summary>この製品で設定できる最大の台数</summary>
    public int MaxUnitCount => _definition.MaxUnitCount;

    /// <summary>台数の入力欄の文字列（Enter などで確定するまでは、台数そのものは変わらない）</summary>
    [ObservableProperty]
    private string _unitCountText = "1";

    /// <summary>現在の項目の位置</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CurrentItem), nameof(IsValueInput), nameof(IsConfirmation), nameof(IsTask), nameof(CurrentRangeText))]
    private int _currentItemIndex;

    /// <summary>現在の台の位置</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CurrentUnitText))]
    private int _currentUnitIndex;

    /// <summary>現在の台の表示（「1台目」の形）</summary>
    public string CurrentUnitText => $"{CurrentUnitIndex + 1}台目";

    /// <summary>入力欄の文字列</summary>
    [ObservableProperty]
    private string _inputText = "";

    /// <summary>現在の項目</summary>
    public InspectionItem CurrentItem => _session.Items[CurrentItemIndex];

    /// <summary>値の入力欄を使う項目か（測定項目・入力項目）</summary>
    public bool IsValueInput => CurrentItem.Kind is ItemKind.Measurement or ItemKind.Input;

    /// <summary>合格／不合格のボタンを使う項目か</summary>
    public bool IsConfirmation => CurrentItem.Kind == ItemKind.Confirmation;

    /// <summary>完了ボタンを使う項目か</summary>
    public bool IsTask => CurrentItem.Kind == ItemKind.Task;

    /// <summary>合格範囲の表示（範囲のない項目は空）</summary>
    public string CurrentRangeText => CurrentItem.LowerLimit is { } lower && CurrentItem.UpperLimit is { } upper
        ? $"合格範囲: {lower} ～ {upper}"
        : "";

    /// <summary>起動時に呼ぶ。前回の途中保存があれば、続きから再開するか確認する</summary>
    public void Start() {
        var loaded = _store.Load(Product, _definition.Items);
        switch (loaded.Status) {
            case LoadStatus.Restored:
                if (_dialog.Confirm("前回の途中保存があります。続きから再開しますか？\n（「いいえ」を選ぶと、途中保存を破棄して最初から始めます）")) {
                    ReplaceSession(loaded.Session!);
                } else {
                    DeleteSavedState();
                }
                break;
            case LoadStatus.Failed:
                _dialog.ShowError($"途中保存を復元できませんでした。最初から始めます。\n理由: {loaded.Message}\n壊れたファイルは次の場所に残してあります: {loaded.BrokenFilePath}");
                break;
        }
    }

    /// <summary>表のマスが選ばれたとき（画面から呼ばれる）</summary>
    public void MoveTo(int itemIndex, int unitIndex) {
        _session.MoveTo(itemIndex, unitIndex);
        SyncPosition();
        Save();
    }

    [RelayCommand]
    private void CommitValue() {
        if (!decimal.TryParse(InputText, NumberStyles.Float, CultureInfo.CurrentCulture, out var value)) {
            _dialog.ShowError("数値で入力してください。");
            return;
        }
        // 5-1 は手入力のみ。測定器から取得するようになったら、取得値は isManual: false で渡す
        _session.SetValue(CurrentItemIndex, CurrentUnitIndex, value, isManual: true);
        FinishCurrentCell();
    }

    [RelayCommand]
    private void ConfirmOk() {
        _session.SetConfirmation(CurrentItemIndex, CurrentUnitIndex, isOk: true);
        FinishCurrentCell();
    }

    [RelayCommand]
    private void ConfirmNg() {
        _session.SetConfirmation(CurrentItemIndex, CurrentUnitIndex, isOk: false);
        FinishCurrentCell();
    }

    [RelayCommand]
    private void CompleteTask() {
        _session.CompleteTask(CurrentItemIndex, CurrentUnitIndex);
        FinishCurrentCell();
    }

    [RelayCommand]
    private void MoveNext() {
        if (_session.MoveNext()) {
            SyncPosition();
            Save();
        }
    }

    [RelayCommand]
    private void MovePrevious() {
        if (_session.MovePrevious()) {
            SyncPosition();
            Save();
        }
    }

    [RelayCommand]
    private void IncreaseUnitCount() {
        if (UnitCount < MaxUnitCount) {
            TryChangeUnitCount(UnitCount + 1);
        }
    }

    [RelayCommand]
    private void DecreaseUnitCount() => TryChangeUnitCount(UnitCount - 1);

    /// <summary>直接入力された台数を反映する（Enter キーの確定、入力欄からフォーカスが外れたとき）</summary>
    [RelayCommand]
    private void ApplyUnitCountText() {
        if (!int.TryParse(UnitCountText, NumberStyles.None, CultureInfo.CurrentCulture, out var unitCount) || unitCount < 1) {
            // 数字でない、0 などは、入力前の台数に戻す
            UnitCountText = UnitCount.ToString();
            return;
        }
        if (unitCount > MaxUnitCount) {
            _dialog.ShowError($"{Product} の台数は最大 {MaxUnitCount} 台までです。");
            UnitCountText = UnitCount.ToString();
            return;
        }
        TryChangeUnitCount(unitCount);
    }

    // 台数を変更する。結果のある台が削除されるときは確認し、やめた場合や台数が変わらない場合は表示を元に戻す
    private void TryChangeUnitCount(int unitCount) {
        if (unitCount >= 1 && unitCount != UnitCount) {
            // 台数を減らすと、末尾の台から削除される
            var firstRemoved = unitCount;
            var hasResult = unitCount < UnitCount && Enumerable.Range(firstRemoved, UnitCount - firstRemoved).Any(_session.HasAnyResult);
            if (!hasResult || _dialog.Confirm(CreateRemoveMessage(firstRemoved + 1, UnitCount))) {
                _session.ChangeUnitCount(unitCount);
                Rows = CreateRows();
                OnPropertyChanged(nameof(UnitCount));
                SyncPosition();
                Save();
            }
        }
        UnitCountText = UnitCount.ToString();
    }

    private static string CreateRemoveMessage(int firstUnitNumber, int lastUnitNumber) {
        var range = firstUnitNumber == lastUnitNumber ? $"{firstUnitNumber}台目" : $"{firstUnitNumber}台目～{lastUnitNumber}台目";
        return $"{range}の結果が削除されます。台数を減らしますか？";
    }

    // 確定した後の共通処理: マスを更新し、保存して、次の項目へ進む
    private void FinishCurrentCell() {
        Rows[CurrentItemIndex].Cells[CurrentUnitIndex].Refresh();
        _session.MoveNext();
        SyncPosition();
        Save();
    }

    // セッションの現在位置を画面用のプロパティへ反映し、入力欄を移動先の値に合わせる
    private void SyncPosition() {
        CurrentItemIndex = _session.CurrentItemIndex;
        CurrentUnitIndex = _session.CurrentUnitIndex;
        InputText = _session.GetResult(CurrentItemIndex, CurrentUnitIndex).Value?.ToString() ?? "";
    }

    private void ReplaceSession(InspectionSession session) {
        _session = session;
        Rows = CreateRows();
        OnPropertyChanged(nameof(UnitCount));
        UnitCountText = UnitCount.ToString();
        SyncPosition();
    }

    private List<RowViewModel> CreateRows() {
        var rows = new List<RowViewModel>();
        for (var itemIndex = 0; itemIndex < _session.Items.Count; itemIndex++) {
            var cells = new List<CellViewModel>();
            for (var unitIndex = 0; unitIndex < _session.UnitCount; unitIndex++) {
                cells.Add(new CellViewModel(_session, itemIndex, unitIndex));
            }
            rows.Add(new RowViewModel(_session.Items[itemIndex].Name, cells));
        }
        return rows;
    }

    // 保存に失敗しても検査は続けられるように、例外は通知にとどめる
    private void Save() {
        try {
            _store.Save(Product, _session);
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            _dialog.ShowError($"途中保存に失敗しました。検査は続けられますが、アプリが終了すると結果が失われます。\n{ex.Message}");
        }
    }

    private void DeleteSavedState() {
        try {
            _store.Delete(Product);
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            _dialog.ShowError($"途中保存を削除できませんでした。\n{ex.Message}");
        }
    }
}
