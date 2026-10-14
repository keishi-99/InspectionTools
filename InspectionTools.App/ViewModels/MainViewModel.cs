using System.IO;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using InspectionTools.Excel;
using InspectionTools.Inspection;

namespace InspectionTools.App.ViewModels;

/// <summary>検査画面の状態と操作。処理は InspectionSession と SessionStore を呼ぶだけにして、ここには画面の部品を持ち込まない</summary>
public sealed partial class MainViewModel : ObservableObject {

    private readonly ProductDefinition _definition;
    private readonly SessionStore _store;
    private readonly IDialogService _dialog;
    private InspectionSession _session;

    /// <param name="instrumentSelection">測定器の選択と接続（測定器を使わない起動では null。その場合、測定器の操作は出さない）</param>
    public MainViewModel(ProductDefinition definition, SessionStore store, IDialogService dialog, InstrumentSelectionViewModel? instrumentSelection = null) {
        _definition = definition;
        _store = store;
        _dialog = dialog;
        InstrumentSelection = instrumentSelection;
        if (instrumentSelection is not null) {
            // 接続の状態が変わったら、検査画面の上部の表示も更新する
            instrumentSelection.PropertyChanged += (_, e) => {
                if (e.PropertyName == nameof(InstrumentSelectionViewModel.StatusText)) {
                    OnPropertyChanged(nameof(InstrumentStatusText));
                }
            };
        }
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

    /// <summary>測定器の選択と接続（測定器を使わない起動では null）</summary>
    public InstrumentSelectionViewModel? InstrumentSelection { get; }

    /// <summary>測定器の操作を出すか</summary>
    public bool HasInstrumentSelection => InstrumentSelection is not null;

    /// <summary>測定器の接続の状態の表示</summary>
    public string InstrumentStatusText => InstrumentSelection?.StatusText ?? "";

    /// <summary>測定器に接続しているか</summary>
    public bool IsInstrumentConnected => InstrumentSelection?.IsConnected ?? false;

    /// <summary>測定器の選択と接続の画面を開く</summary>
    [RelayCommand]
    private void OpenInstrumentSelection() {
        if (InstrumentSelection is null) {
            return;
        }
        InstrumentSelection.Refresh();
        _dialog.ShowInstrumentSelection(InstrumentSelection);
    }

    /// <summary>アプリを閉じる前に呼ぶ。測定器に接続していれば、解除する（測定器には何も送らない）</summary>
    public void Shutdown() {
        InstrumentSelection?.Disconnect();
    }

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

    /// <summary>検査を終了し、転記先のテンプレートを選んでもらって、結果を転記する</summary>
    [RelayCommand]
    private void FinishInspection() {
        var templatePath = _dialog.SelectTemplateFilePath();
        if (templatePath is null) {
            return;
        }

        var settings = TransferSettings.From(_definition.Transfer);
        var unitCount = UnitCount;
        var isSavedStateKept = false;
        try {
            // 違う製品のテンプレートで合格範囲を突き合わせても意味がないので、先に製品を確認する
            TemplateVerifier.Verify(templatePath, settings.Check);
            var templateRows = TemplateReader.ReadRows(templatePath, TemplateLayout.From(_definition.Transfer));
            var mismatches = RangeComparer.Compare(_session.Items, templateRows);
            if (mismatches.Count > 0 && !_dialog.Confirm(CreateMismatchMessage(mismatches))) {
                return;
            }

            // 未実施のセルは、転記先の値を消して空欄にする（上書きの確認は不要という要件）ため、消えることを先に知らせる
            var undoneCount = CountUndoneCells();
            if (undoneCount > 0 && !_dialog.Confirm($"未実施の項目が {undoneCount} か所あります。転記先の該当セルは空欄になります。転記しますか？")) {
                return;
            }

            ResultTransfer.Transfer(templatePath, settings, _session, _store, Product);
        } catch (Exception ex) when (ex is TemplateReadException or TemplateVerificationException or TemplateWriteException) {
            // 途中保存は残っているので、原因を直してもう一度押せばやり直せる
            _dialog.ShowError($"転記できませんでした。\n{ex.Message}\n\n検査結果は途中保存に残してあります。原因を直してから、もう一度「検査終了」を押してください。");
            return;
        } catch (SavedStateDeleteException ex) {
            isSavedStateKept = true;
            _dialog.ShowError($"{ex.Message}\n\n次回の起動で「続きから再開しますか？」と出たときは、「いいえ」を選んでください（転記は済んでいます）。");
        }

        if (!isSavedStateKept) {
            _dialog.ShowInfo($"{unitCount}台分の結果を転記しました。\n{templatePath}");
        }
        // 転記が済んだ結果を残したままにすると、続けて操作したときに途中保存が作り直されるため、新しい検査にする
        ReplaceSession(new InspectionSession(_definition.Items, unitCount: 1));
    }

    // 合格範囲の食い違いを、確認ダイアログに出す文面にする（多いときは先頭の数件だけ）
    private static string CreateMismatchMessage(IReadOnlyList<RangeMismatch> mismatches) {
        const int MaxLines = 10;
        var lines = mismatches.Take(MaxLines).Select(mismatch => $"・{mismatch.Row}行 {mismatch.ItemName}: {mismatch.Reason}");
        var omitted = mismatches.Count > MaxLines ? $"\n（ほか {mismatches.Count - MaxLines} 件）" : "";
        return $"アプリの定義と、テンプレートの合格範囲が一致しません。\n{string.Join("\n", lines)}{omitted}\n\nこのまま転記しますか？（Excel の判定は、テンプレート側の合格範囲で行われます）";
    }

    // 転記する項目（作業項目以外）のうち、まだ実施していないセルの数を数える
    private int CountUndoneCells() {
        var count = 0;
        for (var unitIndex = 0; unitIndex < _session.UnitCount; unitIndex++) {
            for (var itemIndex = 0; itemIndex < _session.Items.Count; itemIndex++) {
                var isRecorded = _session.Items[itemIndex].Kind != ItemKind.Task;
                if (isRecorded && _session.GetResult(itemIndex, unitIndex).State == ResultState.NotDone) {
                    count++;
                }
            }
        }
        return count;
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
