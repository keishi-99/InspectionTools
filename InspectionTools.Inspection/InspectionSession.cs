namespace InspectionTools.Inspection;

/// <summary>検査中の状態（台数、項目 × 台の結果、現在位置）を持つ</summary>
public sealed class InspectionSession {

    // 台ごとの結果の並び（_units[台][項目]）
    private readonly List<ItemResult[]> _units = [];

    public InspectionSession(IReadOnlyList<InspectionItem> items, int unitCount) {
        if (items.Count == 0) {
            throw new ArgumentException("項目が1つもありません。", nameof(items));
        }
        ValidateUnitCount(unitCount);

        Items = items;
        for (var i = 0; i < unitCount; i++) {
            _units.Add(CreateEmptyUnit());
        }
    }

    /// <summary>検査項目の定義</summary>
    public IReadOnlyList<InspectionItem> Items { get; }

    /// <summary>台数</summary>
    public int UnitCount => _units.Count;

    /// <summary>現在の項目の位置</summary>
    public int CurrentItemIndex { get; private set; }

    /// <summary>現在の台の位置</summary>
    public int CurrentUnitIndex { get; private set; }

    /// <summary>指定したセル（項目 × 台）の結果を返す</summary>
    public ItemResult GetResult(int itemIndex, int unitIndex) {
        ValidatePosition(itemIndex, unitIndex);
        return _units[unitIndex][itemIndex];
    }

    /// <summary>値を設定し、合格範囲と比べて判定する（測定項目と入力項目のみ。再測定は上書き）</summary>
    public void SetValue(int itemIndex, int unitIndex, decimal value, bool isManual) {
        ValidatePosition(itemIndex, unitIndex);
        var item = Items[itemIndex];
        if (item.Kind is not (ItemKind.Measurement or ItemKind.Input)) {
            throw new InvalidOperationException($"値を設定できない項目です: {item.Name}（{item.Kind}）");
        }

        var result = _units[unitIndex][itemIndex];
        result.Value = value;
        result.IsManual = isManual;
        // 合格範囲は両端を含む（テンプレートの判定式 下限 <= 値 <= 上限 と同じ）
        result.State = item.LowerLimit <= value && value <= item.UpperLimit ? ResultState.Ok : ResultState.Ng;
    }

    /// <summary>入力した文字を設定し、実施済みにする（文字入力の項目のみ。前後の空白は取り除き、空欄は受け付けない。やり直しは上書き）</summary>
    public void SetText(int itemIndex, int unitIndex, string text) {
        ValidatePosition(itemIndex, unitIndex);
        var item = Items[itemIndex];
        if (item.Kind != ItemKind.Text) {
            throw new InvalidOperationException($"文字を設定できない項目です: {item.Name}（{item.Kind}）");
        }
        if (string.IsNullOrWhiteSpace(text)) {
            throw new ArgumentException($"文字が空です: {item.Name}", nameof(text));
        }

        var result = _units[unitIndex][itemIndex];
        result.Text = text.Trim();
        result.State = ResultState.Ok;
    }

    /// <summary>人が決めた合否を設定する（確認項目のみ）</summary>
    public void SetConfirmation(int itemIndex, int unitIndex, bool isOk) {
        ValidatePosition(itemIndex, unitIndex);
        var item = Items[itemIndex];
        if (item.Kind != ItemKind.Confirmation) {
            throw new InvalidOperationException($"合否を設定できない項目です: {item.Name}（{item.Kind}）");
        }

        _units[unitIndex][itemIndex].State = isOk ? ResultState.Ok : ResultState.Ng;
    }

    /// <summary>作業が完了したことを設定する（作業項目のみ）</summary>
    public void CompleteTask(int itemIndex, int unitIndex) {
        ValidatePosition(itemIndex, unitIndex);
        var item = Items[itemIndex];
        if (item.Kind != ItemKind.Task) {
            throw new InvalidOperationException($"作業項目ではありません: {item.Name}（{item.Kind}）");
        }

        _units[unitIndex][itemIndex].State = ResultState.Done;
    }

    /// <summary>指定した台に、実施済みの項目が1つでもあるか（台数を減らす前の確認ダイアログの判断に使う）</summary>
    public bool HasAnyResult(int unitIndex) {
        ValidatePosition(0, unitIndex);
        return _units[unitIndex].Any(result => result.State != ResultState.NotDone);
    }

    /// <summary>台数を変更する。増やした台は未実施で追加し、減らすときは末尾の台から削除する</summary>
    public void ChangeUnitCount(int unitCount) {
        ValidateUnitCount(unitCount);

        while (_units.Count < unitCount) {
            _units.Add(CreateEmptyUnit());
        }
        if (_units.Count > unitCount) {
            _units.RemoveRange(unitCount, _units.Count - unitCount);
        }

        // 現在の台が削除された場合は、最後の台に移す
        CurrentUnitIndex = Math.Min(CurrentUnitIndex, unitCount - 1);
    }

    /// <summary>任意のセル（項目 × 台）へ移動する</summary>
    public void MoveTo(int itemIndex, int unitIndex) {
        ValidatePosition(itemIndex, unitIndex);
        CurrentItemIndex = itemIndex;
        CurrentUnitIndex = unitIndex;
    }

    /// <summary>台ごとに進める順で、次のセルへ移動する。最後のセルのときは移動せず false を返す</summary>
    public bool MoveNext() {
        if (CurrentItemIndex < Items.Count - 1) {
            CurrentItemIndex++;
            return true;
        }
        if (CurrentUnitIndex < UnitCount - 1) {
            CurrentUnitIndex++;
            CurrentItemIndex = 0;
            return true;
        }
        return false;
    }

    /// <summary>台ごとに進める順で、前のセルへ移動する。最初のセルのときは移動せず false を返す</summary>
    public bool MovePrevious() {
        if (CurrentItemIndex > 0) {
            CurrentItemIndex--;
            return true;
        }
        if (CurrentUnitIndex > 0) {
            CurrentUnitIndex--;
            CurrentItemIndex = Items.Count - 1;
            return true;
        }
        return false;
    }

    /// <summary>途中保存用に、実施済みのセルだけを書き出す</summary>
    public SessionSnapshot ToSnapshot(string product) {
        var cells = new List<CellSnapshot>();
        for (var unitIndex = 0; unitIndex < UnitCount; unitIndex++) {
            for (var itemIndex = 0; itemIndex < Items.Count; itemIndex++) {
                var result = _units[unitIndex][itemIndex];
                if (result.State == ResultState.NotDone) {
                    continue;
                }
                cells.Add(new CellSnapshot(unitIndex, itemIndex, result.Value, result.State, result.IsManual, result.Text));
            }
        }
        return new SessionSnapshot(product, Items.Count, UnitCount, CurrentItemIndex, CurrentUnitIndex, cells);
    }

    /// <summary>途中保存の内容から、検査中の状態を復元する。壊れている、または項目の定義と合わないときは SessionRestoreException</summary>
    public static InspectionSession FromSnapshot(SessionSnapshot snapshot, string product, IReadOnlyList<InspectionItem> items) {
        ValidateSnapshot(snapshot, product, items);

        var session = new InspectionSession(items, snapshot.UnitCount);
        foreach (var cell in snapshot.Cells) {
            session.RestoreCell(cell);
        }
        session.MoveTo(snapshot.CurrentItemIndex, snapshot.CurrentUnitIndex);
        return session;
    }

    // 保存内容が、いまの項目の定義と合っているか確認する（合わなければ SessionRestoreException を投げる）
    private static void ValidateSnapshot(SessionSnapshot snapshot, string product, IReadOnlyList<InspectionItem> items) {
        if (snapshot.Product != product) {
            throw new SessionRestoreException($"製品が違います: {snapshot.Product}（保存時） / {product}（現在）");
        }
        if (snapshot.ItemCount != items.Count) {
            throw new SessionRestoreException($"項目数が違います: {snapshot.ItemCount}（保存時） / {items.Count}（現在）");
        }
        if (snapshot.UnitCount < 1) {
            throw new SessionRestoreException($"台数が不正です: {snapshot.UnitCount}（保存時）");
        }
        if (snapshot.CurrentItemIndex < 0 || snapshot.CurrentItemIndex >= items.Count) {
            throw new SessionRestoreException($"項目の位置が範囲外です: {snapshot.CurrentItemIndex}（保存時） / {items.Count}（現在）");
        }
        if (snapshot.CurrentUnitIndex < 0 || snapshot.CurrentUnitIndex >= snapshot.UnitCount) {
            throw new SessionRestoreException($"台の位置が範囲外です: {snapshot.CurrentUnitIndex}（保存時） / {snapshot.UnitCount}（現在）");
        }
        foreach (var cell in snapshot.Cells) {
            if (cell.Item < 0 || cell.Item >= items.Count) {
                throw new SessionRestoreException($"項目の位置が範囲外です: {cell.Item}（保存時） / {items.Count}（現在）");
            }
            if (cell.Unit < 0 || cell.Unit >= snapshot.UnitCount) {
                throw new SessionRestoreException($"台の位置が範囲外です: {cell.Unit}（保存時） / {snapshot.UnitCount}（現在）");
            }
        }
    }

    // 保存された1セルを、通常の設定処理を通して復元する（値の項目は、定義が変わっていても合格範囲で判定し直すため）
    private void RestoreCell(CellSnapshot cell) {
        if (cell.State == ResultState.NotDone) {
            return;
        }

        var item = Items[cell.Item];
        switch (item.Kind) {
            case ItemKind.Measurement:
            case ItemKind.Input:
                if (cell.Value is null) {
                    throw new SessionRestoreException($"値がありません: {item.Name}（台 {cell.Unit + 1}）");
                }
                SetValue(cell.Item, cell.Unit, cell.Value.Value, cell.IsManual);
                break;
            case ItemKind.Text when cell.State == ResultState.Ok:
                if (string.IsNullOrWhiteSpace(cell.Text)) {
                    throw new SessionRestoreException($"文字がありません: {item.Name}（台 {cell.Unit + 1}）");
                }
                SetText(cell.Item, cell.Unit, cell.Text);
                break;
            case ItemKind.Confirmation when cell.State is ResultState.Ok or ResultState.Ng:
                SetConfirmation(cell.Item, cell.Unit, cell.State == ResultState.Ok);
                break;
            case ItemKind.Task when cell.State == ResultState.Done:
                CompleteTask(cell.Item, cell.Unit);
                break;
            default:
                throw new SessionRestoreException($"項目の種類と状態が合いません: {item.Name}（{item.Kind}、{cell.State}）");
        }
    }

    // 全項目が未実施の1台分の結果を作る
    private ItemResult[] CreateEmptyUnit() => Enumerable.Range(0, Items.Count).Select(_ => new ItemResult()).ToArray();

    // 台数は1以上でなければならない
    private static void ValidateUnitCount(int unitCount) {
        ArgumentOutOfRangeException.ThrowIfLessThan(unitCount, 1);
    }

    // 項目と台の位置が範囲内か確認する
    private void ValidatePosition(int itemIndex, int unitIndex) {
        ArgumentOutOfRangeException.ThrowIfNegative(itemIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(itemIndex, Items.Count);
        ArgumentOutOfRangeException.ThrowIfNegative(unitIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(unitIndex, UnitCount);
    }
}
