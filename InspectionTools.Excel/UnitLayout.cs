namespace InspectionTools.Excel;

/// <summary>台ごとの結果をテンプレートのどの列・行に書くか。台数が多いときに、同じ列のまま下の行へ折り返す製品にも対応する</summary>
public sealed record UnitLayout {

    public UnitLayout(string firstUnitColumn, int? unitsPerBlock = null, int blockRowOffset = 0, string? confirmationColumn = null) {
        FirstUnitColumn = ToColumnNumber(firstUnitColumn);
        ConfirmationColumn = confirmationColumn is null ? null : ToColumnNumber(confirmationColumn);
        if (unitsPerBlock is < 1) {
            throw new ArgumentOutOfRangeException(nameof(unitsPerBlock), "折り返す台数は1以上にしてください。");
        }
        if (unitsPerBlock is not null && blockRowOffset < 1) {
            throw new ArgumentOutOfRangeException(nameof(blockRowOffset), "折り返すときは、下の行までの行数（1以上）が必要です。");
        }

        UnitsPerBlock = unitsPerBlock;
        BlockRowOffset = blockRowOffset;
    }

    /// <summary>1台目を書く列の番号（A列が 1）</summary>
    public int FirstUnitColumn { get; }

    /// <summary>確認項目の1台目を書く列の番号（チェックボックスのリンク先など。値の列と同じセルに書く製品は null）</summary>
    public int? ConfirmationColumn { get; }

    /// <summary>折り返すまでの台数（折り返さない製品は null）</summary>
    public int? UnitsPerBlock { get; }

    /// <summary>次のまとまりへ移るときに、下へ進める行数</summary>
    public int BlockRowOffset { get; }

    /// <summary>項目の行と台の位置（0 から数える）から、書き込むセルの位置を求める</summary>
    public CellPosition GetCell(int itemRow, int unitIndex) => GetCell(FirstUnitColumn, itemRow, unitIndex);

    /// <summary>確認項目の結果を書くセルの位置を求める。確認の列がない製品は、値のセルと同じ</summary>
    public CellPosition GetConfirmationCell(int itemRow, int unitIndex) => GetCell(ConfirmationColumn ?? FirstUnitColumn, itemRow, unitIndex);

    // 1台目の列を指定して、台の位置からセルを求める（値の列と確認の列で、同じ折り返しの規則を使う）
    private CellPosition GetCell(int firstColumn, int itemRow, int unitIndex) {
        ArgumentOutOfRangeException.ThrowIfNegative(unitIndex);

        if (UnitsPerBlock is null) {
            return new CellPosition(itemRow, firstColumn + unitIndex);
        }
        else {
            var blockIndex = unitIndex / UnitsPerBlock.Value;
            var blockOffset = unitIndex % UnitsPerBlock.Value;
            return new CellPosition(itemRow + blockIndex * BlockRowOffset, firstColumn + blockOffset);
        }
    }

    // 「F」「AI」のような列の文字を、列の番号（A が 1）にする
    private static int ToColumnNumber(string letters) {
        if (string.IsNullOrEmpty(letters) || !letters.All(char.IsAsciiLetter)) {
            throw new ArgumentException($"列の指定が正しくありません: '{letters}'（例: F, AI）", nameof(letters));
        }

        var number = 0;
        foreach (var letter in letters.ToUpperInvariant()) {
            number = number * 26 + (letter - 'A' + 1);
        }
        return number;
    }
}
