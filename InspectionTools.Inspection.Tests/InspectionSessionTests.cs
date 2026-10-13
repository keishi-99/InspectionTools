namespace InspectionTools.Inspection.Tests;

/// <summary>検査中の状態（値の設定と判定、台数の変更、現在位置の移動）が要件どおりに動くことを確認する</summary>
public class InspectionSessionTests {

    // 項目の位置（0: 測定、1: 入力、2: 確認、3: 作業）
    private const int MeasurementIndex = 0;
    private const int InputIndex = 1;
    private const int ConfirmationIndex = 2;
    private const int TaskIndex = 3;

    private static InspectionSession CreateSession(int unitCount = 2) {
        List<InspectionItem> items = [
            new("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m),
            new("温度入力", ItemKind.Input, excelRow: 16, lowerLimit: -50.2m, upperLimit: -49.8m),
            new("電源投入", ItemKind.Confirmation, excelRow: 5),
            new("調整作業", ItemKind.Task),
        ];
        return new InspectionSession(items, unitCount);
    }

    [Fact]
    public void AllCellsAreNotDoneAtStart() {
        var session = CreateSession(unitCount: 2);

        for (var unit = 0; unit < session.UnitCount; unit++) {
            for (var item = 0; item < session.Items.Count; item++) {
                Assert.Equal(ResultState.NotDone, session.GetResult(item, unit).State);
            }
        }
    }

    [Fact]
    public void ConstructorRejectsEmptyItemsAndZeroUnits() {
        Assert.Throws<ArgumentException>(() => new InspectionSession([], 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => CreateSession(unitCount: 0));
    }

    [Theory]
    [InlineData("146.745", true)]
    [InlineData("0", true)]
    [InlineData("230", true)]
    [InlineData("230.001", false)]
    [InlineData("-0.001", false)]
    public void SetValueJudgesWithInclusiveRange(string text, bool expectedOk) {
        var session = CreateSession();

        session.SetValue(MeasurementIndex, 0, decimal.Parse(text), isManual: false);

        Assert.Equal(expectedOk ? ResultState.Ok : ResultState.Ng, session.GetResult(MeasurementIndex, 0).State);
    }

    [Fact]
    public void SetValueKeepsValueAndManualFlag() {
        var session = CreateSession();

        session.SetValue(InputIndex, 1, -50.003m, isManual: true);

        var result = session.GetResult(InputIndex, 1);
        Assert.Equal(-50.003m, result.Value);
        Assert.True(result.IsManual);
        Assert.Equal(ResultState.Ok, result.State);
    }

    [Fact]
    public void SetValueOverwritesAndRejudges() {
        var session = CreateSession();
        session.SetValue(MeasurementIndex, 0, 300m, isManual: false);

        session.SetValue(MeasurementIndex, 0, 100m, isManual: true);

        var result = session.GetResult(MeasurementIndex, 0);
        Assert.Equal(100m, result.Value);
        Assert.Equal(ResultState.Ok, result.State);
        Assert.True(result.IsManual);
    }

    [Fact]
    public void SetValueOnConfirmationOrTaskThrows() {
        var session = CreateSession();

        Assert.Throws<InvalidOperationException>(() => session.SetValue(ConfirmationIndex, 0, 1m, isManual: false));
        Assert.Throws<InvalidOperationException>(() => session.SetValue(TaskIndex, 0, 1m, isManual: false));
    }

    [Theory]
    [InlineData(true, ResultState.Ok)]
    [InlineData(false, ResultState.Ng)]
    public void SetConfirmationSetsState(bool isOk, ResultState expected) {
        var session = CreateSession();

        session.SetConfirmation(ConfirmationIndex, 0, isOk);

        var result = session.GetResult(ConfirmationIndex, 0);
        Assert.Equal(expected, result.State);
        Assert.Null(result.Value);
    }

    [Fact]
    public void SetConfirmationOnOtherKindsThrows() {
        var session = CreateSession();

        Assert.Throws<InvalidOperationException>(() => session.SetConfirmation(MeasurementIndex, 0, true));
        Assert.Throws<InvalidOperationException>(() => session.SetConfirmation(TaskIndex, 0, true));
    }

    [Fact]
    public void CompleteTaskSetsDoneOnlyForTaskItems() {
        var session = CreateSession();

        session.CompleteTask(TaskIndex, 0);

        Assert.Equal(ResultState.Done, session.GetResult(TaskIndex, 0).State);
        Assert.Throws<InvalidOperationException>(() => session.CompleteTask(MeasurementIndex, 0));
    }

    [Fact]
    public void OutOfRangePositionThrows() {
        var session = CreateSession(unitCount: 2);

        Assert.Throws<ArgumentOutOfRangeException>(() => session.GetResult(4, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => session.GetResult(0, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => session.GetResult(-1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => session.MoveTo(0, 2));
    }

    [Fact]
    public void HasAnyResultIsTrueOnlyForUnitsWithResults() {
        var session = CreateSession(unitCount: 2);
        session.CompleteTask(TaskIndex, 1);

        Assert.False(session.HasAnyResult(0));
        Assert.True(session.HasAnyResult(1));
    }

    [Fact]
    public void IncreasingUnitCountKeepsExistingResultsAndAddsEmptyUnits() {
        var session = CreateSession(unitCount: 1);
        session.SetValue(MeasurementIndex, 0, 100m, isManual: false);

        session.ChangeUnitCount(3);

        Assert.Equal(3, session.UnitCount);
        Assert.Equal(100m, session.GetResult(MeasurementIndex, 0).Value);
        Assert.Equal(ResultState.NotDone, session.GetResult(MeasurementIndex, 2).State);
    }

    [Fact]
    public void DecreasingUnitCountRemovesLastUnitsAndMovesCurrentPosition() {
        var session = CreateSession(unitCount: 3);
        session.MoveTo(1, 2);

        session.ChangeUnitCount(2);

        Assert.Equal(2, session.UnitCount);
        Assert.Equal(1, session.CurrentUnitIndex);
        Assert.Equal(1, session.CurrentItemIndex);
        Assert.Throws<ArgumentOutOfRangeException>(() => session.GetResult(0, 2));
    }

    [Fact]
    public void ChangeUnitCountToZeroThrows() {
        var session = CreateSession();

        Assert.Throws<ArgumentOutOfRangeException>(() => session.ChangeUnitCount(0));
    }

    [Fact]
    public void MoveToJumpsToAnyCell() {
        var session = CreateSession(unitCount: 2);

        session.MoveTo(itemIndex: 3, unitIndex: 1);

        Assert.Equal(3, session.CurrentItemIndex);
        Assert.Equal(1, session.CurrentUnitIndex);
    }

    [Fact]
    public void MoveNextGoesToNextItemInTheSameUnit() {
        var session = CreateSession(unitCount: 2);

        var moved = session.MoveNext();

        Assert.True(moved);
        Assert.Equal(1, session.CurrentItemIndex);
        Assert.Equal(0, session.CurrentUnitIndex);
    }

    [Fact]
    public void MoveNextAtLastItemGoesToFirstItemOfNextUnit() {
        var session = CreateSession(unitCount: 2);
        session.MoveTo(TaskIndex, 0);

        var moved = session.MoveNext();

        Assert.True(moved);
        Assert.Equal(0, session.CurrentItemIndex);
        Assert.Equal(1, session.CurrentUnitIndex);
    }

    [Fact]
    public void MoveNextAtLastCellStaysAndReturnsFalse() {
        var session = CreateSession(unitCount: 2);
        session.MoveTo(TaskIndex, 1);

        var moved = session.MoveNext();

        Assert.False(moved);
        Assert.Equal(TaskIndex, session.CurrentItemIndex);
        Assert.Equal(1, session.CurrentUnitIndex);
    }

    [Fact]
    public void MovePreviousGoesBackInTheSameUnit() {
        var session = CreateSession(unitCount: 2);
        session.MoveTo(2, 0);

        var moved = session.MovePrevious();

        Assert.True(moved);
        Assert.Equal(1, session.CurrentItemIndex);
        Assert.Equal(0, session.CurrentUnitIndex);
    }

    [Fact]
    public void MovePreviousAtFirstItemGoesToLastItemOfPreviousUnit() {
        var session = CreateSession(unitCount: 2);
        session.MoveTo(0, 1);

        var moved = session.MovePrevious();

        Assert.True(moved);
        Assert.Equal(TaskIndex, session.CurrentItemIndex);
        Assert.Equal(0, session.CurrentUnitIndex);
    }

    [Fact]
    public void MovePreviousAtFirstCellStaysAndReturnsFalse() {
        var session = CreateSession(unitCount: 2);

        var moved = session.MovePrevious();

        Assert.False(moved);
        Assert.Equal(0, session.CurrentItemIndex);
        Assert.Equal(0, session.CurrentUnitIndex);
    }
}
