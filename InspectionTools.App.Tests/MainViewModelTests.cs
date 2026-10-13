using InspectionTools.App.ViewModels;
using InspectionTools.Inspection;

namespace InspectionTools.App.Tests;

/// <summary>検査画面の ViewModel の動き（入力、判定、移動、台数の変更、途中保存と復元）を、画面なしで確認する</summary>
public sealed class MainViewModelTests : IDisposable {

    private const string Product = "EL5000AC";
    private const int MaxUnitCount = 30;

    // 確認の答えを順番に返し、通知された内容を記録するダミー
    private sealed class FakeDialogService : IDialogService {
        public Queue<bool> ConfirmAnswers { get; } = new();
        public List<string> ConfirmMessages { get; } = [];
        public List<string> Errors { get; } = [];

        public bool Confirm(string message) {
            ConfirmMessages.Add(message);
            return ConfirmAnswers.Dequeue();
        }

        public void ShowError(string message) => Errors.Add(message);
    }

    private readonly string _directory = Path.Combine(Path.GetTempPath(), "InspectionToolsTests", Guid.NewGuid().ToString("N"));
    private readonly FakeDialogService _dialog = new();
    private readonly SessionStore _store;

    public MainViewModelTests() {
        _store = new SessionStore(_directory);
    }

    public void Dispose() {
        if (Directory.Exists(_directory)) {
            Directory.Delete(_directory, recursive: true);
        }
        // 保存の失敗を再現するテストでは、フォルダの代わりにファイルを置くので、そちらも片付ける
        if (File.Exists(_directory)) {
            File.Delete(_directory);
        }
    }

    private string SavedPath => Path.Combine(_directory, Product + ".json");

    // 0: 測定、1: 入力、2: 確認、3: 作業 の4項目
    private static ProductDefinition CreateDefinition() => new(Product, MaxUnitCount, [
        new InspectionItem("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m),
        new InspectionItem("温度入力", ItemKind.Input, excelRow: 16, lowerLimit: -50.2m, upperLimit: -49.8m),
        new InspectionItem("電源投入", ItemKind.Confirmation, excelRow: 5),
        new InspectionItem("調整作業", ItemKind.Task),
    ]);

    private MainViewModel CreateViewModel() => new(CreateDefinition(), _store, _dialog);

    [Fact]
    public void StartsWithOneUnitAtTheFirstCell() {
        var viewModel = CreateViewModel();

        Assert.Equal(1, viewModel.UnitCount);
        Assert.Equal(0, viewModel.CurrentItemIndex);
        Assert.Equal(0, viewModel.CurrentUnitIndex);
        Assert.Equal(4, viewModel.Rows.Count);
        Assert.True(viewModel.IsValueInput);
    }

    [Fact]
    public void CommittingValueInRangeMarksOkShowsValueAndMovesToNextItem() {
        var viewModel = CreateViewModel();
        viewModel.InputText = "146.745";

        viewModel.CommitValueCommand.Execute(null);

        var cell = viewModel.Rows[0].Cells[0];
        Assert.Equal(ResultState.Ok, cell.State);
        Assert.Equal("146.745", cell.Display);
        Assert.Equal(1, viewModel.CurrentItemIndex);
        Assert.Equal("", viewModel.InputText);
    }

    [Fact]
    public void CommittingValueOutOfRangeMarksNg() {
        var viewModel = CreateViewModel();
        viewModel.InputText = "231";

        viewModel.CommitValueCommand.Execute(null);

        Assert.Equal(ResultState.Ng, viewModel.Rows[0].Cells[0].State);
    }

    [Theory]
    [InlineData("")]
    [InlineData("abc")]
    public void CommittingNonNumberShowsErrorAndChangesNothing(string text) {
        var viewModel = CreateViewModel();
        viewModel.InputText = text;

        viewModel.CommitValueCommand.Execute(null);

        Assert.Single(_dialog.Errors);
        Assert.Equal(ResultState.NotDone, viewModel.Rows[0].Cells[0].State);
        Assert.Equal(0, viewModel.CurrentItemIndex);
    }

    [Fact]
    public void ConfirmationAndTaskAreSetByTheirOwnCommands() {
        var viewModel = CreateViewModel();
        viewModel.MoveTo(2, 0);
        Assert.True(viewModel.IsConfirmation);

        viewModel.ConfirmNgCommand.Execute(null);
        Assert.Equal(ResultState.Ng, viewModel.Rows[2].Cells[0].State);
        Assert.Equal("不合格", viewModel.Rows[2].Cells[0].Display);

        Assert.True(viewModel.IsTask);
        viewModel.CompleteTaskCommand.Execute(null);
        Assert.Equal(ResultState.Done, viewModel.Rows[3].Cells[0].State);
        Assert.Equal("完了", viewModel.Rows[3].Cells[0].Display);
    }

    [Fact]
    public void CommittingAtTheLastCellStaysThere() {
        var viewModel = CreateViewModel();
        viewModel.MoveTo(3, 0);

        viewModel.CompleteTaskCommand.Execute(null);

        Assert.Equal(3, viewModel.CurrentItemIndex);
        Assert.Equal(0, viewModel.CurrentUnitIndex);
    }

    [Fact]
    public void MovingToCellWithValueShowsItInTheInputBox() {
        var viewModel = CreateViewModel();
        viewModel.InputText = "100";
        viewModel.CommitValueCommand.Execute(null);

        viewModel.MoveTo(0, 0);

        Assert.Equal("100", viewModel.InputText);
    }

    [Fact]
    public void IncreasingUnitCountAddsAColumnToEveryRow() {
        var viewModel = CreateViewModel();

        viewModel.IncreaseUnitCountCommand.Execute(null);

        Assert.Equal(2, viewModel.UnitCount);
        Assert.All(viewModel.Rows, row => Assert.Equal(2, row.Cells.Count));
    }

    [Fact]
    public void DecreasingUnitCountWithoutResultsDoesNotAsk() {
        var viewModel = CreateViewModel();
        viewModel.IncreaseUnitCountCommand.Execute(null);

        viewModel.DecreaseUnitCountCommand.Execute(null);

        Assert.Equal(1, viewModel.UnitCount);
        Assert.Empty(_dialog.ConfirmMessages);
    }

    [Theory]
    [InlineData(true, 1)]
    [InlineData(false, 2)]
    public void DecreasingUnitCountWithResultsAsksAndFollowsTheAnswer(bool answer, int expectedUnitCount) {
        var viewModel = CreateViewModel();
        viewModel.IncreaseUnitCountCommand.Execute(null);
        viewModel.MoveTo(0, 1);
        viewModel.InputText = "100";
        viewModel.CommitValueCommand.Execute(null);
        _dialog.ConfirmAnswers.Enqueue(answer);

        viewModel.DecreaseUnitCountCommand.Execute(null);

        Assert.Single(_dialog.ConfirmMessages);
        Assert.Equal(expectedUnitCount, viewModel.UnitCount);
    }

    [Fact]
    public void UnitCountTextFollowsTheButtons() {
        var viewModel = CreateViewModel();

        viewModel.IncreaseUnitCountCommand.Execute(null);

        Assert.Equal("2", viewModel.UnitCountText);
    }

    [Fact]
    public void AppliedUnitCountTextChangesTheUnitCountAtOnce() {
        var viewModel = CreateViewModel();
        viewModel.UnitCountText = "5";

        viewModel.ApplyUnitCountTextCommand.Execute(null);

        Assert.Equal(5, viewModel.UnitCount);
        Assert.All(viewModel.Rows, row => Assert.Equal(5, row.Cells.Count));
    }

    [Theory]
    [InlineData("")]
    [InlineData("abc")]
    [InlineData("0")]
    [InlineData("-3")]
    [InlineData("2.5")]
    public void InvalidUnitCountTextIsRevertedWithoutMessage(string text) {
        var viewModel = CreateViewModel();
        viewModel.UnitCountText = text;

        viewModel.ApplyUnitCountTextCommand.Execute(null);

        Assert.Equal(1, viewModel.UnitCount);
        Assert.Equal("1", viewModel.UnitCountText);
        Assert.Empty(_dialog.Errors);
    }

    [Fact]
    public void UnitCountTextOverTheMaximumShowsErrorAndIsReverted() {
        var viewModel = CreateViewModel();
        viewModel.UnitCountText = (MaxUnitCount + 1).ToString();

        viewModel.ApplyUnitCountTextCommand.Execute(null);

        Assert.Single(_dialog.Errors);
        Assert.Equal(1, viewModel.UnitCount);
        Assert.Equal("1", viewModel.UnitCountText);
    }

    [Fact]
    public void IncreaseButtonStopsAtTheMaximum() {
        var viewModel = CreateViewModel();
        viewModel.UnitCountText = MaxUnitCount.ToString();
        viewModel.ApplyUnitCountTextCommand.Execute(null);

        viewModel.IncreaseUnitCountCommand.Execute(null);

        Assert.Equal(MaxUnitCount, viewModel.UnitCount);
        Assert.Empty(_dialog.Errors);
    }

    [Fact]
    public void DecreasingByTextAsksOnceWithTheRemovedRangeWhenAnyHasResults() {
        var viewModel = CreateViewModel();
        viewModel.UnitCountText = "4";
        viewModel.ApplyUnitCountTextCommand.Execute(null);
        viewModel.MoveTo(0, 2);
        viewModel.InputText = "100";
        viewModel.CommitValueCommand.Execute(null);
        _dialog.ConfirmAnswers.Enqueue(true);

        viewModel.UnitCountText = "2";
        viewModel.ApplyUnitCountTextCommand.Execute(null);

        Assert.Equal("3台目～4台目の結果が削除されます。台数を減らしますか？", Assert.Single(_dialog.ConfirmMessages));
        Assert.Equal(2, viewModel.UnitCount);
    }

    [Fact]
    public void DecliningTheConfirmationKeepsTheUnitCountAndRestoresTheText() {
        var viewModel = CreateViewModel();
        viewModel.UnitCountText = "3";
        viewModel.ApplyUnitCountTextCommand.Execute(null);
        viewModel.MoveTo(0, 2);
        viewModel.InputText = "100";
        viewModel.CommitValueCommand.Execute(null);
        _dialog.ConfirmAnswers.Enqueue(false);

        viewModel.UnitCountText = "1";
        viewModel.ApplyUnitCountTextCommand.Execute(null);

        Assert.Equal(3, viewModel.UnitCount);
        Assert.Equal("3", viewModel.UnitCountText);
        Assert.Equal("100", viewModel.Rows[0].Cells[2].Display);
    }

    [Fact]
    public void DecreasingByTextDoesNotAskWhenOnlyEmptyUnitsAreRemoved() {
        var viewModel = CreateViewModel();
        viewModel.UnitCountText = "5";
        viewModel.ApplyUnitCountTextCommand.Execute(null);

        viewModel.UnitCountText = "2";
        viewModel.ApplyUnitCountTextCommand.Execute(null);

        Assert.Equal(2, viewModel.UnitCount);
        Assert.Empty(_dialog.ConfirmMessages);
    }

    [Fact]
    public void UnitCountCannotGoBelowOne() {
        var viewModel = CreateViewModel();

        viewModel.DecreaseUnitCountCommand.Execute(null);

        Assert.Equal(1, viewModel.UnitCount);
    }

    [Fact]
    public void EveryChangeIsSavedAndCanBeRestoredAfterRestart() {
        var viewModel = CreateViewModel();
        viewModel.IncreaseUnitCountCommand.Execute(null);
        viewModel.InputText = "146.745";
        viewModel.CommitValueCommand.Execute(null);
        Assert.True(File.Exists(SavedPath));

        var restarted = CreateViewModel();
        _dialog.ConfirmAnswers.Enqueue(true);
        restarted.Start();

        Assert.Equal(2, restarted.UnitCount);
        Assert.Equal("2", restarted.UnitCountText);
        Assert.Equal("146.745", restarted.Rows[0].Cells[0].Display);
        Assert.Equal(1, restarted.CurrentItemIndex);
    }

    [Fact]
    public void DecliningToResumeDiscardsTheSavedStateAndStartsOver() {
        var viewModel = CreateViewModel();
        viewModel.InputText = "100";
        viewModel.CommitValueCommand.Execute(null);

        var restarted = CreateViewModel();
        _dialog.ConfirmAnswers.Enqueue(false);
        restarted.Start();

        Assert.Equal(ResultState.NotDone, restarted.Rows[0].Cells[0].State);
        Assert.False(File.Exists(SavedPath));
    }

    [Fact]
    public void StartWithoutSavedStateDoesNotAsk() {
        var viewModel = CreateViewModel();

        viewModel.Start();

        Assert.Empty(_dialog.ConfirmMessages);
        Assert.Empty(_dialog.Errors);
    }

    [Fact]
    public void StartWithBrokenSavedStateReportsItAndStartsOver() {
        Directory.CreateDirectory(_directory);
        File.WriteAllText(SavedPath, "{ これはJSONではない");
        var viewModel = CreateViewModel();

        viewModel.Start();

        Assert.Single(_dialog.Errors);
        Assert.Equal(ResultState.NotDone, viewModel.Rows[0].Cells[0].State);
    }

    [Fact]
    public void FailingToSaveReportsAnErrorWithoutLosingTheInput() {
        // 保存先のフォルダの位置にファイルを置いて、保存を失敗させる
        Directory.CreateDirectory(Path.GetDirectoryName(_directory)!);
        File.WriteAllText(_directory, "ファイル");
        var viewModel = CreateViewModel();
        viewModel.InputText = "100";

        viewModel.CommitValueCommand.Execute(null);

        Assert.Single(_dialog.Errors);
        Assert.Equal(ResultState.Ok, viewModel.Rows[0].Cells[0].State);
    }
}
