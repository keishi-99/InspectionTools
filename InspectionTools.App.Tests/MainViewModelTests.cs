using ClosedXML.Excel;
using InspectionTools.App.Input;
using InspectionTools.App.ViewModels;
using InspectionTools.Inspection;

namespace InspectionTools.App.Tests;

/// <summary>検査画面の ViewModel の動き（入力、判定、移動、台数の変更、途中保存と復元）を、画面なしで確認する</summary>
public sealed class MainViewModelTests : IDisposable {

    private const string Product = "EL5000AC";
    private const int MaxUnitCount = 30;

    private static readonly TransferDefinition Transfer = new("ACdata", "ACsetting", 5, 50, "A3", "EL5000", "F", UnitsPerBlock: null, BlockRowOffset: 0);

    // 確認の答えを順番に返し、通知された内容を記録するダミー
    private sealed class FakeDialogService : IDialogService {
        public Queue<bool> ConfirmAnswers { get; } = new();
        public List<string> ConfirmMessages { get; } = [];
        public List<string> Errors { get; } = [];
        public List<string> Infos { get; } = [];

        /// <summary>転記先として選ばれるファイル（null なら、選ばずに閉じたことになる）</summary>
        public string? TemplatePathToSelect { get; set; }

        /// <summary>確認ダイアログが開かれた瞬間に呼ばれる(そのときの画面の状態を調べるため)</summary>
        public Action<string>? OnConfirm { get; set; }

        public bool Confirm(string message) {
            OnConfirm?.Invoke(message);
            ConfirmMessages.Add(message);
            return ConfirmAnswers.Dequeue();
        }

        public void ShowError(string message) => Errors.Add(message);

        public List<string> Warnings { get; } = [];

        public void ShowWarning(string message) => Warnings.Add(message);

        public void ShowInfo(string message) => Infos.Add(message);

        /// <summary>測定器の選択画面が開かれた回数</summary>
        public int InstrumentSelectionShownCount { get; private set; }

        public void ShowInstrumentSelection(InstrumentSelectionViewModel viewModel) => InstrumentSelectionShownCount++;

        public string? SelectTemplateFilePath() => TemplatePathToSelect;
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

    private string TemplatePath => Path.Combine(_directory, "template.xlsx");

    private string SavedPath => Path.Combine(_directory, Product + ".json");

    // 0: 測定、1: 入力、2: 確認、3: 作業 の4項目
    private static ProductDefinition CreateDefinition() => new(Product, MaxUnitCount, Transfer, [
        new InspectionItem("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m),
        new InspectionItem("温度入力", ItemKind.Input, excelRow: 16, lowerLimit: -50.2m, upperLimit: -49.8m),
        new InspectionItem("電源投入", ItemKind.Confirmation, excelRow: 5),
        new InspectionItem("調整作業", ItemKind.Task, description: "DCS を 4mA にして調整する"),
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

    [Fact]
    public void NonNumberMessageShowsTheEnteredText() {
        var viewModel = CreateViewModel();
        viewModel.InputText = "1.2.3";

        viewModel.CommitValueCommand.Execute(null);

        Assert.Contains("'1.2.3' は数値として読めません", Assert.Single(_dialog.Errors));
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
        viewModel.MoveToAsync(2, 0);
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
    public void OnlyTheCurrentItemRowIsMarkedAsCurrent() {
        var viewModel = CreateViewModel();
        Assert.Equal([true, false, false, false], viewModel.Rows.Select(row => row.IsCurrent));

        viewModel.MoveToAsync(2, 0);

        Assert.Equal([false, false, true, false], viewModel.Rows.Select(row => row.IsCurrent));
    }

    [Fact]
    public void CurrentRowMarkSurvivesChangingTheUnitCount() {
        var viewModel = CreateViewModel();
        viewModel.MoveToAsync(2, 0);

        viewModel.IncreaseUnitCountCommand.Execute(null);

        Assert.Equal([false, false, true, false], viewModel.Rows.Select(row => row.IsCurrent));
    }

    [Fact]
    public void KeyActionAvailabilityFollowsTheKindOfTheCurrentItem() {
        var viewModel = CreateViewModel();

        // 測定項目(測定器に接続していないので、値の取得は使えない)
        Assert.False(viewModel.IsKeyActionAvailable(KeyAction.AcquireValue));
        Assert.True(viewModel.IsKeyActionAvailable(KeyAction.Confirm));
        Assert.False(viewModel.IsKeyActionAvailable(KeyAction.Pass));
        Assert.False(viewModel.IsKeyActionAvailable(KeyAction.Fail));
        Assert.False(viewModel.IsKeyActionAvailable(KeyAction.Reapply));
        Assert.True(viewModel.IsKeyActionAvailable(KeyAction.Next));

        // 確認項目
        viewModel.MoveToAsync(2, 0);
        Assert.False(viewModel.IsKeyActionAvailable(KeyAction.Confirm));
        Assert.True(viewModel.IsKeyActionAvailable(KeyAction.Pass));
        Assert.True(viewModel.IsKeyActionAvailable(KeyAction.Fail));

        // 作業項目(確定のキーで、作業完了になる)
        viewModel.MoveToAsync(3, 0);
        Assert.True(viewModel.IsKeyActionAvailable(KeyAction.Confirm));
        Assert.False(viewModel.IsKeyActionAvailable(KeyAction.Pass));
    }

    [Fact]
    public void DescriptionIsShownOnlyForItemsThatHaveOne() {
        var viewModel = CreateViewModel();
        Assert.False(viewModel.HasDescription);
        Assert.Equal("", viewModel.CurrentDescription);

        viewModel.MoveToAsync(3, 0);

        Assert.True(viewModel.HasDescription);
        Assert.Equal("DCS を 4mA にして調整する", viewModel.CurrentDescription);
    }

    [Fact]
    public void CommittingAtTheLastCellStaysThere() {
        var viewModel = CreateViewModel();
        viewModel.MoveToAsync(3, 0);

        viewModel.CompleteTaskCommand.Execute(null);

        Assert.Equal(3, viewModel.CurrentItemIndex);
        Assert.Equal(0, viewModel.CurrentUnitIndex);
    }

    [Fact]
    public void MovingToCellWithValueShowsItInTheInputBox() {
        var viewModel = CreateViewModel();
        viewModel.InputText = "100";
        viewModel.CommitValueCommand.Execute(null);

        viewModel.MoveToAsync(0, 0);

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
        viewModel.MoveToAsync(0, 1);
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
        viewModel.MoveToAsync(0, 2);
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
        viewModel.MoveToAsync(0, 2);
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

    [Fact]
    public void SaveFailureIsReportedOnlyOnceWhileItKeepsFailing() {
        // 項目を進めるたびに保存するので、失敗が続いても、通知は最初の1回だけにする
        Directory.CreateDirectory(Path.GetDirectoryName(_directory)!);
        File.WriteAllText(_directory, "ファイル");
        var viewModel = CreateViewModel();

        viewModel.MoveToAsync(1, 0);
        viewModel.MoveToAsync(2, 0);
        viewModel.MoveToAsync(3, 0);

        Assert.Single(_dialog.Errors);
        Assert.Contains("繰り返しません", _dialog.Errors[0]);
    }

    [Fact]
    public void SaveFailureIsReportedAgainAfterASuccessfulSave() {
        Directory.CreateDirectory(Path.GetDirectoryName(_directory)!);
        File.WriteAllText(_directory, "ファイル");
        var viewModel = CreateViewModel();
        viewModel.MoveToAsync(1, 0);
        Assert.Single(_dialog.Errors);

        // 保存できるようにしてから、保存に成功させる。そのあと、また失敗させる
        File.Delete(_directory);
        viewModel.MoveToAsync(2, 0);
        Directory.Delete(_directory, recursive: true);
        File.WriteAllText(_directory, "ファイル");
        viewModel.MoveToAsync(3, 0);

        Assert.Equal(2, _dialog.Errors.Count);
    }

    [Fact]
    public void RestoreFailureExplainsThatTheDefinitionMayHaveChanged() {
        Directory.CreateDirectory(_directory);
        File.WriteAllText(SavedPath, "{ これはJSONではない");
        var viewModel = CreateViewModel();

        viewModel.Start();

        var message = Assert.Single(_dialog.Errors);
        Assert.Contains("定義(項目)が変わった", message);
        Assert.Contains("最初から始めます", message);
    }

    // ===== 文字入力（シリアルなど） =====

    // 0: 文字入力、1: 入力
    private MainViewModel CreateViewModelWithTextItem() => new(new ProductDefinition(Product, MaxUnitCount, Transfer, [
        new InspectionItem("LCDシリアル", ItemKind.Text, excelRow: 7),
        new InspectionItem("温度入力", ItemKind.Input, excelRow: 16, lowerLimit: -50.2m, upperLimit: -49.8m),
    ]), _store, _dialog);

    [Fact]
    public void TextItemUsesTheInputBoxWithoutRange() {
        var viewModel = CreateViewModelWithTextItem();

        Assert.True(viewModel.IsValueInput);
        Assert.True(viewModel.IsTextInput);
        Assert.Equal("", viewModel.CurrentRangeText);
        Assert.False(viewModel.HasAcquisition);
    }

    [Fact]
    public void CommittingTextShowsItInTheCellMarksOkAndMovesOn() {
        var viewModel = CreateViewModelWithTextItem();
        viewModel.InputText = "  0012345Z ";

        viewModel.CommitValueCommand.Execute(null);

        var cell = viewModel.Rows[0].Cells[0];
        Assert.Equal(ResultState.Ok, cell.State);
        Assert.Equal("0012345Z", cell.Display);
        Assert.Equal(1, viewModel.CurrentItemIndex);
        Assert.Equal("", viewModel.InputText);
        Assert.False(viewModel.IsTextInput);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void CommittingBlankTextShowsErrorAndChangesNothing(string text) {
        var viewModel = CreateViewModelWithTextItem();
        viewModel.InputText = text;

        viewModel.CommitValueCommand.Execute(null);

        Assert.Contains("入力されていません", Assert.Single(_dialog.Errors));
        Assert.Equal(ResultState.NotDone, viewModel.Rows[0].Cells[0].State);
        Assert.Equal(0, viewModel.CurrentItemIndex);
    }

    [Fact]
    public void ReturningToACommittedTextCellShowsTheTextInTheInputBox() {
        var viewModel = CreateViewModelWithTextItem();
        viewModel.InputText = "ABC-1";
        viewModel.CommitValueCommand.Execute(null);

        viewModel.MoveToAsync(0, 0);

        Assert.Equal("ABC-1", viewModel.InputText);
    }

    [Fact]
    public async Task ConfirmKeyCommitsTheText() {
        var viewModel = CreateViewModelWithTextItem();
        viewModel.InputText = "ABC-1";

        await viewModel.ExecuteKeyActionAsync(KeyAction.Confirm);

        Assert.Equal("ABC-1", viewModel.Rows[0].Cells[0].Display);
        Assert.True(viewModel.IsKeyActionAvailable(KeyAction.Confirm));
    }

    [Fact]
    public void TextIsRestoredFromTheSavedStateWhenTheAppIsStartedAgain() {
        var viewModel = CreateViewModelWithTextItem();
        viewModel.InputText = "ABC-1";
        viewModel.CommitValueCommand.Execute(null);

        var restored = _store.Load(Product, [
            new InspectionItem("LCDシリアル", ItemKind.Text, excelRow: 7),
            new InspectionItem("温度入力", ItemKind.Input, excelRow: 16, lowerLimit: -50.2m, upperLimit: -49.8m),
        ]);

        Assert.Equal(LoadStatus.Restored, restored.Status);
        Assert.Equal("ABC-1", restored.Session!.GetResult(0, 0).Text);
    }

    // ===== 初期値（defaultValue） =====

    // 0: 初期値のある入力項目、1: 初期値のない入力項目
    private static List<InspectionItem> CreateDefaultValueItems() => [
        new("ノッチ周波数", ItemKind.Input, excelRow: 23, lowerLimit: 6000m, upperLimit: 6500m, defaultValue: 6250m),
        new("温度入力", ItemKind.Input, excelRow: 16, lowerLimit: -50.2m, upperLimit: -49.8m),
    ];

    private MainViewModel CreateViewModelWithDefaultValue() =>
        new(new ProductDefinition(Product, MaxUnitCount, Transfer, CreateDefaultValueItems()), _store, _dialog);

    // 保存された途中の結果を読み直して、最初のマスが手入力として記録されているかを調べる
    private bool IsFirstCellManual() =>
        _store.Load(Product, CreateDefaultValueItems()).Session!.GetResult(0, 0).IsManual;

    [Fact]
    public void DefaultValueIsShownInTheInputBoxOfAnItemWithoutResult() {
        var viewModel = CreateViewModelWithDefaultValue();

        Assert.Equal("6250", viewModel.InputText);
    }

    [Fact]
    public void ItemWithoutDefaultValueStartsWithAnEmptyInputBox() {
        var viewModel = CreateViewModelWithDefaultValue();
        viewModel.CommitValueCommand.Execute(null);

        Assert.Equal(1, viewModel.CurrentItemIndex);
        Assert.Equal("", viewModel.InputText);
    }

    // 初期値は人が打ち込んだ値ではないので、そのまま確定しても手入力にはしない
    [Fact]
    public void CommittingTheDefaultValueAsIsIsNotRecordedAsManual() {
        var viewModel = CreateViewModelWithDefaultValue();

        viewModel.CommitValueCommand.Execute(null);

        Assert.Equal(ResultState.Ok, viewModel.Rows[0].Cells[0].State);
        Assert.False(IsFirstCellManual());
    }

    [Fact]
    public void EditingTheDefaultValueIsRecordedAsManual() {
        var viewModel = CreateViewModelWithDefaultValue();
        viewModel.InputText = "6300";

        viewModel.CommitValueCommand.Execute(null);

        Assert.True(IsFirstCellManual());
    }

    [Fact]
    public void ReturningToACommittedCellShowsTheResultInsteadOfTheDefaultValue() {
        var viewModel = CreateViewModelWithDefaultValue();
        viewModel.InputText = "6300";
        viewModel.CommitValueCommand.Execute(null);

        viewModel.MoveToAsync(0, 0);

        Assert.Equal("6300", viewModel.InputText);
    }

    // ===== 検査終了（転記） =====

    // 定義の4項目（行8、16、5と作業）に対応する、仮のテンプレートを作る
    private void CreateTemplate(string productName = "EL5000", double measurementUpper = 230) {
        Directory.CreateDirectory(_directory);
        using var workbook = new XLWorkbook();
        workbook.AddWorksheet("ACdata").Cell("A3").Value = productName;
        var setting = workbook.AddWorksheet("ACsetting");
        setting.Cell("F5").Value = "良";
        setting.Cell("F8").Value = 0;
        setting.Cell("G8").Value = measurementUpper;
        setting.Cell("F16").Value = -50.2;
        setting.Cell("G16").Value = -49.8;
        workbook.SaveAs(TemplatePath);
    }

    // 1台目の4項目をすべて実施した状態にする
    private MainViewModel CreateFinishedViewModel() {
        var viewModel = CreateViewModel();
        viewModel.InputText = "146.745";
        viewModel.CommitValueCommand.Execute(null);
        viewModel.InputText = "-50";
        viewModel.CommitValueCommand.Execute(null);
        viewModel.ConfirmOkCommand.Execute(null);
        viewModel.CompleteTaskCommand.Execute(null);
        return viewModel;
    }

    private double? ReadNumber(string cell) {
        using var workbook = new XLWorkbook(TemplatePath);
        var value = workbook.Worksheet("ACdata").Cell(cell).Value;
        return value.IsNumber ? value.GetNumber() : null;
    }

    [Fact]
    public async Task FinishingTransfersTheResultsAndStartsANewInspection() {
        CreateTemplate();
        _dialog.TemplatePathToSelect = TemplatePath;
        var viewModel = CreateFinishedViewModel();

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Equal(146.745, ReadNumber("F8"));
        Assert.Equal(-50, ReadNumber("F16"));
        Assert.Empty(_dialog.Errors);
        Assert.Contains("1台分", Assert.Single(_dialog.Infos));
        Assert.False(File.Exists(SavedPath));
        Assert.Equal(ResultState.NotDone, viewModel.Rows[0].Cells[0].State);
        Assert.Equal(1, viewModel.UnitCount);
    }

    // ===== 転記中のスピナー =====

    // 転記は時間がかかるので、背景で行い、その間はスピナーを出して画面の操作を止める。busy の変化と、そのときの文言を記録する
    private static List<(bool IsBusy, string Text)> RecordBusyChanges(MainViewModel viewModel) {
        var changes = new List<(bool, string)>();
        viewModel.PropertyChanged += (_, e) => {
            if (e.PropertyName == nameof(MainViewModel.IsInstrumentBusy)) {
                changes.Add((viewModel.IsInstrumentBusy, viewModel.BusyText));
            }
        };
        return changes;
    }

    [Fact]
    public void BusyTextIsAboutInstrumentCommunicationAtFirst() {
        Assert.Equal("測定器と通信中…", CreateViewModel().BusyText);
    }

    [Fact]
    public async Task SpinnerIsShownWhileTransferringAndClearedAfterwards() {
        CreateTemplate();
        _dialog.TemplatePathToSelect = TemplatePath;
        var viewModel = CreateFinishedViewModel();
        var changes = RecordBusyChanges(viewModel);

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        // 合格範囲の読み取りと、書き込みの2回、「Excel に転記しています…」を出して、最後は消えている
        Assert.Contains((true, "Excel に転記しています…"), changes);
        Assert.Equal(2, changes.Count(change => change.IsBusy));
        Assert.False(viewModel.IsInstrumentBusy);
    }

    // 確認ダイアログが出ている間は、スピナーを止める(「転記しています」と出たまま、確認を求めないため)
    [Fact]
    public async Task SpinnerIsClearedWhileAConfirmationIsShown() {
        CreateTemplate();
        _dialog.TemplatePathToSelect = TemplatePath;
        var viewModel = CreateViewModel();
        bool? busyWhenAsked = null;
        _dialog.OnConfirm = _ => busyWhenAsked = viewModel.IsInstrumentBusy;
        _dialog.ConfirmAnswers.Enqueue(false);

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.False(busyWhenAsked);
        Assert.False(viewModel.IsInstrumentBusy);
        Assert.Empty(_dialog.Infos);
    }

    [Fact]
    public async Task SpinnerIsClearedWhenTheTransferFails() {
        CreateTemplate(productName: "EL9240");
        _dialog.TemplatePathToSelect = TemplatePath;
        var viewModel = CreateFinishedViewModel();

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Single(_dialog.Errors);
        Assert.False(viewModel.IsInstrumentBusy);
    }

    [Fact]
    public async Task SpinnerIsNotShownWhenTheFileDialogIsClosedWithoutChoosing() {
        _dialog.TemplatePathToSelect = null;
        var viewModel = CreateFinishedViewModel();
        var changes = RecordBusyChanges(viewModel);

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Empty(changes);
    }

    // 合格範囲のシートも製品の目印もないテンプレート（PAF5）では、合格範囲の突き合わせと製品の確認をせずに転記する。確認項目はリンク先の列に TRUE を書く
    [Fact]
    public async Task FinishingWithoutSettingSheetAndProductCheckTransfersAndWritesConfirmationToItsColumn() {
        Directory.CreateDirectory(_directory);
        using (var workbook = new XLWorkbook()) {
            workbook.AddWorksheet("ACdata");
            workbook.SaveAs(TemplatePath);
        }
        _dialog.TemplatePathToSelect = TemplatePath;
        var transfer = new TransferDefinition("ACdata", null, 5, 50, null, null, "F", UnitsPerBlock: null, BlockRowOffset: 0, ConfirmationColumn: "K");
        var viewModel = new MainViewModel(new ProductDefinition(Product, MaxUnitCount, transfer, CreateDefinition().Items), _store, _dialog);
        viewModel.InputText = "146.745";
        viewModel.CommitValueCommand.Execute(null);
        viewModel.InputText = "-50";
        viewModel.CommitValueCommand.Execute(null);
        viewModel.ConfirmOkCommand.Execute(null);
        viewModel.CompleteTaskCommand.Execute(null);

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Empty(_dialog.Errors);
        Assert.Empty(_dialog.ConfirmMessages);
        Assert.Equal(146.745, ReadNumber("F8"));
        using var result = new XLWorkbook(TemplatePath);
        Assert.True(result.Worksheet("ACdata").Cell("K5").Value.GetBoolean());
    }

    // 台を1台ずつ別のシートに書く製品（PAF5 では5台ずつ）では、2台目が2枚目のシートに転記され、各シートの台数のセル(C1)に台数が入る
    [Fact]
    public async Task FinishingWithTwoUnitsOnTwoSheetsWritesEachUnitToItsSheet() {
        Directory.CreateDirectory(_directory);
        using (var workbook = new XLWorkbook()) {
            workbook.AddWorksheet("ACdata");
            workbook.AddWorksheet("ACdata2");
            workbook.SaveAs(TemplatePath);
        }
        _dialog.TemplatePathToSelect = TemplatePath;
        var transfer = new TransferDefinition("ACdata", null, 5, 50, null, null, "F", UnitsPerBlock: null, BlockRowOffset: 0,
            DataSheets: ["ACdata", "ACdata2"], UnitsPerSheet: 1, UnitCountCell: "C1");
        var viewModel = new MainViewModel(new ProductDefinition(Product, 2, transfer, CreateDefinition().Items), _store, _dialog);
        viewModel.IncreaseUnitCountCommand.Execute(null);
        foreach (var value in new[] { "146.745", "100" }) {
            viewModel.InputText = value;
            viewModel.CommitValueCommand.Execute(null);
            viewModel.InputText = "-50";
            viewModel.CommitValueCommand.Execute(null);
            viewModel.ConfirmOkCommand.Execute(null);
            viewModel.CompleteTaskCommand.Execute(null);
        }

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Empty(_dialog.Errors);
        using var result = new XLWorkbook(TemplatePath);
        Assert.Equal(146.745, result.Worksheet("ACdata").Cell("F8").Value.GetNumber());
        Assert.Equal(100, result.Worksheet("ACdata2").Cell("F8").Value.GetNumber());
        Assert.Equal(1, result.Worksheet("ACdata").Cell("C1").Value.GetNumber());
        Assert.Equal(1, result.Worksheet("ACdata2").Cell("C1").Value.GetNumber());
    }

    [Fact]
    public async Task ClosingTheFileDialogWithoutChoosingDoesNothing() {
        _dialog.TemplatePathToSelect = null;
        var viewModel = CreateFinishedViewModel();

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Empty(_dialog.Errors);
        Assert.Empty(_dialog.Infos);
        Assert.True(File.Exists(SavedPath));
        Assert.Equal(ResultState.Ok, viewModel.Rows[0].Cells[0].State);
    }

    [Fact]
    public async Task TemplateOfAnotherProductIsRejectedAndTheResultsAreKept() {
        CreateTemplate(productName: "EL9240");
        _dialog.TemplatePathToSelect = TemplatePath;
        var viewModel = CreateFinishedViewModel();

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Single(_dialog.Errors);
        Assert.Empty(_dialog.Infos);
        Assert.True(File.Exists(SavedPath));
        Assert.Equal(ResultState.Ok, viewModel.Rows[0].Cells[0].State);
        Assert.Null(ReadNumber("F8"));
    }

    [Fact]
    public async Task TemplateOpenedByAnotherProgramIsReportedAndTheResultsAreKept() {
        CreateTemplate();
        _dialog.TemplatePathToSelect = TemplatePath;
        var viewModel = CreateFinishedViewModel();
        using var otherProgram = new FileStream(TemplatePath, FileMode.Open, FileAccess.Read, FileShare.Read);

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Single(_dialog.Errors);
        Assert.True(File.Exists(SavedPath));
        Assert.Equal(ResultState.Ok, viewModel.Rows[0].Cells[0].State);
    }

    [Theory]
    [InlineData(true, 146.745)]
    [InlineData(false, null)]
    public async Task DifferentRangeInTheTemplateAsksAndFollowsTheAnswer(bool answer, double? expectedValue) {
        // 定義は上限230、テンプレートは上限200
        CreateTemplate(measurementUpper: 200);
        _dialog.TemplatePathToSelect = TemplatePath;
        _dialog.ConfirmAnswers.Enqueue(answer);
        var viewModel = CreateFinishedViewModel();

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Contains("消費電流", Assert.Single(_dialog.ConfirmMessages));
        Assert.Equal(expectedValue, ReadNumber("F8"));
    }

    [Theory]
    [InlineData(true, 146.745)]
    [InlineData(false, null)]
    public async Task UndoneItemsAskBeforeTransferAndFollowTheAnswer(bool answer, double? expectedValue) {
        CreateTemplate();
        _dialog.TemplatePathToSelect = TemplatePath;
        _dialog.ConfirmAnswers.Enqueue(answer);
        var viewModel = CreateViewModel();
        // 測定項目だけ実施し、入力項目と確認項目は未実施のまま（作業項目は数えない）
        viewModel.InputText = "146.745";
        viewModel.CommitValueCommand.Execute(null);

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Contains("2 か所", Assert.Single(_dialog.ConfirmMessages));
        Assert.Equal(expectedValue, ReadNumber("F8"));
        Assert.Equal(answer, !File.Exists(SavedPath));
    }

    [Fact]
    public async Task FailingToDeleteTheSavedStateAfterTransferStillStartsANewInspection() {
        CreateTemplate();
        _dialog.TemplatePathToSelect = TemplatePath;
        var viewModel = CreateFinishedViewModel();
        // 途中保存のファイルを別のプログラムが開いている状態にして、削除だけを失敗させる
        using var otherProgram = new FileStream(SavedPath, FileMode.Open, FileAccess.Read, FileShare.None);

        await viewModel.FinishInspectionCommand.ExecuteAsync(null);

        Assert.Equal(146.745, ReadNumber("F8"));
        Assert.Contains("転記は完了", Assert.Single(_dialog.Errors));
        Assert.Empty(_dialog.Infos);
        Assert.Equal(ResultState.NotDone, viewModel.Rows[0].Cells[0].State);
    }
}
