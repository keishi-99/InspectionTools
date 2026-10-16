using InspectionTools.Common;
using InspectionTools.Inspection;

namespace InspectionTools.Products.Tests;

/// <summary>EL5000AC の手順表が、製品の定義（products/EL5000AC.json）と、旧アプリのコマンドと食い違っていないことを確認する</summary>
public class El5000AcProcedureTests {

    private static readonly ProductDefinition Definition =
        ProductDefinitionLoader.LoadFile(RepositoryFiles.Find("products", "EL5000AC.json"));

    private static readonly ProductProcedure Procedure = El5000AcProcedure.Create();

    private static readonly InstrumentRole[] MeasuringRoles = [InstrumentRole.Dmm01, InstrumentRole.Dmm02, InstrumentRole.Cnt];

    // 転記先の行から、その項目の動作 ID を引く
    private static string? ActionIdOfRow(int excelRow) => Definition.Items.Single(item => item.ExcelRow == excelRow).ActionId;

    // ===== 定義と手順表の整合 =====

    [Fact]
    public void ProcedureIsForTheSameProductAsTheDefinition() {
        Assert.Equal(Definition.Product, Procedure.Product);
    }

    [Fact]
    public void EveryActionIdInTheDefinitionExistsInTheProcedure() {
        var unknown = Definition.Items
            .Where(item => item.ActionId is not null && !Procedure.Actions.ContainsKey(item.ActionId))
            .Select(item => $"{item.Name}: {item.ActionId}");

        Assert.Empty(unknown);
    }

    [Fact]
    public void EveryActionInTheProcedureIsUsedByAnItem() {
        var usedIds = Definition.Items.Select(item => item.ActionId).ToHashSet();

        Assert.Empty(Procedure.Actions.Keys.Where(id => !usedIds.Contains(id)));
    }

    [Fact]
    public void EverySettingNamedByAnActionHasACommand() {
        var missing = Procedure.Actions
            .SelectMany(action => action.Value.Settings.Select(setting => (action.Key, setting)))
            .Where(pair => !Procedure.Profiles[pair.setting.Role].States.ContainsKey(pair.setting.Name))
            .Select(pair => $"{pair.Key}: {pair.setting.Role} {pair.setting.Name}");

        Assert.Empty(missing);
    }

    [Fact]
    public void NoActionSetsTheSameRoleTwice() {
        var duplicated = Procedure.Actions
            .Where(action => action.Value.Settings.GroupBy(setting => setting.Role).Any(group => group.Count() > 1))
            .Select(action => action.Key);

        Assert.Empty(duplicated);
    }

    [Fact]
    public void IdleStatesAndInitialStatesHaveCommands() {
        foreach (var (role, state) in Procedure.IdleStates) {
            Assert.True(Procedure.Profiles[role].States.ContainsKey(state), $"{role} の待機の状態 {state} のコマンドがありません。");
        }
        foreach (var role in new[] { InstrumentRole.Dmm01, InstrumentRole.Dmm02, InstrumentRole.Fg, InstrumentRole.Dcs }) {
            Assert.True(Procedure.Profiles[role].States.ContainsKey(Procedure.Profiles[role].InitialState), $"{role} の初期状態のコマンドがありません。");
        }
    }

    [Fact]
    public void OnlyMeasuringRolesAcquireValues() {
        var acquisitionRoles = Procedure.Actions.Values.Where(action => action.Acquisition is not null).Select(action => action.Acquisition!.Role);

        Assert.All(acquisitionRoles, role => Assert.Contains(role, MeasuringRoles));
    }

    // EL5000AC が使う役割（OSC は PAF5 などの製品が使う役割で、EL5000AC は使わない）
    private static readonly InstrumentRole[] UsedRoles = [InstrumentRole.Dmm01, InstrumentRole.Dmm02, InstrumentRole.Fg, InstrumentRole.Dcs, InstrumentRole.Cnt];

    [Fact]
    public void ProfileExistsForEveryRoleOfTheProcedure() {
        Assert.Equal(UsedRoles.Order(), Procedure.Profiles.Keys.Order());
    }

    [Fact]
    public void RequirementsExistForEveryRoleAndMatchTheOldAppSelectionLists() {
        Assert.Equal(UsedRoles.Order(), Procedure.Requirements.Keys.Order());
        // 通信方式の一覧は配列なので、レコードの等価比較（参照の比較）ではなく、中身を別々に比べる
        void AssertRequirement(InstrumentRole role, string category, int[] signalTypes) {
            Assert.Equal(category, Procedure.Requirements[role].Category);
            Assert.Equal(signalTypes, Procedure.Requirements[role].SignalTypes);
        }

        AssertRequirement(InstrumentRole.Dmm01, "デジタルマルチメータ", [1, 2]);
        AssertRequirement(InstrumentRole.Dmm02, "デジタルマルチメータ", [1, 2]);
        AssertRequirement(InstrumentRole.Fg, "ファンクションジェネレータ", [2]);
        AssertRequirement(InstrumentRole.Dcs, "電流電圧発生器", [2]);
        AssertRequirement(InstrumentRole.Cnt, "ユニバーサルカウンタ", [3]);
    }

    [Fact]
    public void ItemsThatAcquireValuesAreExactlyTheMeasurementItems() {
        var acquiring = Definition.Items.Where(item => Procedure.FindAction(item.ActionId).Acquisition is not null).Select(item => item.Name);
        var measurement = Definition.Items.Where(item => item.Kind == ItemKind.Measurement).Select(item => item.Name);

        Assert.Equal(measurement, acquiring);
    }

    // ===== 項目ごとの動作 =====

    [Fact]
    public void CurrentConsumptionUsesDmm01AndNotDmm02() {
        var action = Procedure.FindAction(ActionIdOfRow(8));

        Assert.Equal([new InstrumentSetting(InstrumentRole.Dmm01, "ACI")], action.Settings);
        Assert.Equal(InstrumentRole.Dmm01, action.Acquisition!.Role);
    }

    [Fact]
    public void VoltageRowsUseDmm02InVoltageMode() {
        for (var row = 9; row <= 15; row++) {
            Assert.Equal("dmm02-dcv", ActionIdOfRow(row));
        }
    }

    [Fact]
    public void AnalogInputRowsUseTheDcsStepsOfInput1ThenInput2() {
        string[] steps = ["4mA", "12mA", "20mA", "1V", "3V", "5V"];

        Assert.Equal(steps.Select(step => $"dcs-{step}"), Enumerable.Range(18, 6).Select(ActionIdOfRow));
        Assert.Equal(steps.Select(step => $"dcs-{step}"), Enumerable.Range(24, 6).Select(ActionIdOfRow));
    }

    [Fact]
    public void AnalogOutputRowsUseDmm02InCurrentModeThenVoltageMode() {
        Assert.All(Enumerable.Range(30, 3), row => Assert.Equal("dmm02-dci", ActionIdOfRow(row)));
        Assert.All(Enumerable.Range(33, 3), row => Assert.Equal("dmm02-dcv", ActionIdOfRow(row)));
    }

    [Fact]
    public void PulseOutputRowsUseTheCounter() {
        Assert.All(Enumerable.Range(36, 3), row => Assert.Equal("cnt-pulse-width", ActionIdOfRow(row)));
    }

    [Fact]
    public void OnlyPulseInput2UsesTheFunctionGenerator() {
        var rowsUsingFg = Definition.Items.Where(item => item.ActionId == "fg-on").Select(item => item.ExcelRow);

        Assert.Equal([47], rowsUsingFg);
    }

    [Fact]
    public void RowsThatDoNotUseInstrumentsHaveNoActionId() {
        // 行39〜46のパルス入力は別の装置を使う（アプリが管理するのは、行47の FG だけ）
        int[] rows = [5, 6, 7, 16, 17, .. Enumerable.Range(39, 8), 48, 49, 50];

        Assert.All(rows, row => Assert.Null(ActionIdOfRow(row)));
    }

    [Fact]
    public void TaskItemsStartWithTwoPt100CalibrationsWithoutInstrumentThenDcsStepsTwice() {
        var tasks = Definition.Items.Where(item => item.Kind == ItemKind.Task).Select(item => item.ActionId).ToList();

        Assert.Equal(
            [null, null, "dcs-4mA", "dcs-20mA", "dcs-5V", "dcs-1V", "dcs-4mA", "dcs-20mA", "dcs-5V", "dcs-1V"],
            tasks);
    }

    [Fact]
    public void DcsActionIdMatchesTheStepWrittenInTheItemName() {
        // 項目名の末尾（4mA、1V など）と動作 ID が食い違うと、名前と違う電圧を出してしまう
        var dcsItems = Definition.Items.Where(item => item.ActionId is not null && item.ActionId.StartsWith("dcs-"));

        Assert.All(dcsItems, item => Assert.Equal($"dcs-{item.Name.Split(' ')[^1]}", item.ActionId));
        Assert.Equal(8 + 12, dcsItems.Count());
    }

    // ===== 旧アプリのコマンドとの照合 =====

    [Fact]
    public void EveryCommandStringIsWrittenInTheOldApp() {
        var oldSource = File.ReadAllText(RepositoryFiles.Find("InspectionTools", "Product", "EL5000ACUserControl.xaml.cs"));
        var allCommands = Procedure.Profiles.Values
            .SelectMany(profile => new[] { profile.Init, profile.AfterRead }.Concat(profile.States.Values))
            .OfType<SwitchCommand>();

        var missing = allCommands
            .SelectMany(command => new[] { command.Adc, command.Visa, command.Gpib })
            .Where(text => text.Length > 0 && !oldSource.Contains($"\"{text}\""))
            .Distinct();

        Assert.Empty(missing);
    }
}
