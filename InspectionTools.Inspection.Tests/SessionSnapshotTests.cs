namespace InspectionTools.Inspection.Tests;

/// <summary>検査中の状態を途中保存の形に書き出し、そこから復元できることを確認する</summary>
public class SessionSnapshotTests {

    private const string Product = "EL5000AC";

    // 値・確認・作業の各種類に結果が入り、現在位置が先頭でない状態を作る
    private static InspectionSession CreateSessionWithResults() {
        var session = new InspectionSession(TestItems.Create(), unitCount: 2);
        session.SetValue(0, 0, 146.745m, isManual: false);
        session.SetValue(1, 0, -50.003m, isManual: true);
        session.SetConfirmation(2, 0, isOk: false);
        session.CompleteTask(3, 1);
        session.MoveTo(itemIndex: 1, unitIndex: 1);
        return session;
    }

    private static InspectionSession Restore(SessionSnapshot snapshot) =>
        InspectionSession.FromSnapshot(snapshot, Product, TestItems.Create());

    // 保存内容を壊した形に変えたとき、復元が SessionRestoreException で失敗することを確認する
    private static void AssertRejected(Func<SessionSnapshot, SessionSnapshot> breakSnapshot) {
        var snapshot = breakSnapshot(CreateSessionWithResults().ToSnapshot(Product));

        Assert.Throws<SessionRestoreException>(() => Restore(snapshot));
    }

    [Fact]
    public void ToSnapshotContainsOnlyDoneCellsAndPosition() {
        var snapshot = CreateSessionWithResults().ToSnapshot(Product);

        Assert.Equal(Product, snapshot.Product);
        Assert.Equal(4, snapshot.ItemCount);
        Assert.Equal(2, snapshot.UnitCount);
        Assert.Equal(1, snapshot.CurrentItemIndex);
        Assert.Equal(1, snapshot.CurrentUnitIndex);
        Assert.Equal(4, snapshot.Cells.Count);
    }

    // ===== 文字入力の項目 =====

    private static List<InspectionItem> CreateTextItems() => [
        new("LCDシリアル", ItemKind.Text, excelRow: 7),
        new("消費電流", ItemKind.Measurement, excelRow: 8, lowerLimit: 0m, upperLimit: 230m),
    ];

    [Fact]
    public void TextIsKeptThroughSnapshotAndRestore() {
        var original = new InspectionSession(CreateTextItems(), unitCount: 2);
        original.SetText(0, 1, "0012345Z");

        var restored = InspectionSession.FromSnapshot(original.ToSnapshot(Product), Product, CreateTextItems());

        Assert.Equal("0012345Z", restored.GetResult(0, 1).Text);
        Assert.Equal(ResultState.Ok, restored.GetResult(0, 1).State);
        Assert.Equal(ResultState.NotDone, restored.GetResult(0, 0).State);
    }

    // 文字入力を足す前に保存された途中保存（"text" のない形）も、そのまま読める
    [Fact]
    public void SnapshotSavedBeforeTextItemsWereAddedCanStillBeLoaded() {
        const string OldJson = """
            { "product": "EL5000AC", "itemCount": 4, "unitCount": 1, "currentItemIndex": 1, "currentUnitIndex": 0,
              "cells": [ { "unit": 0, "item": 0, "value": 146.745, "state": "Ok", "isManual": false } ] }
            """;
        var snapshot = System.Text.Json.JsonSerializer.Deserialize<SessionSnapshot>(
            OldJson,
            new System.Text.Json.JsonSerializerOptions(System.Text.Json.JsonSerializerDefaults.Web) { Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() } })!;

        var restored = Restore(snapshot);

        Assert.Equal(146.745m, restored.GetResult(0, 0).Value);
        Assert.Null(snapshot.Cells[0].Text);
    }

    [Fact]
    public void TextCellWithoutTextIsRejected() {
        var snapshot = new SessionSnapshot(Product, 2, 1, 0, 0, [new CellSnapshot(0, 0, null, ResultState.Ok, false)]);

        Assert.Throws<SessionRestoreException>(() => InspectionSession.FromSnapshot(snapshot, Product, CreateTextItems()));
    }

    [Fact]
    public void SnapshotRoundTripRestoresEveryCellAndPosition() {
        var original = CreateSessionWithResults();

        var restored = Restore(original.ToSnapshot(Product));

        Assert.Equal(original.UnitCount, restored.UnitCount);
        Assert.Equal(original.CurrentItemIndex, restored.CurrentItemIndex);
        Assert.Equal(original.CurrentUnitIndex, restored.CurrentUnitIndex);
        for (var unit = 0; unit < original.UnitCount; unit++) {
            for (var item = 0; item < original.Items.Count; item++) {
                var expected = original.GetResult(item, unit);
                var actual = restored.GetResult(item, unit);
                Assert.Equal(expected.Value, actual.Value);
                Assert.Equal(expected.State, actual.State);
                Assert.Equal(expected.IsManual, actual.IsManual);
            }
        }
    }

    // 項目の定義（合格範囲）が保存後に変わっていても、いまの範囲で判定し直す
    [Fact]
    public void RestoreRejudgesValueWithCurrentRange() {
        var snapshot = CreateSessionWithResults().ToSnapshot(Product) with {
            Cells = [new CellSnapshot(Unit: 0, Item: 0, Value: 300m, State: ResultState.Ok, IsManual: false)],
        };

        var restored = Restore(snapshot);

        Assert.Equal(ResultState.Ng, restored.GetResult(0, 0).State);
    }

    [Fact]
    public void RestoreRejectsValueCellWithoutValue() {
        AssertRejected(snapshot => snapshot with {
            Cells = [new CellSnapshot(Unit: 0, Item: 0, Value: null, State: ResultState.Ok, IsManual: false)],
        });
    }

    [Fact]
    public void RestoreRejectsStateThatDoesNotMatchItemKind() {
        // 確認項目に「完了」、作業項目に「合格」は、種類と合わない
        AssertRejected(snapshot => snapshot with {
            Cells = [new CellSnapshot(Unit: 0, Item: 2, Value: null, State: ResultState.Done, IsManual: false)],
        });
        AssertRejected(snapshot => snapshot with {
            Cells = [new CellSnapshot(Unit: 0, Item: 3, Value: null, State: ResultState.Ok, IsManual: false)],
        });
    }

    [Fact]
    public void RestoreRejectsDifferentProduct() {
        AssertRejected(snapshot => snapshot with { Product = "EL5000DC" });
    }

    [Fact]
    public void RestoreRejectsDifferentItemCount() {
        AssertRejected(snapshot => snapshot with { ItemCount = 3 });
    }

    [Fact]
    public void RestoreRejectsUnitCountLessThanOne() {
        AssertRejected(snapshot => snapshot with { UnitCount = 0 });
    }

    // セルがあると「台の位置が範囲外」の確認で偶然弾かれてしまうため、セルがない保存内容でも確認する
    [Fact]
    public void RestoreRejectsUnitCountLessThanOneEvenWithoutCells() {
        AssertRejected(snapshot => snapshot with { UnitCount = 0, Cells = [] });
    }

    [Fact]
    public void RestoreRejectsCurrentPositionOutOfRange() {
        AssertRejected(snapshot => snapshot with { CurrentItemIndex = 4 });
        AssertRejected(snapshot => snapshot with { CurrentItemIndex = -1 });
        AssertRejected(snapshot => snapshot with { CurrentUnitIndex = 2 });
    }

    [Fact]
    public void RestoreRejectsCellOutOfRange() {
        AssertRejected(snapshot => snapshot with {
            Cells = [new CellSnapshot(Unit: 0, Item: 4, Value: 1m, State: ResultState.Ok, IsManual: false)],
        });
        AssertRejected(snapshot => snapshot with {
            Cells = [new CellSnapshot(Unit: 2, Item: 0, Value: 1m, State: ResultState.Ok, IsManual: false)],
        });
    }
}
