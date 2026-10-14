namespace InspectionTools.App.Tests;

/// <summary>多重起動の防止を確認する</summary>
public sealed class SingleInstanceGuardTests {

    private static string UniqueName() => $"InspectionToolsTests-{Guid.NewGuid():N}";

    [Fact]
    public void FirstInstanceIsAllowedAndSecondIsNot() {
        var name = UniqueName();

        using var first = new SingleInstanceGuard(name);
        using var second = new SingleInstanceGuard(name);

        Assert.True(first.IsFirstInstance);
        Assert.False(second.IsFirstInstance);
    }

    [Fact]
    public void DifferentNamesDoNotInterfere() {
        using var first = new SingleInstanceGuard(UniqueName());
        using var second = new SingleInstanceGuard(UniqueName());

        Assert.True(first.IsFirstInstance);
        Assert.True(second.IsFirstInstance);
    }

    [Fact]
    public void CanStartAgainAfterTheFirstIsDisposed() {
        var name = UniqueName();
        new SingleInstanceGuard(name).Dispose();

        using var again = new SingleInstanceGuard(name);

        Assert.True(again.IsFirstInstance);
    }
}
