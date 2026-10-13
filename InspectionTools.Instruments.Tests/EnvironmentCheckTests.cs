using InspectionTools.Common;

namespace InspectionTools.Instruments.Tests;

/// <summary>MSBuild でのビルドと dotnet test --no-build が動くことを確認するための仮のテスト</summary>
public class EnvironmentCheckTests {

    [Fact]
    public void CanReferenceInstrumentsLibrary() {
        var dmm = new DmmInstClass();

        Assert.Equal(0, dmm.SettingNumber);
    }
}
