namespace InspectionTools.App.Tests;

/// <summary>テストの実行場所から上の階層をたどって、リポジトリにあるファイルを探す</summary>
internal static class RepositoryFiles {

    public static string Find(params string[] relativePath) {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent) {
            if (File.Exists(Path.Combine(directory.FullName, "InspectionTools.sln"))) {
                return Path.Combine([directory.FullName, .. relativePath]);
            }
        }
        throw new FileNotFoundException("リポジトリのルート（InspectionTools.sln）が見つかりません。");
    }
}
