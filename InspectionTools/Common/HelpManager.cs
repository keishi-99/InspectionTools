using System.IO;
using System.Text.Json;

namespace InspectionTools.Common {
    public static class HelpManager {

        private static Dictionary<string, List<HelpEntry>>? _helpTexts;

        // JSONファイルを読み込んでヘルプテキストを_helpTexts辞書に格納する
        public static void LoadHelpFile(string path) {
            if (!File.Exists(path)) {
                _helpTexts = [];
                return;
            }

            string json = File.ReadAllText(path);
            var rawData = JsonSerializer.Deserialize<Dictionary<string, List<Dictionary<string, string>>>>(json);

            _helpTexts = [];

            if (rawData == null) return;

            foreach (var kvp in rawData) {
                var entries = new List<HelpEntry>();

                foreach (var item in kvp.Value) {
                    foreach (var pair in item) {
                        entries.Add(new HelpEntry(pair.Key, pair.Value));
                    }
                }

                _helpTexts[kvp.Key] = entries;
            }
        }

        // 指定ページ名のヘルプエントリ一覧を取得する
        public static IReadOnlyList<HelpEntry> GetHelpData(string pageName) {
            if (_helpTexts == null)
                throw new InvalidOperationException("HelpManager: JSONファイルが読み込まれていません。");

            if (!_helpTexts.TryGetValue(pageName, out var entries))
                return [];

            return entries;
        }
    }
}