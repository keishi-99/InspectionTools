using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace InspectionTools.Common {
    public static class HelpManager {

        // JSONエントリのデシリアライズ用DTO
        private record HelpEntryDto(
            [property: JsonPropertyName("keys")] string[] Keys,
            [property: JsonPropertyName("description")] string Description
        );

        private static Dictionary<string, List<HelpEntry>>? _helpTexts;

        // JSONファイルを読み込んでヘルプテキストを_helpTexts辞書に格納する
        public static void LoadHelpFile(string path) {
            if (!File.Exists(path)) {
                _helpTexts = [];
                return;
            }

            string json = File.ReadAllText(path);

            Dictionary<string, List<HelpEntryDto>>? rawData;
            try {
                rawData = JsonSerializer.Deserialize<Dictionary<string, List<HelpEntryDto>>>(json);
            } catch (JsonException) {
                // JSONが不正な形式の場合は空データで初期化してクラッシュを防ぐ
                _helpTexts = [];
                return;
            }

            _helpTexts = [];

            if (rawData == null) return;

            foreach (var kvp in rawData) {
                var entries = new List<HelpEntry>();

                foreach (var item in kvp.Value) {
                    entries.Add(new HelpEntry(item.Keys, item.Description));
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