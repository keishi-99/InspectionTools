using System.IO;
using System.Text.Json;

namespace InspectionTools.Common {
    public static class HelpManager {

        //public class HelpEntry {
        //    public string Key { get; set; } = string.Empty;
        //    public string Description { get; set; } = string.Empty;
        //}

        //private static Dictionary<string, List<HelpEntry>>? s_helpData;


        //public static void LoadHelpFile(string path) {
        //    if (!File.Exists(path)) {
        //        s_helpData = [];
        //        return;
        //    }

        //    string json = File.ReadAllText(path);
        //    s_helpData = JsonSerializer.Deserialize<Dictionary<string, List<HelpEntry>>>(json);
        //}

        //public static List<HelpEntry> GetHelpList(string key) {
        //    if (s_helpData == null) {
        //        throw new InvalidOperationException("Helpファイルが読み込まれていません。");
        //    }

        //    return s_helpData.TryGetValue(key, out var list)
        //        ? list
        //        : [new() { Key = "なし", Description = "ヘルプ情報が登録されていません。" }];
        //}


        //    private static Dictionary<string, string[]>? s_helpTexts;

        //    // JSONファイルの読み込み
        //    public static void LoadHelpFile(string path) {
        //        if (!File.Exists(path)) {
        //            s_helpTexts = [];
        //            return;
        //        }

        //        string json = File.ReadAllText(path);
        //        s_helpTexts = JsonSerializer.Deserialize<Dictionary<string, string[]>>(json);
        //    }

        //    // 指定キーのヘルプテキストを取得
        //    public static string GetHelpText(string key) {
        //        if (s_helpTexts == null) {
        //            throw new InvalidOperationException("HelpManager: JSONファイルが読み込まれていません。");
        //        }

        //        if (!s_helpTexts.TryGetValue(key, out var lines) || lines.Length == 0) {
        //            return "ヘルプ情報が登録されていません。";
        //        }

        //        return string.Join(Environment.NewLine, lines);
        //    }
        //}

        private static Dictionary<string, (string[] Keys, string[] Descriptions)>? s_helpTexts;

        // JSONファイルの読み込み
        public static void LoadHelpFile(string path) {
            if (!File.Exists(path)) {
                s_helpTexts = [];
                return;
            }

            string json = File.ReadAllText(path);

            // 一旦、Dictionary<string, List<Dictionary<string, string>>> で読み込む
            var rawData = JsonSerializer.Deserialize<Dictionary<string, List<Dictionary<string, string>>>>(json);

            s_helpTexts = [];

            if (rawData == null) {
                return;
            }

            foreach (var kvp in rawData) {
                var keys = new List<string>();
                var descriptions = new List<string>();

                foreach (var item in kvp.Value) {
                    foreach (var pair in item) {
                        keys.Add(pair.Key);
                        descriptions.Add(pair.Value);
                    }
                }

                s_helpTexts[kvp.Key] = (keys.ToArray(), descriptions.ToArray());
            }
        }

        // 指定キーのヘルプテキストを取得（KeyとDescriptionを結合して表示用）
        public static string GetHelpText(string key) {
            if (s_helpTexts == null) {
                throw new InvalidOperationException("HelpManager: JSONファイルが読み込まれていません。");
            }

            if (!s_helpTexts.TryGetValue(key, out var data) || data.Keys.Length == 0) {
                return "ヘルプ情報が登録されていません。";
            }

            // Key と Description を結合して見やすく表示
            var lines = data.Keys.Zip(data.Descriptions, (k, d) => $"{k}\t= {d}");
            return string.Join(Environment.NewLine, lines);
        }

        public static (string[] Keys, string[] Descriptions) GetHelpData(string pageName) {
            if (s_helpTexts == null) {
                throw new InvalidOperationException("HelpManager: JSONファイルが読み込まれていません。");
            }

            if (!s_helpTexts.TryGetValue(pageName, out var data)) {
                return (Array.Empty<string>(), Array.Empty<string>());
            }

            return data;
        }
    }
}
