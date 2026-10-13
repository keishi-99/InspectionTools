using System.Text.Json;
using System.Text.Json.Serialization;

namespace InspectionTools.Inspection;

/// <summary>途中保存の読み込み結果の種類</summary>
public enum LoadStatus {
    /// <summary>保存ファイルがない</summary>
    NoSavedData,
    /// <summary>復元できた</summary>
    Restored,
    /// <summary>保存ファイルが壊れている、または定義と合わず、復元できなかった</summary>
    Failed,
}

/// <summary>途中保存の読み込み結果</summary>
/// <param name="Status">結果の種類</param>
/// <param name="Session">復元できた状態（Restored のときだけ入る）</param>
/// <param name="Message">復元できなかった理由（Failed のときだけ入る）</param>
/// <param name="BrokenFilePath">復元できなかったファイルを退避した先（Failed のときだけ入る）</param>
public sealed record LoadResult(LoadStatus Status, InspectionSession? Session, string? Message, string? BrokenFilePath);

/// <summary>検査中の状態を、製品ごとに1つの JSON ファイルへ保存・復元する</summary>
public sealed class SessionStore(string directory, TimeProvider? timeProvider = null) {

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web) {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
    };

    private readonly string _directory = directory;
    private readonly TimeProvider _timeProvider = timeProvider ?? TimeProvider.System;

    /// <summary>状態を保存する（同じ製品のファイルは上書きする）</summary>
    public void Save(string product, InspectionSession session) {
        var path = GetPath(product);
        Directory.CreateDirectory(_directory);

        // 書き込み中にアプリが落ちても、古い完全なファイルが残るように、別名に最後まで書いてから置き換える
        var temporaryPath = path + ".tmp";
        File.WriteAllText(temporaryPath, JsonSerializer.Serialize(session.ToSnapshot(product), JsonOptions));
        File.Move(temporaryPath, path, overwrite: true);
    }

    /// <summary>保存された状態を復元する。復元できないときは、壊れたファイルを別名で残して Failed を返す</summary>
    public LoadResult Load(string product, IReadOnlyList<InspectionItem> items) {
        var path = GetPath(product);
        if (!File.Exists(path)) {
            return new LoadResult(LoadStatus.NoSavedData, null, null, null);
        }

        try {
            var snapshot = JsonSerializer.Deserialize<SessionSnapshot>(File.ReadAllText(path), JsonOptions)
                ?? throw new SessionRestoreException("保存ファイルの内容が空です。");
            var session = InspectionSession.FromSnapshot(snapshot, product, items);
            return new LoadResult(LoadStatus.Restored, session, null, null);
        } catch (Exception ex) when (ex is JsonException or SessionRestoreException) {
            // 次回の保存で上書きされて失われないように、元のファイルは別名で残す
            var brokenPath = Path.Combine(_directory, $"{product}.broken-{_timeProvider.GetLocalNow():yyyyMMdd-HHmmss}.json");
            File.Move(path, brokenPath);
            return new LoadResult(LoadStatus.Failed, null, ex.Message, brokenPath);
        }
    }

    /// <summary>保存された状態を削除する（転記が終わったあと、次の起動で「続きから再開しますか」と聞かれないようにするため）。ファイルがなければ何もしない。退避した壊れたファイルは消さない</summary>
    public void Delete(string product) {
        var path = GetPath(product);
        if (File.Exists(path)) {
            File.Delete(path);
        }
    }

    // 製品名から保存ファイルのパスを作る（ファイル名に使えない文字を含む製品名は受け付けない）
    private string GetPath(string product) {
        if (string.IsNullOrWhiteSpace(product) || product.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
            throw new ArgumentException($"保存ファイルの名前に使えない製品名です: {product}", nameof(product));
        }
        return Path.Combine(_directory, product + ".json");
    }
}
