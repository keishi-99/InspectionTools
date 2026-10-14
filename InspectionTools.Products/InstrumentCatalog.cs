using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace InspectionTools.Products;

/// <summary>測定器の一覧（VisaAddress.xml。旧アプリと同じ形式）。画面に依存せず、読み込みと候補の絞り込みだけを行う</summary>
public sealed class InstrumentCatalog {

    private InstrumentCatalog(IReadOnlyList<InstrumentEntry> entries, IReadOnlyList<string> warnings) {
        Entries = entries;
        Warnings = warnings;
    }

    /// <summary>使える測定器の一覧（ファイルの順）</summary>
    public IReadOnlyList<InstrumentEntry> Entries { get; }

    /// <summary>一覧に入れなかった項目の理由（書き方の誤りがあっても、ほかの測定器は使えるようにするため、読み飛ばして知らせる）</summary>
    public IReadOnlyList<string> Warnings { get; }

    /// <summary>ファイルから読み込む。ファイルがない、読み込めない、XML として正しくないときは InstrumentCatalogException</summary>
    public static InstrumentCatalog LoadFile(string path) {
        string xml;
        try {
            xml = File.ReadAllText(path);
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
            throw new InstrumentCatalogException($"測定器の一覧（{Path.GetFileName(path)}）を読み込めません: {ex.Message}", ex);
        }
        return Parse(xml);
    }

    /// <summary>XML の文字列から読み込む</summary>
    public static InstrumentCatalog Parse(string xml) {
        XDocument document;
        try {
            document = XDocument.Parse(xml);
        } catch (XmlException ex) {
            throw new InstrumentCatalogException($"測定器の一覧が XML として正しくありません: {ex.Message}", ex);
        }

        var entries = new List<InstrumentEntry>();
        var warnings = new List<string>();
        // 要素の名前空間（http://tempuri.org/dsInstList.xsd）に左右されないよう、名前だけで探す
        var elements = document.Descendants().Where(element => element.Name.LocalName == "tblInstList").ToList();
        for (var index = 0; index < elements.Count; index++) {
            var entry = CreateEntry(elements[index], index + 1, warnings);
            if (entry is null) {
                continue;
            }

            if (entries.Any(other => other.Category == entry.Category && other.Name == entry.Name)) {
                warnings.Add($"{index + 1}番目の測定器（{entry.Name}）: 同じカテゴリに同じ名前があるため、読み飛ばしました。");
                continue;
            }
            entries.Add(entry);
        }
        return new InstrumentCatalog(entries, warnings);
    }

    /// <summary>役割の条件（カテゴリと通信方式）に合う測定器を、ファイルの順に返す</summary>
    public IReadOnlyList<InstrumentEntry> FindCandidates(RoleRequirement requirement) {
        return Entries
            .Where(entry => entry.Category == requirement.Category && requirement.SignalTypes.Contains(entry.SignalType))
            .ToList();
    }

    // 1台分の要素を InstrumentEntry にする。使えない書き方のときは、理由を警告に入れて null を返す
    private static InstrumentEntry? CreateEntry(XElement element, int position, List<string> warnings) {
        var category = ReadText(element, "Category");
        var name = ReadText(element, "Name");
        var visaAddress = ReadText(element, "VisaAddress");
        var signalTypeText = ReadText(element, "SignalType");
        var label = $"{position}番目の測定器（{(name.Length > 0 ? name : "名前なし")}）";

        if (category.Length == 0 || name.Length == 0 || visaAddress.Length == 0) {
            warnings.Add($"{label}: Category、Name、VisaAddress のいずれかが空のため、読み飛ばしました。");
            return null;
        }
        if (!int.TryParse(signalTypeText, NumberStyles.None, CultureInfo.InvariantCulture, out var signalType) || signalType is < 1 or > 4) {
            warnings.Add($"{label}: SignalType は 1〜4 の整数で指定してください（現在: '{signalTypeText}'）。読み飛ばしました。");
            return null;
        }
        return new InstrumentEntry(category, name, visaAddress, signalType);
    }

    // 子要素の文字を、前後の空白を除いて読む（要素がなければ空文字）
    private static string ReadText(XElement element, string name) {
        return element.Elements().FirstOrDefault(child => child.Name.LocalName == name)?.Value.Trim() ?? string.Empty;
    }
}
