namespace InspectionTools.Excel;

/// <summary>テンプレートのどのシートの、どの行を読むか</summary>
/// <param name="DataSheet">項目名が書かれた値の入力シート（EL5000AC では ACdata）</param>
/// <param name="SettingSheet">合格範囲が書かれたシート（EL5000AC では ACsetting）</param>
/// <param name="FirstRow">読み始める行</param>
/// <param name="LastRow">読み終える行</param>
public sealed record TemplateLayout(string DataSheet, string SettingSheet, int FirstRow, int LastRow);
