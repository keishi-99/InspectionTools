namespace InspectionTools.App.ViewModels;

/// <summary>ユーザーへの確認・通知の窓口（ViewModel から画面の部品を直接触らず、テストで差し替えられるようにする）</summary>
public interface IDialogService {

    /// <summary>「はい／いいえ」で答えてもらう。はいなら true</summary>
    bool Confirm(string message);

    /// <summary>エラーを知らせる</summary>
    void ShowError(string message);

    /// <summary>処理が終わったことなどを知らせる</summary>
    void ShowInfo(string message);

    /// <summary>転記先のテンプレート（Excel ブック）を選んでもらう。選ばずに閉じたときは null</summary>
    string? SelectTemplateFilePath();
}
