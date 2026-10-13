using InspectionTools.Inspection;

namespace InspectionTools.Excel;

/// <summary>検査結果をテンプレートへ転記し、成功したときだけ途中保存を削除する</summary>
public static class ResultTransfer {

    /// <summary>
    /// 転記が成功したら、その製品の途中保存を削除する。転記に失敗したときは例外がそのまま投げられ、途中保存は残る
    /// （残っていれば、原因を直してから転記をやり直せるため）
    /// </summary>
    public static void Transfer(string templatePath, TransferSettings settings, InspectionSession session, SessionStore store, string product) {
        TemplateWriter.Write(templatePath, settings, session);

        // 転記が済んだ結果が、次の起動で「続きから再開しますか」と出ないように、ここで初めて消す
        store.Delete(product);
    }
}
