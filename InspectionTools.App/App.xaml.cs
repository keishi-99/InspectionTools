using System.IO;
using System.Windows;
using InspectionTools.App.Input;
using InspectionTools.App.ViewModels;
using InspectionTools.Common;
using InspectionTools.Inspection;
using InspectionTools.Products;

namespace InspectionTools.App;

public partial class App : Application {

    protected override void OnStartup(StartupEventArgs e) {
        base.OnStartup(e);

        var dialog = new WpfDialogService();

        // 製品の一覧は、products フォルダの定義ファイルから作る。途中保存も設定ファイルも、実行ファイルのフォルダに置く（AppData には置かない）
        var catalog = ProductCatalog.Scan(Path.Combine(AppContext.BaseDirectory, "products"));
        var store = new SessionStore(Path.Combine(AppContext.BaseDirectory, "sessions"));
        var shell = new ShellViewModel(catalog, item => OpenProduct(item, store, dialog), dialog);

        // キー割り当ては、実行ファイルのフォルダの設定ファイルだけで決まる。書き間違いや、キーのない操作は、起動時に知らせる
        var keyBindings = KeyBindingLoader.Load(Path.Combine(AppContext.BaseDirectory, KeyBindingLoader.FileName));
        if (keyBindings.Warnings.Count > 0) {
            dialog.ShowError("キー割り当ての設定に、確認が必要な点があります。\n" + string.Join("\n", keyBindings.Warnings.Select(warning => $"・{warning}")));
        }

        var window = new MainWindow(shell, keyBindings.Bindings);
        dialog.Owner = window;
        window.Show();
    }

    // 選ばれた製品の検査画面を作る。手順表のある製品だけ、測定器の選択と接続を付ける（測定器の一覧 VisaAddress.xml は、開くたびに読み直す）
    private static MainViewModel OpenProduct(ProductMenuItem item, SessionStore store, WpfDialogService dialog) {
        var definition = ProductDefinitionLoader.LoadFile(item.DefinitionPath);
        InstrumentSelectionViewModel? instrumentSelection = null;
        var procedure = ProductProcedures.TryCreate(definition.Product);
        if (procedure is not null) {
            instrumentSelection = new InstrumentSelectionViewModel(
                procedure,
                new InstrumentConnection(procedure, new DeviceControllerConnection()),
                Path.Combine(AppContext.BaseDirectory, "VisaAddress.xml"));
        }
        return new MainViewModel(definition, store, dialog, instrumentSelection);
    }
}
