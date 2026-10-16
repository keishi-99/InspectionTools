using System.IO;
using System.Windows;
using InspectionTools.App.Input;
using InspectionTools.App.Settings;
using InspectionTools.App.ViewModels;
using InspectionTools.Common;
using InspectionTools.Inspection;
using InspectionTools.Products;

namespace InspectionTools.App;

public partial class App : Application {

    // 他のユーザーのセッションとも取り合わないように、Global を付ける
    private SingleInstanceGuard? _singleInstance;

    // 画面の配色(ライト/ダーク)
    private ThemeController? _themeController;

    protected override void OnStartup(StartupEventArgs e) {
        base.OnStartup(e);

        // 2つ目の起動は、知らせて終了する
        _singleInstance = new SingleInstanceGuard(@"Global\InspectionTools.App");
        if (!_singleInstance.IsFirstInstance) {
            MessageBox.Show("すでに起動しています。\n起動中のアプリを使ってください。", "多重起動", MessageBoxButton.OK, MessageBoxImage.Information);
            Shutdown();
            return;
        }

        var dialog = new WpfDialogService();

        // 配色は、実行ファイルのフォルダの settings.json から読み、ウィンドウを作る前に適用する(切り替えると、ここに保存する)
        _themeController = new ThemeController(Path.Combine(AppContext.BaseDirectory, AppSettingsStore.FileName), new MaterialThemeApplier(), dialog);
        _themeController.Initialize();

        // 製品の一覧は、products フォルダの定義ファイルから作る。途中保存も設定ファイルも、実行ファイルのフォルダに置く（AppData には置かない）
        var catalog = ProductCatalog.Scan(Path.Combine(AppContext.BaseDirectory, "products"));
        var store = new SessionStore(Path.Combine(AppContext.BaseDirectory, "sessions"));
        var shell = new ShellViewModel(catalog, item => OpenProduct(item, store, dialog), dialog);

        // キー割り当ては、実行ファイルのフォルダの設定ファイルだけで決まる。書き間違いや、キーのない操作は、起動時に知らせる
        var keyBindings = KeyBindingLoader.Load(Path.Combine(AppContext.BaseDirectory, KeyBindingLoader.FileName));
        if (keyBindings.Warnings.Count > 0) {
            dialog.ShowWarning("キー割り当ての設定に、確認が必要な点があります。\n" + string.Join("\n", keyBindings.Warnings.Select(warning => $"・{warning}")));
        }

        var window = new MainWindow(shell, keyBindings.Bindings, dialog, _themeController);
        dialog.Owner = window;
        window.Show();
    }

    // 選ばれた製品の検査画面を作る。手順表のある製品だけ、測定器の選択と接続を付ける（測定器の一覧 VisaAddress.xml は、開くたびに読み直す）
    private static MainViewModel OpenProduct(ProductMenuItem item, SessionStore store, WpfDialogService dialog) {
        var definition = ProductDefinitionLoader.LoadFile(item.DefinitionPath);
        InstrumentSelectionViewModel? instrumentSelection = null;
        var procedure = ProductProcedures.TryCreate(definition.Product);
        if (procedure is not null) {
            // 定義の actionId が手順表にないと、検査中に例外になる。開く時点で見つけて、メニューに戻す
            ProductDefinitionChecker.EnsureActionsExist(definition, procedure);
            instrumentSelection = new InstrumentSelectionViewModel(
                procedure,
                new InstrumentConnection(procedure, new DeviceControllerConnection()),
                Path.Combine(AppContext.BaseDirectory, "VisaAddress.xml"));
        }
        return new MainViewModel(definition, store, dialog, instrumentSelection);
    }

    protected override void OnExit(ExitEventArgs e) {
        _singleInstance?.Dispose();
        base.OnExit(e);
    }
}
