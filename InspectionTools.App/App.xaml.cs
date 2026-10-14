using System.IO;
using System.Windows;
using InspectionTools.App.Input;
using InspectionTools.App.ViewModels;
using InspectionTools.Common;
using InspectionTools.Inspection;
using InspectionTools.Products;

namespace InspectionTools.App;

public partial class App : Application {

    // 5-1 では製品を EL5000AC に固定する（製品の選択メニューは後の手順で追加する）
    private const string ProductName = "EL5000AC";

    protected override void OnStartup(StartupEventArgs e) {
        base.OnStartup(e);

        var dialog = new WpfDialogService();
        try {
            var definition = ProductDefinitionLoader.LoadFile(
                Path.Combine(AppContext.BaseDirectory, "products", ProductName + ".json"));
            // 途中保存も設定ファイルも、実行ファイルのフォルダに置く（AppData には置かない）
            var store = new SessionStore(Path.Combine(AppContext.BaseDirectory, "sessions"));

            // 測定器の一覧（VisaAddress.xml）は、実行ファイルのフォルダに置く（開くたびに読み直す）
            var procedure = El5000AcProcedure.Create();
            var instrumentSelection = new InstrumentSelectionViewModel(
                procedure,
                new InstrumentConnection(procedure, new DeviceControllerConnection()),
                Path.Combine(AppContext.BaseDirectory, "VisaAddress.xml"));

            var viewModel = new MainViewModel(definition, store, dialog, instrumentSelection);
            viewModel.Start();

            // キー割り当ては、実行ファイルのフォルダの設定ファイルだけで決まる。書き間違いや、キーのない操作は、起動時に知らせる
            var keyBindings = KeyBindingLoader.Load(Path.Combine(AppContext.BaseDirectory, KeyBindingLoader.FileName));
            if (keyBindings.Warnings.Count > 0) {
                dialog.ShowError("キー割り当ての設定に、確認が必要な点があります。\n" + string.Join("\n", keyBindings.Warnings.Select(warning => $"・{warning}")));
            }

            var window = new MainWindow(viewModel, keyBindings.Bindings);
            dialog.Owner = window;
            window.Show();
        } catch (Exception ex) when (ex is ProductDefinitionException or IOException) {
            dialog.ShowError($"製品の定義を読み込めませんでした。\n{ex.Message}");
            Shutdown();
        }
    }
}
