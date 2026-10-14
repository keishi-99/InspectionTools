using System.IO;
using System.Windows;
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
            var store = new SessionStore(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "InspectionTools", "sessions"));

            // 測定器の一覧（VisaAddress.xml）は、実行ファイルのフォルダに置く（開くたびに読み直す）
            var procedure = El5000AcProcedure.Create();
            var instrumentSelection = new InstrumentSelectionViewModel(
                procedure,
                new InstrumentConnection(procedure, new DeviceControllerConnection()),
                Path.Combine(AppContext.BaseDirectory, "VisaAddress.xml"));

            var viewModel = new MainViewModel(definition, store, dialog, instrumentSelection);
            viewModel.Start();

            var window = new MainWindow(viewModel);
            dialog.Owner = window;
            window.Show();
        } catch (Exception ex) when (ex is ProductDefinitionException or IOException) {
            dialog.ShowError($"製品の定義を読み込めませんでした。\n{ex.Message}");
            Shutdown();
        }
    }
}
