using System.Windows;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        private bool _isHelpVisible = false;
        private string _pageName = string.Empty;

        private MainMenu.SubMenuUserControl? _subMenu;

        public MainWindow() {
            InitializeComponent();
            Common.HelpManager.LoadHelpFile("help.json");
            ShowMainMenu();
        }

        private void ShowMainMenu() {

            _subMenu = new MainMenu.SubMenuUserControl();
            _subMenu.BackToMainRequested += (_, __) => ShowMainMenu();
            _subMenu.HelpButtonClicked += OnHelpButtonClicked;
            _subMenu.SetButtonEnabled("ProductListButton", false);
            _subMenu.SetButtonEnabled("InstListButton", true);
            SubMenuContentArea.Content = _subMenu;

            var mainMenu = new MainMenu.MainMenuUserControl();
            mainMenu.PageSelected += OnPageSelected;
            MainMenuContentArea.Content = mainMenu;

            this.Title = "Menu";
            _pageName = "MainMenu";
            if (_isHelpVisible) {
                var (keys, descriptions) = Common.HelpManager.GetHelpData(_pageName);
                HelpTextBlock1.Text = string.Join(Environment.NewLine, keys);
                HelpTextBlock2.Text = string.Join(Environment.NewLine, descriptions);
            }
            HotKeyHelpScrollViewer.Height = mainMenu.Height + (_subMenu?.Height ?? 0);
        }

        private void OnPageSelected(string pageName) {

            _subMenu?.SetButtonEnabled("ProductListButton", true);

            UserControl? page = pageName switch {
                "EL0122FI" => new Product.EL0122FIUserControl(),
                "EL0122" => new Product.EL0122UserControl(),
                "EL0137" => new Product.EL0137UserControl(),
                "EL1812" => new Product.EL1812UserControl(),
                "EL3801" => new Product.EL3801UserControl(),
                "EL4001" => new Product.EL4001UserControl(),
                "EL9100" => new Product.EL9100UserControl(),
                "EL9240" => new Product.EL9240UserControl(),

                "PA14" => new Product.PA14UserControl(),
                "PAF5amp" => new Product.PAF5ampUserControl(),
                "PAF5" => new Product.PAF5UserControl(),

                "DFPDX" => new Product.DFPDXUserControl(),
                "MassFlow" => new Product.MassFlowUserControl(),
                _ => null
            };

            if (page is not null) {

                this.Title = pageName;
                _pageName = pageName;
                MainMenuContentArea.Content = page;

                if (page is MainMenu.SubMenuUserControl.ISubMenuAware s) {
                    s.SetSubMenuControl(_subMenu);
                }

                if (_isHelpVisible) {
                    var (keys, descriptions) = Common.HelpManager.GetHelpData(_pageName);
                    HelpTextBlock1.Text = string.Join(Environment.NewLine, keys);
                    HelpTextBlock2.Text = string.Join(Environment.NewLine, descriptions);
                }
                HotKeyHelpScrollViewer.Height = page.Height + (_subMenu?.ActualHeight ?? 0);
            }
        }
        private void OnHelpButtonClicked(object? sender, EventArgs e) {
            _isHelpVisible = !_isHelpVisible;

            var margin = _isHelpVisible ? new Thickness(10) : new Thickness(0);
            HelpTextBlock1.Margin = margin;
            HelpTextBlock2.Margin = margin;

            if (_isHelpVisible) {
                var (keys, descriptions) = Common.HelpManager.GetHelpData(_pageName);
                HelpTextBlock1.Text = string.Join(Environment.NewLine, keys);
                HelpTextBlock2.Text = string.Join(Environment.NewLine, descriptions);
            }
            else {
                HelpTextBlock1.Text = string.Empty;
                HelpTextBlock2.Text = string.Empty;
            }
        }


    }
}