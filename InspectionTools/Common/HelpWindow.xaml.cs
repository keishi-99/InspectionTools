using System.Collections.Generic;
using System.Windows;

namespace InspectionTools.Common {
    public partial class HelpWindow : Window {
        public HelpWindow() {
            InitializeComponent();

            // コンテンツに合わせて高さを自動調整するが、画面の作業領域を超える場合は上限を設ける
            Loaded += (s, e) => {
                double workAreaHeight = System.Windows.SystemParameters.WorkArea.Height;
                if (Height > workAreaHeight) {
                    SizeToContent = SizeToContent.Manual;
                    Height = workAreaHeight;
                }
            };
        }

        // 表示するヘルプエントリ一覧を更新する
        public void UpdateHelpData(IReadOnlyList<HelpEntry> entries) {
            HelpItemsControl.ItemsSource = entries;
        }
    }
}
