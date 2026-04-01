using System.Collections.Generic;
using System.Windows;

namespace InspectionTools.Common {
    public partial class HelpWindow : Window {
        public HelpWindow() {
            InitializeComponent();

            // 画面の作業領域を超えないよう最大高を設定する（ページ切り替え後の動的更新にも対応）
            MaxHeight = SystemParameters.WorkArea.Height;
        }

        // 表示するヘルプエントリ一覧を更新する
        public void UpdateHelpData(IReadOnlyList<HelpEntry> entries) {
            HelpItemsControl.ItemsSource = entries;
        }
    }
}
