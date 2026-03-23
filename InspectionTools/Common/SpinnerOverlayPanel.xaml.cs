using System.Windows.Media;
using System.Windows.Media.Animation;
using UserControl = System.Windows.Controls.UserControl;

namespace InspectionTools.Tool {
    /// <summary>
    /// SpinnerOverlayPanel.xaml の相互作用ロジック
    /// </summary>
    public partial class SpinnerOverlayPanel : UserControl {
        public SpinnerOverlayPanel() {
            InitializeComponent();

            // スピナーを約1秒で1回転するアニメーションを開始（OverlayPanel と同速: 6°/16ms ≈ 375°/s）
            var animation = new DoubleAnimation {
                From = 0,
                To = 360,
                Duration = TimeSpan.FromMilliseconds(960),
                RepeatBehavior = RepeatBehavior.Forever
            };
            SpinnerRotation.BeginAnimation(RotateTransform.AngleProperty, animation);
        }
    }
}
