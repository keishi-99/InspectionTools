using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace InspectionTools.Tool {
    /// <summary>
    /// WaitingCircle.xaml の相互作用ロジック
    /// </summary>
    public partial class WaitingCircle : UserControl {
        public static readonly DependencyProperty s_circleColorProperty =
            DependencyProperty.Register(
                "CircleColor", // プロパティ名を指定
                typeof(Color), // プロパティの型を指定
                typeof(WaitingCircle), // プロパティを所有する型を指定
                new UIPropertyMetadata(Color.FromRgb(90, 117, 153),
                    (d, e) => { ((WaitingCircle)d).OnCircleColorPropertyChanged(e); }));
        public Color CircleColor {
            get => (Color)GetValue(s_circleColorProperty); set => SetValue(s_circleColorProperty, value);
        }

        public WaitingCircle() {
            InitializeComponent();

            double cx = 50.0;
            double cy = 50.0;
            double r = 45.0;
            int cnt = 14;
            double deg = 360.0 / cnt;
            double degS = deg * 0.2;
            for (int i = 0; i < cnt; ++i) {
                var si1 = Math.Sin((270.0 - (i * deg)) / 180.0 * Math.PI);
                var co1 = Math.Cos((270.0 - (i * deg)) / 180.0 * Math.PI);
                var si2 = Math.Sin((270.0 - ((i + 1) * deg) + degS) / 180.0 * Math.PI);
                var co2 = Math.Cos((270.0 - ((i + 1) * deg) + degS) / 180.0 * Math.PI);
                var x1 = (r * co1) + cx;
                var y1 = (r * si1) + cy;
                var x2 = (r * co2) + cx;
                var y2 = (r * si2) + cy;

                var path = new Path {
                    Data = Geometry.Parse(string.Format("M {0},{1} A {2},{2} 0 0 0 {3},{4}", x1, y1, r, x2, y2)),
                    Stroke = new SolidColorBrush(Color.FromArgb((byte)(255 - (i * 256 / cnt)), CircleColor.R, CircleColor.G, CircleColor.B)),
                    StrokeThickness = 10.0
                };
                MainCanvas.Children.Add(path);
            }

            var kf = new DoubleAnimationUsingKeyFrames {
                RepeatBehavior = RepeatBehavior.Forever
            };
            for (int i = 0; i < cnt; ++i) {
                kf.KeyFrames.Add(new DiscreteDoubleKeyFrame() {
                    KeyTime = KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(i * 80)),
                    Value = i * deg
                });
            }
            kf.KeyFrames.Add(new DiscreteDoubleKeyFrame() {
                KeyTime = KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(cnt * 80)),
                Value = 0
            });
            MainTrans.BeginAnimation(RotateTransform.AngleProperty, kf);
        }

        public void OnCircleColorPropertyChanged(DependencyPropertyChangedEventArgs _) {
            if (null == MainCanvas) {
                return;
            }

            if (null == MainCanvas.Children) {
                return;
            }

            foreach (var child in MainCanvas.Children) {
                if (child is Shape shp && shp.Stroke is SolidColorBrush sb) {
                    var a = sb.Color.A;
                    shp.Stroke = new SolidColorBrush(Color.FromArgb(a, CircleColor.R, CircleColor.G, CircleColor.B));
                }
            }
        }
    }
}
