using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace InspectionTools.Common {
    /// <summary>
    /// ScreenCaptureWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class ScreenCaptureWindow : Window {


        private System.Windows.Point _position;
        private bool _trimEnable = false;
        private string? _model = null;
        private int _captureWidth = 0;
        private int _captureHeight = 0;

        private Bitmap? _capturedImage;

        public ScreenCaptureWindow() {
            InitializeComponent();
        }

        public Bitmap? Capture(string? model = null, int width = 0, int height = 0) {
            _model = model;

            (_captureWidth, _captureHeight) = model switch {
                "EL9100" => (width, height),
                _ => (0, 0),
            };

            // モーダルで開く
            ShowDialog();

            return _capturedImage;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e) {
            // プライマリスクリーンサイズの取得
            var screen = System.Windows.Forms.Screen.PrimaryScreen
                ?? throw new InvalidOperationException("プライマリスクリーンの取得に失敗しました。");

            // ウィンドウサイズの設定
            this.Left = screen.Bounds.Left;
            this.Top = screen.Bounds.Top;
            this.Width = screen.Bounds.Width;
            this.Height = screen.Bounds.Height;

            // ジオメトリサイズの設定
            this.ScreenArea.Geometry1 = new RectangleGeometry(new Rect(0, 0, screen.Bounds.Width, screen.Bounds.Height));
        }

        private void DrawingPath_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            if (sender is not Path path) {
                return;
            }

            // 開始座標を取得
            var point = e.GetPosition(path);
            _position = point;

            // マウスキャプチャの設定
            _trimEnable = true;
            this.Cursor = System.Windows.Input.Cursors.Cross;
            path.CaptureMouse();
        }

        private void DrawingPath_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) {
            if (sender is not Path path) {
                return;
            }

            // 現在座標を取得
            var point = e.GetPosition(path);

            // マウスキャプチャの終了
            _trimEnable = false;
            this.Cursor = System.Windows.Input.Cursors.Arrow;
            path.ReleaseMouseCapture();

            // 画面キャプチャ

            switch (_model) {
                case "EL9100":
                    _capturedImage = CaptureScreenEL9100(point);
                    break;
                default:
                    _capturedImage = CaptureScreen(point);
                    break;
            }

            // アプリケーションの終了
            this.Close();
        }

        private void DrawingPath_MouseMove(object sender, System.Windows.Input.MouseEventArgs e) {
            if (!_trimEnable) {
                return;
            }

            if (sender is not Path path) {
                return;
            }

            // 現在座標を取得
            var point = e.GetPosition(path);

            switch (_model) {
                case "EL9100":
                    DrawStrokeEL9100(point);
                    break;
                default:
                    DrawStroke(point);
                    break;
            }
        }

        private void DrawStroke(System.Windows.Point point) {
            // 矩形の描画
            var x = _position.X < point.X ? _position.X : point.X;
            var y = _position.Y < point.Y ? _position.Y : point.Y;
            var width = Math.Abs(point.X - _position.X);
            var height = Math.Abs(point.Y - _position.Y);
            this.ScreenArea.Geometry2 = new RectangleGeometry(new Rect(x, y, width, height));
        }

        private Bitmap? CaptureScreen(System.Windows.Point point) {
            // 座標変換
            var start = PointToScreen(_position);
            var end = PointToScreen(point);

            // キャプチャエリアの取得
            var x = start.X < end.X ? (int)start.X : (int)end.X;
            var y = start.Y < end.Y ? (int)start.Y : (int)end.Y;
            var width = (int)Math.Abs(end.X - start.X);
            var height = (int)Math.Abs(end.Y - start.Y);
            if (width == 0 || height == 0) {
                return null;
            }

            // スクリーンイメージの取得
            var bmp = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            using var graph = System.Drawing.Graphics.FromImage(bmp);
            // 画面をコピーする
            graph.CopyFromScreen(new System.Drawing.Point(x, y), new System.Drawing.Point(), bmp.Size);

            return bmp;
        }

        private void DrawStrokeEL9100(System.Windows.Point point) {
            // 矩形の描画
            var x = point.X;
            var y = point.Y;
            var width = Math.Abs(_captureWidth);
            var height = Math.Abs(_captureHeight);
            this.ScreenArea.Geometry2 = new RectangleGeometry(new Rect(x, y, width, height));
        }
        private Bitmap? CaptureScreenEL9100(System.Windows.Point point) {

            // キャプチャエリアの取得
            var x = (int)point.X;
            var y = (int)point.Y;
            var width = Math.Abs(_captureWidth);
            var height = Math.Abs(_captureHeight);
            if (width == 0 || height == 0) {
                return null;
            }

            // スクリーンイメージの取得
            var bmp = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            using var graph = System.Drawing.Graphics.FromImage(bmp);
            // 画面をコピーする
            graph.CopyFromScreen(new System.Drawing.Point(x, y), new System.Drawing.Point(), bmp.Size);

            return bmp;
        }
    }
}
