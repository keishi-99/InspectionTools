using System.Drawing.Drawing2D;

namespace InspectionTools.Common {
    /// <summary>
    /// 処理中オーバーレイ表示パネル（WinForms用スピナー付き半透明オーバーレイ）
    /// </summary>
    public sealed class OverlayPanel : Panel {
        private readonly System.Windows.Forms.Timer _timer;
        private float _angle;

        // OnPaint が約60fps で呼ばれるため、Pen をフィールドにキャッシュしてGC負荷を軽減
        private readonly Pen _bgPen;
        private readonly Pen _arcPen;

        public OverlayPanel() {
            SetStyle(
                ControlStyles.SupportsTransparentBackColor |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.UserPaint |
                ControlStyles.OptimizedDoubleBuffer,
                true);
            BackColor = Color.Transparent;

            _bgPen = new Pen(Color.FromArgb(220, 220, 220), 5f) {
                StartCap = LineCap.Round,
                EndCap = LineCap.Round
            };
            _arcPen = new Pen(Color.SteelBlue, 5f) {
                StartCap = LineCap.Round,
                EndCap = LineCap.Round
            };

            // 約60fps でスピナーを回転
            _timer = new System.Windows.Forms.Timer { Interval = 16 };
            _timer.Tick += (_, _) => {
                _angle = (_angle + 6f) % 360f;
                Invalidate();
            };
            _timer.Start();
        }

        protected override void OnPaint(PaintEventArgs e) {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // スピナー（薄い背景円 + 回転する円弧）をコントロール中央に描画
            const int SpinnerSize = 96;
            var spinnerRect = new RectangleF(
                (Width - SpinnerSize) / 2f,
                (Height - SpinnerSize) / 2f,
                SpinnerSize,
                SpinnerSize);

            g.DrawEllipse(_bgPen, spinnerRect);
            g.DrawArc(_arcPen, spinnerRect, _angle, 90f);
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                _timer.Dispose();
                _bgPen.Dispose();
                _arcPen.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
