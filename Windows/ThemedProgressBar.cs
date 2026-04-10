using System.Drawing;
using System.Windows.Forms;

namespace LibreDocToPdf
{
    public class ThemedProgressBar : ProgressBar
    {
        public Color BarColor { get; set; } = Color.FromArgb(0, 120, 212);
        public Color BarBackColor { get; set; } = Color.FromArgb(230, 230, 230);
        public Color BorderColor { get; set; } = Color.FromArgb(180, 180, 180);

        public ThemedProgressBar()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle rect = ClientRectangle;

            using (var bgBrush = new SolidBrush(BarBackColor))
                e.Graphics.FillRectangle(bgBrush, rect);

            if (Maximum > Minimum && Value > Minimum)
            {
                float fraction = (float)(Value - Minimum) / (Maximum - Minimum);
                int fillWidth = (int)(rect.Width * fraction);
                if (fillWidth > 0)
                {
                    using (var fillBrush = new SolidBrush(BarColor))
                        e.Graphics.FillRectangle(fillBrush, rect.X, rect.Y, fillWidth, rect.Height);
                }
            }

            using (var pen = new Pen(BorderColor))
                e.Graphics.DrawRectangle(pen, 0, 0, rect.Width - 1, rect.Height - 1);
        }
    }
}
