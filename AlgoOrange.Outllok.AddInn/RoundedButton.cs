using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace AlgoOrange.Outllok.AddInn
{
    public class RoundedButton : Button
    {
        public int CornerRadius { get; set; } = 20; // Default corner radius

        protected override void OnPaint(PaintEventArgs pevent)
        {
            base.OnPaint(pevent);
            GraphicsPath graphicsPath = new GraphicsPath();
            int radius = CornerRadius;
            graphicsPath.AddArc(0, 0, radius, radius, 180, 90);
            graphicsPath.AddArc(ClientSize.Width - radius, 0, radius, radius, 270, 90);
            graphicsPath.AddArc(ClientSize.Width - radius, ClientSize.Height - radius, radius, radius, 0, 90);
            graphicsPath.AddArc(0, ClientSize.Height - radius, radius, radius, 90, 90);
            graphicsPath.CloseFigure();
            this.Region = new Region(graphicsPath);
            pevent.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            pevent.Graphics.FillPath(new SolidBrush(this.BackColor), graphicsPath);
            pevent.Graphics.DrawPath(new Pen(this.ForeColor), graphicsPath);

            // Draw the text
            TextRenderer.DrawText(pevent.Graphics, this.Text, this.Font, this.ClientRectangle, this.ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }
}
