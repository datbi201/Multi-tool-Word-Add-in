using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace TroLyKyThuat
{
    // Dựng hình từ cùng cây công thức dùng khi chèn Equation vào Word.
    sealed class EquationPreview : Control
    {
        readonly bool stacked;
        MNode formula;

        public EquationPreview(bool stacked)
        {
            this.stacked = stacked;
            SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Color.White;
            Font = new Font("Cambria Math", 16f);
        }

        public void Set(MNode node, string text)
        {
            formula = node;
            AccessibleName = Lang.Msg("Xem trước " + (stacked ? "riêng dòng: " : "cùng dòng: ")) + (text ?? "");
            Invalidate();
        }

        sealed class Box
        {
            public readonly float Width, Above, Below;
            public readonly Action<float, float> Draw;
            public Box(float width, float above, float below, Action<float, float> draw)
            {
                Width = width; Above = above; Below = below; Draw = draw;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (formula == null || Width <= 8 || Height <= 8) return;
            Graphics g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            using (var ink = new SolidBrush(ForeColor))
            using (var pen = new Pen(ForeColor))
            using (var format = StringFormat.GenericTypographic)
            {
                format.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces | StringFormatFlags.NoWrap;
                Box layout = BuildBox(formula, Font.SizeInPoints, g, ink, pen, format);
                float scale = Math.Min(1, Math.Min((Width - 8f) / Math.Max(1, layout.Width),
                    (Height - 8f) / Math.Max(1, layout.Above + layout.Below)));
                g.TranslateTransform((Width - layout.Width * scale) / 2,
                    (Height - (layout.Above + layout.Below) * scale) / 2);
                g.ScaleTransform(scale, scale);
                layout.Draw(0, layout.Above);
            }
        }

        static Box Row(IEnumerable<Box> children)
        {
            var boxes = new List<Box>(children);
            float width = 0, above = 0, below = 0;
            foreach (Box box in boxes)
            {
                width += box.Width;
                above = Math.Max(above, box.Above);
                below = Math.Max(below, box.Below);
            }
            return new Box(width, above, below, delegate(float x, float y)
            {
                foreach (Box box in boxes) { box.Draw(x, y); x += box.Width; }
            });
        }

        Box BuildBox(MNode node, float size, Graphics g, Brush ink, Pen pen, StringFormat format)
        {
            Func<MNode, float, Box> build = (n, pt) => BuildBox(n, pt, g, ink, pen, format);
            Func<string, bool, Box> text = delegate(string value, bool upright)
            {
                FontStyle style = upright ? FontStyle.Regular : FontStyle.Italic;
                float width, above, below;
                using (var font = new Font(Font.FontFamily, size, style))
                {
                    width = g.MeasureString(value, font, int.MaxValue, format).Width;
                    above = font.GetHeight(g) * 0.67f;
                    below = font.GetHeight(g) - above;
                }
                return new Box(width, above, below, delegate(float x, float y)
                {
                    using (var font = new Font(Font.FontFamily, size, style))
                        g.DrawString(value, font, ink, x, y - above, format);
                });
            };
            if (node == null) return Row(new Box[0]);
            switch (node.K)
            {
                case MK.Text: return text(node.T, node.Up);
                case MK.Seq: return Row(node.Kids.ConvertAll(n => build(n, size)));
                case MK.Paren: return Row(new[] { text("(", true), build(node.A, size), text(")", true) });
                case MK.Frac:
                    if (!stacked)
                        return Row(new[] { build(node.A, size), text("/", true), build(node.B, size) });
                    var num = build(Omml.Strip(node.A), size * 0.85f);
                    var den = build(Omml.Strip(node.B), size * 0.85f);
                    float width = Math.Max(num.Width, den.Width) + 6;
                    return new Box(width, num.Above + num.Below + 2, den.Above + den.Below + 2,
                        delegate(float x, float y)
                        {
                            num.Draw(x + (width - num.Width) / 2, y - 2 - num.Below);
                            g.DrawLine(pen, x, y, x + width, y);
                            den.Draw(x + (width - den.Width) / 2, y + 2 + den.Above);
                        });
                case MK.Rad:
                    Box radicand = build(Omml.Strip(node.A), size);
                    float rootWidth = text("√", true).Width;
                    return new Box(rootWidth + radicand.Width + 2, radicand.Above + 3, radicand.Below,
                        delegate(float x, float y)
                        {
                            float top = y - radicand.Above - 2;
                            g.DrawLines(pen, new[] { new PointF(x, y), new PointF(x + rootWidth * 0.3f, y),
                                new PointF(x + rootWidth * 0.55f, y + radicand.Below * 0.7f),
                                new PointF(x + rootWidth, top), new PointF(x + rootWidth + radicand.Width + 2, top) });
                            radicand.Draw(x + rootWidth + 1, y);
                        });
                case MK.Acc:
                    Box body = build(node.A, size);
                    Box accent = text(" " + node.T, true);
                    return new Box(body.Width, body.Above + 4, body.Below, delegate(float x, float y)
                    {
                        body.Draw(x, y);
                        accent.Draw(x + (body.Width - accent.Width) / 2, y - body.Above * 0.42f);
                    });
                default:
                    Box main = build(node.A, size);
                    Box sub = build(node.B, size * 0.7f), sup = build(node.C, size * 0.7f);
                    float up = main.Above * 0.6f, down = main.Below * 0.75f;
                    return new Box(main.Width + Math.Max(sub.Width, sup.Width),
                        Math.Max(main.Above, sup.Above + up), Math.Max(main.Below, sub.Below + down),
                        delegate(float x, float y)
                        {
                            main.Draw(x, y);
                            sub.Draw(x + main.Width, y + down);
                            sup.Draw(x + main.Width, y - up);
                        });
            }
        }
    }
}