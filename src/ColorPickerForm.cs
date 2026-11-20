// Trợ lý Kỹ thuật - Rev0.9.3: bảng chọn màu thống nhất kiểu Word hiện đại (Tự động, Màu giao diện 10x6, Màu chuẩn 10x1, Thêm màu...) trong một bảng duy nhất.
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace TroLyKyThuat
{
    public sealed class ColorPickerForm : Form
    {
        public string SelectedHex { get; private set; }
        readonly string target;
        readonly string currentHex;
        readonly int[] customColors;
        readonly ToolTip tip = new ToolTip();
        bool openingMore;

        const int Cols = 10;
        const int Swatch = 18;
        const int GapX = 4;
        const int Pad = 10;

        Rectangle autoRect;
        Rectangle moreRect;
        readonly Rectangle[] themeRects = new Rectangle[WordPalette.ThemeCount];
        readonly Rectangle[] stdRects = new Rectangle[WordPalette.StdCount];
        int hoverKind = -1; // 0 = auto, 1 = theme, 2 = std, 3 = more
        int hoverIdx = -1;

        public ColorPickerForm(string target, string currentHex, int[] customColors)
        {
            this.target = target ?? "hdrBg";
            this.currentHex = (currentHex ?? "000000").ToUpperInvariant();
            this.customColors = customColors;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            BackColor = Color.White;
            DpiUtil.ConfigureForm(this, 9.5f);
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            KeyPreview = true;

            int gridW = Cols * Swatch + (Cols - 1) * GapX;
            int w = Pad * 2 + gridW;

            int y = Pad;
            autoRect = new Rectangle(Pad, y, gridW, 28);
            y += 32;

            y += 20; // nhãn "Màu giao diện"
            for (int r = 0; r < 6; r++)
            {
                for (int c = 0; c < Cols; c++)
                {
                    int idx = r * Cols + c;
                    themeRects[idx] = new Rectangle(Pad + c * (Swatch + GapX), y, Swatch, Swatch);
                }
                y += Swatch + (r == 0 ? 6 : 1);
            }
            y += 6;

            y += 20; // nhãn "Màu chuẩn"
            for (int c = 0; c < Cols; c++)
                stdRects[c] = new Rectangle(Pad + c * (Swatch + GapX), y, Swatch, Swatch);
            y += Swatch + 10;

            moreRect = new Rectangle(Pad, y, gridW, 28);
            y += 30 + Pad;

            ClientSize = new Size(w, y);

            Point pt = Cursor.Position;
            Rectangle wa = Screen.FromPoint(pt).WorkingArea;
            int fx = Math.Max(wa.Left + 4, Math.Min(pt.X - 12, wa.Right - Width - 4));
            int fy = Math.Max(wa.Top + 4, Math.Min(pt.Y + 8, wa.Bottom - Height - 4));
            Location = new Point(fx, fy);

            tip.AutoPopDelay = 12000;
            tip.InitialDelay = 250;

            Deactivate += delegate
            {
                // Rev0.9.5: không ghi đè kết quả khi đã chọn màu (đóng hộp cũng gây Deactivate, trước đây biến OK thành Cancel nên mọi lần chọn màu bị bỏ).
                if (!openingMore && Visible && SelectedHex == null && DialogResult != DialogResult.OK) { DialogResult = DialogResult.Cancel; Close(); }
            };
            KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); }
            };
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;   // Rev0.9.5.6: chữ mịn cạnh
            g.Clear(Color.White);

            using (var border = new Pen(Color.FromArgb(180, 180, 180)))
                g.DrawRectangle(border, 0, 0, Width - 1, Height - 1);

            // 1. Tự động
            DrawButtonRow(g, autoRect, hoverKind == 0, true, Lang.Msg(WordPalette.AutoLabel(target)), WordPalette.AutoHex(target));

            // 2. Màu giao diện
            DrawHeader(g, Pad, autoRect.Bottom + 6, Lang.T(WordPalette.ThemeSectionLabel));
            for (int i = 0; i < WordPalette.ThemeCount; i++)
                DrawSwatch(g, themeRects[i], WordPalette.ThemeHex(i), hoverKind == 1 && hoverIdx == i);

            // 3. Màu chuẩn
            int stdLabelY = themeRects[WordPalette.ThemeCount - 1].Bottom + 6;
            DrawHeader(g, Pad, stdLabelY, Lang.T(WordPalette.StdSectionLabel));
            for (int i = 0; i < WordPalette.StdCount; i++)
                DrawSwatch(g, stdRects[i], WordPalette.StdHex(i), hoverKind == 2 && hoverIdx == i);

            // 4. Thêm màu...
            using (var sep = new Pen(Color.FromArgb(225, 225, 225)))
                g.DrawLine(sep, Pad, moreRect.Top - 5, Width - Pad, moreRect.Top - 5);
            DrawButtonRow(g, moreRect, hoverKind == 3, false, Lang.T(WordPalette.MoreColorsLabel), null);
        }

        void DrawHeader(Graphics g, int x, int y, string text)
        {
            using (var f = new Font(Font, FontStyle.Bold))
            using (var b = new SolidBrush(Color.FromArgb(68, 68, 68)))
                g.DrawString(text, f, b, x, y);
        }

        void DrawButtonRow(Graphics g, Rectangle r, bool hovered, bool withSwatch, string text, string hex)
        {
            if (hovered)
            {
                using (var bg = new SolidBrush(Color.FromArgb(230, 242, 252))) g.FillRectangle(bg, r);
                using (var bd = new Pen(Color.FromArgb(0, 120, 212))) g.DrawRectangle(bd, r);
            }
            int tx = r.X + 6;
            if (withSwatch && hex != null)
            {
                var sw = new Rectangle(r.X + 6, r.Y + 6, 16, 16);
                using (var b = new SolidBrush(HexColor(hex))) g.FillRectangle(b, sw);
                using (var p = new Pen(Color.FromArgb(120, 120, 120))) g.DrawRectangle(p, sw);
                tx += 24;
            }
            else if (!withSwatch)
            {
                var sw = new Rectangle(r.X + 6, r.Y + 6, 16, 16);
                using (var b = new LinearGradientBrush(sw, Color.FromArgb(220, 60, 60), Color.FromArgb(40, 120, 220), 45f))
                    g.FillRectangle(b, sw);
                using (var p = new Pen(Color.FromArgb(120, 120, 120))) g.DrawRectangle(p, sw);
                tx += 24;
            }
            using (var b = new SolidBrush(Color.FromArgb(30, 30, 30)))
            using (var sf = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center })
                g.DrawString(text, Font, b, new RectangleF(tx, r.Y, r.Right - tx, r.Height), sf);
        }

        void DrawSwatch(Graphics g, Rectangle r, string hex, bool hovered)
        {
            Color c = HexColor(hex);
            using (var b = new SolidBrush(c)) g.FillRectangle(b, r);
            bool active = string.Equals(hex, currentHex, StringComparison.OrdinalIgnoreCase);
            using (var p = new Pen(Color.FromArgb(200, 200, 200))) g.DrawRectangle(p, r);
            if (hovered || active)
            {
                using (var p1 = new Pen(hovered ? Color.FromArgb(210, 71, 38) : Color.FromArgb(0, 120, 212), 2f))
                    g.DrawRectangle(p1, r.X - 1, r.Y - 1, r.Width + 2, r.Height + 2);
                using (var p2 = new Pen(Color.White, 1f))
                    g.DrawRectangle(p2, r.X + 1, r.Y + 1, r.Width - 2, r.Height - 2);
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int nk = -1, ni = -1; string t = null;
            if (autoRect.Contains(e.Location)) { nk = 0; t = Lang.Msg(WordPalette.AutoLabel(target)); }
            else if (moreRect.Contains(e.Location)) { nk = 3; t = Lang.T(WordPalette.MoreColorsLabel); }
            else
            {
                for (int i = 0; i < themeRects.Length; i++) if (themeRects[i].Contains(e.Location)) { nk = 1; ni = i; t = Lang.Msg(WordPalette.ThemeLabel(i)); break; }
                if (nk < 0) for (int i = 0; i < stdRects.Length; i++) if (stdRects[i].Contains(e.Location)) { nk = 2; ni = i; t = Lang.Msg(WordPalette.StdLabel(i)); break; }
            }
            if (nk != hoverKind || ni != hoverIdx)
            {
                hoverKind = nk; hoverIdx = ni;
                tip.SetToolTip(this, t ?? "");
                Invalidate();
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left) return;
            if (autoRect.Contains(e.Location)) { SelectedHex = WordPalette.AutoHex(target); DialogResult = DialogResult.OK; Close(); return; }
            for (int i = 0; i < themeRects.Length; i++)
                if (themeRects[i].Contains(e.Location)) { SelectedHex = WordPalette.ThemeHex(i); DialogResult = DialogResult.OK; Close(); return; }
            for (int i = 0; i < stdRects.Length; i++)
                if (stdRects[i].Contains(e.Location)) { SelectedHex = WordPalette.StdHex(i); DialogResult = DialogResult.OK; Close(); return; }
            if (moreRect.Contains(e.Location))
            {
                openingMore = true;
                try
                {
                    using (var d = new ColorDialog())
                    {
                        d.FullOpen = true; d.AnyColor = true;
                        d.Color = HexColor(currentHex);
                        if (customColors != null) d.CustomColors = customColors;
                        if (d.ShowDialog(this) == DialogResult.OK)
                        {
                            SelectedHex = string.Format("{0:X2}{1:X2}{2:X2}", d.Color.R, d.Color.G, d.Color.B);
                            DialogResult = DialogResult.OK;
                            Close();
                        }
                    }
                }
                finally { openingMore = false; }
            }
        }

        static Color HexColor(string hex)
        {
            try { return Color.FromArgb(Convert.ToInt32(hex.Substring(0, 2), 16), Convert.ToInt32(hex.Substring(2, 2), 16), Convert.ToInt32(hex.Substring(4, 2), 16)); }
            catch { return Color.Gray; }
        }
    }
}
