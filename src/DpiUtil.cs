// ============================================================================
//  DpiUtil.cs - Hỗ trợ DPI sắc nét (Per-Monitor V2) cho mọi Form (Rev0.9.3, bổ sung Rev0.9.5.2)
//  Rev0.9.5.2: thêm theo dõi DPI từng màn hình. Cửa sổ Per-Monitor V2 nhưng WinForms (không có app.config)
//  không tự co giãn khi cửa sổ sang màn hình khác tỷ lệ, nên chữ và nút giữ cỡ 96 DPI: trông nhỏ, mảnh, thiếu nét.
//  Nay form tự đổi cỡ chữ và kích thước theo DPI thật của màn hình đang hiện (WM_DPICHANGED).
// ============================================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TroLyKyThuat
{
    /// <summary>
    /// Quản lý ngữ cảnh DPI Per-Monitor V2 của Windows khi tạo Form WinForms,
    /// tránh hiện tượng Windows bitmap-stretch làm mờ hoặc vỡ chữ trên màn hình HiDPI,
    /// và co giãn cỡ chữ, kích thước theo từng màn hình.
    /// </summary>
    public static class DpiUtil
    {
        // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = (HANDLE)-4
        static readonly IntPtr PerMonitorV2 = new IntPtr(-4);
        // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE    = (HANDLE)-3
        static readonly IntPtr PerMonitorV1 = new IntPtr(-3);
        const int WM_DPICHANGED = 0x02E0;

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetWindowDpiAwarenessContext(IntPtr hwnd);

        [DllImport("user32.dll")]
        static extern int GetAwarenessFromDpiAwarenessContext(IntPtr value);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT { public int Left, Top, Right, Bottom; }

        sealed class DpiState { public int Dpi = 96; public int Zoom = 100; }
        static readonly ConditionalWeakTable<Form, DpiState> States = new ConditionalWeakTable<Form, DpiState>();

        sealed class ContextScope : IDisposable
        {
            readonly IntPtr prev;
            readonly bool active;
            public ContextScope(IntPtr prev, bool active) { this.prev = prev; this.active = active; }
            public void Dispose()
            {
                if (!active || prev == IntPtr.Zero) return;
                try { SetThreadDpiAwarenessContext(prev); }
                catch { }
            }
        }

        /// <summary>
        /// Thử chuyển luồng hiện tại sang Per-Monitor V2 DPI (hoặc V1 nếu OS cũ hơn).
        /// Không ném ngoại lệ trên hệ điều hành chưa hỗ trợ API này.
        /// </summary>
        public static bool TryEnablePerMonitorV2(out IntPtr previousContext)
        {
            previousContext = IntPtr.Zero;
            try
            {
                previousContext = SetThreadDpiAwarenessContext(PerMonitorV2);
                if (previousContext != IntPtr.Zero) return true;
                previousContext = SetThreadDpiAwarenessContext(PerMonitorV1);
                return previousContext != IntPtr.Zero;
            }
            catch
            {
                previousContext = IntPtr.Zero;
                return false;
            }
        }

        /// <summary>Tạo phạm vi using (...) để dựng và tạo handle Form trong ngữ cảnh Per-Monitor V2 rồi hoàn trả ngữ cảnh cũ.</summary>
        public static IDisposable EnterPerMonitorV2()
        {
            IntPtr prev;
            bool ok = TryEnablePerMonitorV2(out prev);
            return new ContextScope(prev, ok);
        }

        /// <summary>Cấu hình chuẩn DPI cho mọi Form thuộc Add-in.</summary>
        public static void ConfigureForm(Form form, float ptSize)
        {
            if (form == null) return;
            form.AutoScaleDimensions = new SizeF(96f, 96f);
            form.AutoScaleMode = AutoScaleMode.Dpi;
            form.Font = CreateUiFont(ptSize, FontStyle.Regular);
        }

        /// <summary>Tạo font giao diện sắc nét (ưu tiên Segoe UI, fallback sang SystemFonts).</summary>
        public static Font CreateUiFont(float ptSize, FontStyle style)
        {
            try
            {
                return new Font("Segoe UI", ptSize, style, GraphicsUnit.Point);
            }
            catch
            {
                try { return new Font(SystemFonts.MessageBoxFont.FontFamily, ptSize, style, GraphicsUnit.Point); }
                catch { return SystemFonts.MessageBoxFont; }
            }
        }

        /// <summary>Lấy hệ số tỷ lệ DPI thực tế của cửa sổ (1.0f = 96 DPI).</summary>
        public static float GetScaleFactor(Control ctrl)
        {
            try
            {
                if (ctrl != null && ctrl.IsHandleCreated)
                {
                    uint dpi = GetDpiForWindow(ctrl.Handle);
                    if (dpi >= 96 && dpi <= 480) return dpi / 96f;
                }
            }
            catch { }
            return 1f;
        }

        // ------------------------------------------------------------------ Rev0.9.5.2: chẩn đoán và co giãn theo màn hình

        /// <summary>Chế độ nhận biết DPI của cửa sổ: 0 không nhận biết (Windows kéo giãn ảnh, chữ mờ), 1 theo hệ thống, 2 theo từng màn hình (sắc nét). -1 nếu không đọc được.</summary>
        public static int WindowAwareness(IntPtr hwnd)
        {
            try { return GetAwarenessFromDpiAwarenessContext(GetWindowDpiAwarenessContext(hwnd)); }
            catch { return -1; }
        }

        /// <summary>DPI thật của màn hình đang chứa cửa sổ (96 = 100%).</summary>
        public static int WindowDpi(IntPtr hwnd)
        {
            try { uint d = GetDpiForWindow(hwnd); return d >= 96 ? (int)d : 96; }
            catch { return 96; }
        }

        static int DeviceDpiOf(Control c)
        {
            try { int d = c.DeviceDpi; return d >= 96 ? d : 96; }
            catch { return 96; }
        }

        static DpiState StateOf(Form f)
        {
            DpiState st;
            if (!States.TryGetValue(f, out st)) { st = new DpiState(); st.Dpi = DeviceDpiOf(f); States.Add(f, st); }
            return st;
        }

        /// <summary>DPI mà bố cục của form đang được dựng theo (đổi sau mỗi lần co giãn).</summary>
        public static int LayoutDpi(Form f) { return StateOf(f).Dpi; }

        public static float LayoutScale(Form form)
        {
            DpiState state = StateOf(form);
            return state.Dpi / 96f * (state.Zoom / 100f);
        }

        public static void SetZoom(Form form, int percent, Action reflow = null)
        {
            if (percent < 50 || percent > 130) throw new ArgumentOutOfRangeException("percent");
            using (EnterPerMonitorV2()) ApplyZoom(form, percent, reflow);
        }

        static void ApplyZoom(Form form, int percent, Action reflow)
        {
            DpiState state = StateOf(form);
            if (state.Zoom == percent) return;
            float factor = (float)percent / state.Zoom;
            var scrolls = new List<KeyValuePair<ScrollableControl, Point>>();
            CollectScrolls(form, scrolls);
            state.Zoom = percent;
            ScaleForm(form, factor, Rectangle.Empty, false);
            if (reflow != null) reflow();
            foreach (var scroll in scrolls)
                scroll.Key.AutoScrollPosition = new Point(
                    (int)Math.Round(-scroll.Value.X * factor),
                    (int)Math.Round(-scroll.Value.Y * factor));
        }

        static void CollectScrolls(Control parent, List<KeyValuePair<ScrollableControl, Point>> scrolls)
        {
            foreach (Control child in parent.Controls)
            {
                var scroll = child as ScrollableControl;
                if (scroll != null && scroll.AutoScroll && scroll.Visible)
                    scrolls.Add(new KeyValuePair<ScrollableControl, Point>(scroll, scroll.AutoScrollPosition));
                CollectScrolls(child, scrolls);
            }
        }

        /// <summary>Mô tả ngắn để ghi vào log chẩn đoán.</summary>
        public static string Describe(Form f)
        {
            if (f == null) return "form null";
            if (!f.IsHandleCreated) return "chưa có handle";
            return "dpiMànHình=" + WindowDpi(f.Handle) + " nhậnBiếtDPI=" + WindowAwareness(f.Handle)
                + " dpiBốCục=" + LayoutDpi(f) + " font=" + f.Font.Name + " " + f.Font.SizeInPoints.ToString("0.0") + "pt vùng=" + f.Bounds;
        }

        /// <summary>
        /// Gọi trong WndProc của Form. Khi Windows báo DPI đổi (kéo cửa sổ sang màn hình khác tỷ lệ), co giãn cỡ chữ và kích thước rồi dùng vùng Windows gợi ý.
        /// Trả true nếu đã xử lý.
        /// </summary>
        public static bool HandleDpiMessage(Form f, ref Message m)
        {
            if (m.Msg != WM_DPICHANGED) return false;
            try
            {
                int newDpi = (int)((long)m.WParam & 0xFFFF);
                RECT r = (RECT)Marshal.PtrToStructure(m.LParam, typeof(RECT));
                ApplyDpi(f, newDpi, new Rectangle(r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top), true);
            }
            catch (Exception ex) { Log.W("DPI đổi: lỗi " + ex.Message); }
            m.Result = IntPtr.Zero;
            return true;
        }

        /// <summary>Gọi khi form vừa hiện: nếu màn hình thật có DPI khác DPI bố cục thì co giãn ngay (trường hợp form dựng ở màn hình chính rồi hiện ở màn hình khác).</summary>
        public static void SyncDpi(Form f)
        {
            try
            {
                if (f == null || !f.IsHandleCreated) return;
                ApplyDpi(f, WindowDpi(f.Handle), Rectangle.Empty, false);
            }
            catch (Exception ex) { Log.W("DPI đồng bộ: lỗi " + ex.Message); }
        }

        static void CollectOwnFonts(Control parent, List<KeyValuePair<Control, Font>> list)
        {
            foreach (Control c in parent.Controls)
            {
                if (!object.ReferenceEquals(c.Font, parent.Font)) list.Add(new KeyValuePair<Control, Font>(c, c.Font));
                CollectOwnFonts(c, list);
            }
        }

        static void ScaleColumns(Control parent, float k)
        {
            foreach (Control c in parent.Controls)
            {
                ListView lv = c as ListView;
                if (lv != null) foreach (ColumnHeader h in lv.Columns) h.Width = (int)Math.Round(h.Width * k);
                ScaleColumns(c, k);
            }
        }

        /// <summary>Giữ cửa sổ trong vùng làm việc của màn hình đang chứa nó (sau khi co giãn có thể cao hơn màn hình).</summary>
        static void ClampToScreen(Form f)
        {
            try
            {
                if (!f.IsHandleCreated) return;
                Rectangle wa = Screen.FromHandle(f.Handle).WorkingArea;
                Rectangle b = f.Bounds;
                int w = Math.Min(b.Width, wa.Width), h = Math.Min(b.Height, wa.Height);
                int x = Math.Max(wa.Left, Math.Min(b.Left, wa.Right - w));
                int y = Math.Max(wa.Top, Math.Min(b.Top, wa.Bottom - h));
                if (x != b.Left || y != b.Top || w != b.Width || h != b.Height) f.SetBounds(x, y, w, h);
            }
            catch { }
        }

        static Font Scaled(Font f, float k)
        {
            return new Font(f.FontFamily, f.Size * k, f.Style, f.Unit);
        }

        static void ApplyDpi(Form f, int newDpi, Rectangle suggested, bool useRect)
        {
            if (f == null || newDpi < 96 || newDpi > 480) return;
            DpiState st = StateOf(f);
            if (newDpi == st.Dpi)
            {
                if (useRect) f.Bounds = suggested;
                return;
            }
            float k = (float)newDpi / (float)st.Dpi;
            int oldDpi = st.Dpi;
            st.Dpi = newDpi;
            ScaleForm(f, k, suggested, useRect);
            Log.W("DPI: " + f.GetType().Name + " " + oldDpi + " -> " + newDpi + " (x" + k.ToString("0.00") + "), font " + f.Font.SizeInPoints.ToString("0.0") + "pt");
        }

        static void ScaleForm(Form f, float k, Rectangle suggested, bool useRect)
        {
            var own = new List<KeyValuePair<Control, Font>>();
            CollectOwnFonts(f, own);          // font riêng của từng điều khiển, lấy trước khi đổi font form
            f.SuspendLayout();
            try
            {
                f.Font = Scaled(f.Font, k);
                foreach (KeyValuePair<Control, Font> kv in own) kv.Key.Font = Scaled(kv.Value, k);
                ScaleColumns(f, k);         // trước khi co giãn kích thước: các cột co động sẽ được tính lại theo bề rộng mới trong ReflowAll
                f.Scale(new SizeF(k, k));
                if (useRect) f.Bounds = suggested;
                ClampToScreen(f);   // cửa sổ cao hơn 100% không được vượt khỏi màn hình nhỏ hơn
            }
            finally { f.ResumeLayout(true); }
        }

        /// <summary>Rev0.9.5.6: tắt ký tự phím tắt (&amp;) trên mọi Label, Button, CheckBox, RadioButton để chữ "Kiểm tra &amp; Lưu" hiện đủ dấu &amp;.</summary>
        public static void NoMnemonics(Control root)
        {
            if (root == null) return;
            try
            {
                var lb = root as Label; if (lb != null) { lb.UseMnemonic = false; lb.UseCompatibleTextRendering = false; }
                var bb = root as ButtonBase; if (bb != null) { bb.UseMnemonic = false; bb.UseCompatibleTextRendering = false; }
                foreach (Control c in root.Controls) NoMnemonics(c);
            }
            catch { }
        }
    }
}
