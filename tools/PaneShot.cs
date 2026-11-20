// Công cụ chẩn đoán (Rev0.9.5.2): dựng Bảng điều khiển trên luồng UI riêng giống cách add-in làm, hiện lên màn hình,
// báo DPI và chế độ nhận biết DPI của cửa sổ rồi chụp đúng điểm ảnh trên màn hình ra PNG để xem độ sắc nét của chữ.
// Cách dùng: bin\pane-shot.exe <tệp.png> [pmv2|system|unaware] [chọn thẻ 0-5]
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using TroLyKyThuat;

static class PaneShot
{
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr h);
    [DllImport("user32.dll")] static extern IntPtr GetWindowDpiAwarenessContext(IntPtr h);
    [DllImport("user32.dll")] static extern int GetAwarenessFromDpiAwarenessContext(IntPtr c);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L, T, R, B; }

    static int Main(string[] args)
    {
        if (Environment.GetEnvironmentVariable("TROLY_LANG") == "en") Lang.Set("en");   // Rev0.9.6: chụp giao diện tiếng Anh (set TROLY_LANG=en)
        string outPng = args.Length > 0 ? args[0] : "pane-shot.png";
        string mode = args.Length > 1 ? args[1] : "pmv2";
        int tab = args.Length > 2 ? int.Parse(args[2]) : 0;
        int screenIdx = args.Length > 3 ? int.Parse(args[3]) : 0;
        int moveTo = args.Length > 4 ? int.Parse(args[4]) : -1;      // sau khi hiện, kéo sang màn hình này (-1: không kéo)
        string kind = args.Length > 5 ? args[5] : "pane";             // pane | settings
        string report = "";
        var t = new Thread(delegate()
        {
            try
            {
                // Mô phỏng đúng Connect.EnsurePane: luồng STA riêng, bật Per-Monitor V2, EnableVisualStyles, tạo form rồi hiện.
                if (mode == "pmv2") { IntPtr prev; DpiUtil.TryEnablePerMonitorV2(out prev); }
                else if (mode == "system") SetThreadDpiAwarenessContext(new IntPtr(-2));
                else if (mode == "unaware") SetThreadDpiAwarenessContext(new IntPtr(-1));
                Application.EnableVisualStyles();
                Form f = kind == "settings" ? (Form)new SettingsForm() : (Form)new PaneForm(new Rec("app"));
                IntPtr h = f.Handle;
                if (screenIdx > 0 && screenIdx < Screen.AllScreens.Length)
                {
                    // chuyển sang màn hình khác (khác độ phân giải và tỷ lệ) trước khi hiện, giống khi người dùng kéo cửa sổ sang đó
                    Rectangle wa2 = Screen.AllScreens[screenIdx].WorkingArea;
                    f.StartPosition = FormStartPosition.Manual;
                    f.Location = new Point(wa2.Left + 40, wa2.Top + 40);
                }
                f.Show();
                for (int i = 0; i < 25; i++) { Application.DoEvents(); Thread.Sleep(60); }
                string before = DpiUtil.Describe(f);
                if (moveTo >= 0 && moveTo < Screen.AllScreens.Length)
                {
                    Rectangle wa3 = Screen.AllScreens[moveTo].WorkingArea;
                    f.Location = new Point(wa3.Left + 40, wa3.Top + 40);      // kéo sang màn hình khác như người dùng kéo cửa sổ
                    for (int i = 0; i < 40; i++) { Application.DoEvents(); Thread.Sleep(60); }
                }
                var tc = FindTab(f);
                if (tc != null && tab > 0 && tab < tc.TabPages.Count) { tc.SelectedIndex = tab; for (int i = 0; i < 15; i++) { Application.DoEvents(); Thread.Sleep(60); } }
                int aw = GetAwarenessFromDpiAwarenessContext(GetWindowDpiAwarenessContext(h));
                RECT r; GetWindowRect(h, out r);
                report = "mode=" + mode + " dpi=" + GetDpiForWindow(h) + " awareness=" + aw + " (0 không, 1 hệ thống, 2 từng màn hình)"
                    + " rect=" + r.L + "," + r.T + "," + (r.R - r.L) + "x" + (r.B - r.T) + " font=" + f.Font.Name + " " + f.Font.SizeInPoints + "pt"
                    + " | trước=[" + before + "] | sau=[" + DpiUtil.Describe(f) + "]";
                ListView lvs = FindList(f);
                if (lvs != null)
                {
                    int sum = 0; foreach (ColumnHeader ch in lvs.Columns) sum += ch.Width;
                    report += " | danhSách: tổng cột " + sum + " px, bề rộng " + lvs.ClientSize.Width + " px " + (sum <= lvs.ClientSize.Width + 2 ? "(VỪA)" : "(TRÀN: có thanh cuộn ngang, cột cuối bị cắt)");
                }
                // Chụp bằng chính luồng này (cùng ngữ cảnh DPI) để lấy điểm ảnh thật trên màn hình
                int w = r.R - r.L, hgt = r.B - r.T;
                using (var bmp = new Bitmap(w, hgt, PixelFormat.Format32bppArgb))
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(r.L, r.T, 0, 0, new Size(w, hgt));
                    bmp.Save(outPng, ImageFormat.Png);
                }
            }
            catch (Exception ex) { report = "LỖI: " + ex; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join(60000);
        Console.WriteLine(report);
        Environment.Exit(0);   // thoát ngay, không chạy FormClosing/Deactivate (tránh ghi đè cài đặt thật của người dùng)
        return 0;
    }

    static ListView FindList(Control c)
    {
        foreach (Control x in c.Controls)
        {
            ListView lv = x as ListView;
            if (lv != null && lv.Visible) return lv;
            var r = FindList(x); if (r != null) return r;
        }
        return null;
    }

    static TabControl FindTab(Control c)
    {
        foreach (Control x in c.Controls) { if (x is TabControl) return (TabControl)x; var r = FindTab(x); if (r != null) return r; }
        return null;
    }
}
