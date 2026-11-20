// Rev0.9.5.2: kiểm thử co giãn theo DPI từng màn hình (chữ Bảng điều khiển và hộp Cài đặt không được nhỏ, mờ, vỡ nét khi sang màn hình khác tỷ lệ).
// Mỗi kiểm tra gửi thông điệp WM_DPICHANGED giả vào form thật rồi đọc lại cỡ chữ, kích thước, vùng hiện.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using TroLyKyThuat;

static class TestDpi
{
    static int fails = 0, total = 0;
    static void Eq(object a, object b, string name)
    {
        total++;
        bool ok;
        if (a is double && b is double) ok = Math.Abs((double)a - (double)b) < 1e-9;
        else ok = Equals(a, b);
        if (!ok) fails++;
        Console.WriteLine((ok ? "PASS " : "FAIL ") + name + (ok ? "" : "  (nhận " + a + ", mong " + b + ")"));
    }
    static void Tr(bool cond, string name) { Eq(cond, true, name); }
    static void Near(double a, double b, double tol, string name)
    {
        total++;
        bool ok = Math.Abs(a - b) <= tol;
        if (!ok) fails++;
        Console.WriteLine((ok ? "PASS " : "FAIL ") + name + (ok ? "" : "  (nhận " + a.ToString("0.###") + ", mong " + b.ToString("0.###") + " ± " + tol + ")"));
    }

    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L, T, R, B; }

    /// <summary>Gửi WM_DPICHANGED giả: dpi mới và vùng Windows gợi ý. Trả true nếu DpiUtil nhận xử lý.</summary>
    static bool Send(Form f, int dpi, Rectangle suggested)
    {
        IntPtr p = Marshal.AllocHGlobal(16);
        try
        {
            RECT r; r.L = suggested.Left; r.T = suggested.Top; r.R = suggested.Right; r.B = suggested.Bottom;
            Marshal.StructureToPtr(r, p, false);
            Message m = Message.Create(f.Handle, 0x02E0, (IntPtr)((dpi << 16) | dpi), p);
            return DpiUtil.HandleDpiMessage(f, ref m);
        }
        finally { Marshal.FreeHGlobal(p); }
    }

    static void Collect(Control parent, Dictionary<Control, float> sizes)
    {
        foreach (Control c in parent.Controls) { sizes[c] = c.Font.SizeInPoints; Collect(c, sizes); }
    }

    static void RunAllTests()
    {
        // 1. Form nhỏ dựng như Bảng điều khiển: font form 11pt, một nhãn có font riêng đậm, một nút cỡ cố định
        Form f;
        Label bold; Button btn; Panel pnl;
        using (DpiUtil.EnterPerMonitorV2())
        {
            f = new Form();
            DpiUtil.ConfigureForm(f, 11f);
            f.Width = 560; f.Height = 700;
            bold = new Label(); bold.Text = "Tiêu đề"; bold.Font = new Font(f.Font, FontStyle.Bold); bold.Left = 10; bold.Top = 10; f.Controls.Add(bold);
            pnl = new Panel(); pnl.Left = 10; pnl.Top = 60; pnl.Width = 300; pnl.Height = 200; f.Controls.Add(pnl);
            btn = new Button(); btn.Text = "Quét"; btn.Left = 10; btn.Top = 10; btn.Width = 100; btn.Height = 30; pnl.Controls.Add(btn);
            IntPtr h = f.Handle;
            Eq(DpiUtil.WindowAwareness(h), 2, "(a) form tạo trong ngữ cảnh Per-Monitor V2 có nhận biết DPI = 2 (sắc nét, không bị kéo giãn ảnh)");
        }

        int baseDpi = DpiUtil.LayoutDpi(f);
        Tr(baseDpi >= 96, "(a) DPI bố cục ban đầu >= 96 (" + baseDpi + ")");
        float baseFont = f.Font.SizeInPoints, baseBold = bold.Font.SizeInPoints;
        int baseBtnW = btn.Width, baseBtnH = btn.Height;
        Near(baseFont, 11.0, 0.01, "(a) font form ban đầu 11pt");

        // 2. Sang màn hình 150%: chữ, nút co đúng 1,5 lần, vùng hiện lấy theo gợi ý của Windows
        int dpi150 = baseDpi * 3 / 2;
        Rectangle sug = new Rectangle(100, 80, 840, 900);
        Eq(Send(f, dpi150, sug), true, "(b) WM_DPICHANGED được DpiUtil xử lý");
        Near(f.Font.SizeInPoints, baseFont * 1.5, 0.05, "(b) font form co 1,5 lần");
        Near(bold.Font.SizeInPoints, baseBold * 1.5, 0.05, "(b) nhãn có font riêng cũng co 1,5 lần (không bị bỏ sót)");
        Near(btn.Width, baseBtnW * 1.5, 2, "(b) nút rộng co 1,5 lần");
        Near(btn.Height, baseBtnH * 1.5, 2, "(b) nút cao co 1,5 lần");
        Eq(DpiUtil.LayoutDpi(f), dpi150, "(b) DPI bố cục được ghi nhận");
        Tr(f.Bounds.Width <= sug.Width + 1 && f.Bounds.Height <= sug.Height + 1, "(b) vùng hiện theo gợi ý của Windows (" + f.Bounds + ")");

        // 3. Gửi lại cùng DPI: không co giãn lần nữa
        float fontAfter = f.Font.SizeInPoints; int btnAfter = btn.Width;
        Send(f, dpi150, sug);
        Near(f.Font.SizeInPoints, fontAfter, 0.001, "(c) gửi lại cùng DPI: font không đổi");
        Eq(btn.Width, btnAfter, "(c) gửi lại cùng DPI: nút không đổi");

        // 4. Về lại màn hình ban đầu: trở về cỡ cũ
        Send(f, baseDpi, new Rectangle(100, 80, 560, 700));
        Near(f.Font.SizeInPoints, baseFont, 0.1, "(d) về DPI ban đầu: font trở về cỡ cũ");
        Near(bold.Font.SizeInPoints, baseBold, 0.1, "(d) về DPI ban đầu: nhãn riêng trở về cỡ cũ");
        Near(btn.Width, baseBtnW, 2, "(d) về DPI ban đầu: nút trở về cỡ cũ");

        // 5. DPI quá thấp bị bỏ qua, thông điệp khác không bị chặn
        float f0 = f.Font.SizeInPoints;
        Send(f, 48, new Rectangle(0, 0, 100, 100));
        Near(f.Font.SizeInPoints, f0, 0.001, "(e) DPI < 96 bị bỏ qua");
        Message other = Message.Create(f.Handle, 0x0000, IntPtr.Zero, IntPtr.Zero);
        Eq(DpiUtil.HandleDpiMessage(f, ref other), false, "(e) thông điệp khác WM_DPICHANGED không bị DpiUtil nuốt");

        // 6. Gợi ý lớn hơn màn hình: bị cắt về vùng làm việc của màn hình
        Send(f, dpi150, new Rectangle(0, 0, 50000, 50000));
        Rectangle wa = Screen.FromHandle(f.Handle).WorkingArea;
        Tr(f.Bounds.Width <= wa.Width && f.Bounds.Height <= wa.Height, "(f) cửa sổ không vượt quá vùng làm việc của màn hình (" + f.Bounds + " trong " + wa + ")");

        // 7. SyncDpi đưa DPI bố cục khớp với DPI thật của cửa sổ
        DpiUtil.SyncDpi(f);
        Eq(DpiUtil.LayoutDpi(f), DpiUtil.WindowDpi(f.Handle), "(g) SyncDpi: DPI bố cục bằng DPI thật của cửa sổ");
        float zoomBase = f.Font.SizeInPoints;
        DpiUtil.SetZoom(f, 70);
        Near(f.Font.SizeInPoints, zoomBase * 0.7, 0.05, "Zoom 70% thu nhỏ font");
        int physicalDpi = DpiUtil.LayoutDpi(f);
        Near(DpiUtil.LayoutScale(f), physicalDpi / 96.0 * 0.7, 0.001, "Zoom kết hợp DPI");
        Send(f, physicalDpi * 3 / 2, f.Bounds);
        Near(f.Font.SizeInPoints, zoomBase * 1.05, 0.05, "Đổi màn hình duy trì zoom 70%");
        Send(f, physicalDpi, f.Bounds);
        DpiUtil.SetZoom(f, 100);
        Near(f.Font.SizeInPoints, zoomBase, 0.05, "Trở về 100% khôi phục font");
        f.Dispose();

        // 8. Bảng điều khiển thật (PaneForm) và hộp Cài đặt thật: mọi điều khiển co đều 1,5 lần
        Form pane;
        using (DpiUtil.EnterPerMonitorV2()) { pane = new PaneForm(new Rec("app")); IntPtr h = pane.Handle; }
        CheckWholeForm(pane, "PaneForm");
        pane.Dispose();

        Form set;
        using (DpiUtil.EnterPerMonitorV2()) { set = new SettingsForm(); IntPtr h = set.Handle; }
        CheckWholeForm(set, "SettingsForm");
        set.Dispose();
    }

    static ListView FindListView(Control parent)
    {
        foreach (Control c in parent.Controls)
        {
            ListView lv = c as ListView;
            if (lv != null) return lv;
            ListView r = FindListView(c);
            if (r != null) return r;
        }
        return null;
    }

    static void CheckWholeForm(Form form, string name)
    {
        int b = DpiUtil.LayoutDpi(form);
        var before = new Dictionary<Control, float>();
        Collect(form, before);
        float formBefore = form.Font.SizeInPoints;
        // Vùng gợi ý đặt cùng vị trí hiện tại để không kéo cửa sổ sang màn hình khác (nếu kéo, Windows tự gửi thêm một WM_DPICHANGED thật theo màn hình đích)
        ListView lv0 = FindListView(form);
        int col0Before = lv0 != null && lv0.Columns.Count > 0 ? lv0.Columns[0].Width : 0;
        int ox = form.Left, oy = form.Top;
        Send(form, b * 3 / 2, new Rectangle(ox, oy, form.Width * 3 / 2, Math.Min(form.Height * 3 / 2, 600)));
        Near(form.Font.SizeInPoints, formBefore * 1.5, 0.05, "(h) " + name + ": font form co 1,5 lần");
        if (lv0 != null && col0Before > 0)
        {
            Near(lv0.Columns[0].Width, col0Before * 1.5, 3, "(h) " + name + ": cột cố định của danh sách (cột 'Mức') co 1,5 lần để không cắt chữ");
            // Tổng cột so với bề rộng danh sách chỉ kiểm được khi form đã hiện và bố cục đã chạy: xem bin\pane-shot.exe (dòng "danhSách=...")
        }
        int bad = 0, count = 0; string first = "";
        foreach (KeyValuePair<Control, float> kv in before)
        {
            count++;
            double ratio = kv.Key.Font.SizeInPoints / kv.Value;
            if (Math.Abs(ratio - 1.5) > 0.03) { bad++; if (first == "") first = kv.Key.GetType().Name + "[" + kv.Key.Text + "] " + ratio.ToString("0.00"); }
        }
        Eq(bad, 0, "(h) " + name + ": cả " + count + " điều khiển con đều co chữ 1,5 lần" + (bad > 0 ? " (sai: " + first + ")" : ""));
        Send(form, b, new Rectangle(form.Left, form.Top, 560, 600));
        Near(form.Font.SizeInPoints, formBefore, 0.1, "(h) " + name + ": về DPI cũ trả lại cỡ chữ");
        var minus = (Button)form.Controls.Find("ZoomOut", true)[0];
        var plus = (Button)form.Controls.Find("ZoomIn", true)[0];
        Tr(minus.Parent == plus.Parent && minus.Parent.Parent.Dock == DockStyle.Top, name + ": có đủ nút thu/phóng");
        var tab = form.Controls.Count > 0 ? form.Controls[0] as TabControl : null;
        if (tab != null) tab.SelectedIndex = 1;
        int selected = tab == null ? -1 : tab.SelectedIndex;
        Send(form, 192, form.Bounds);
        float font200 = form.Font.SizeInPoints;
        DpiUtil.SetZoom(form, 70);
        Near(form.Font.SizeInPoints, font200 * 0.7, 0.05, name + ": thu nhỏ ở DPI 200%");
        Tr(tab == null || tab.SelectedIndex == selected, name + ": giữ nguyên thẻ đang chọn");
        DpiUtil.SetZoom(form, 100);
        Near(form.Font.SizeInPoints, font200, 0.05, name + ": khôi phục 100% ở DPI 200%");

    }

    static void TestZoomButtons()
    {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".txt");
        try
        {
            using (var form = new Form())
            {
                DpiUtil.ConfigureForm(form, 11f);
                var input = new TextBox { Dock = DockStyle.Fill, Text = "Nội dung đang nhập" };
                form.Controls.Add(input);
                UiZoom.AttachWithFile(form, input, path, null);
                form.Show();
                Application.DoEvents();
                var minus = (Button)form.Controls.Find("ZoomOut", true)[0];
                var plus = (Button)form.Controls.Find("ZoomIn", true)[0];
                var label = form.Controls.Find("ZoomPercent", true)[0];
                Tr(input.Top >= minus.Parent.Bottom, "Thanh zoom ở trên vùng nội dung");
                float size = form.Font.SizeInPoints;
                for (int i = 0; i < 8; i++) minus.PerformClick();
                Eq(label.Text, "50%", "Năm lần thu nhỏ (bấm dư) dừng ở 50%");
                Tr(!minus.Enabled, "Nút thu nhỏ bị khóa ở 50%");
                Near(form.Font.SizeInPoints, size * 0.5, 0.05, "Thu nhỏ có tác dụng ngay");
                Eq(input.Text, "Nội dung đang nhập", "Zoom giữ nội dung đang nhập");
                for (int i = 0; i < 10; i++) plus.PerformClick();
                Eq(label.Text, "130%", "Phóng to (bấm dư) dừng ở 130%");
                Tr(!plus.Enabled, "Nút phóng to bị khóa ở 130%");
                Near(form.Font.SizeInPoints, size * 1.3, 0.05, "Phóng to có tác dụng ngay");
                Eq(File.ReadAllText(path), "130", "Lưu tỷ lệ mới");
            }
            using (var form = new Form())
            {
                var content = new Panel { Dock = DockStyle.Fill };
                form.Controls.Add(content);
                UiZoom.AttachWithFile(form, content, path, null);
                Eq(form.Controls.Find("ZoomPercent", true)[0].Text, "130%", "Mở lại đọc tỷ lệ đã lưu");
            }
        }
        finally { File.Delete(path); }
    }

    [STAThread]

    // Rev0.9.5.6: vệ sinh giao diện: không còn ký tự phím tắt (&) làm mất chữ, chữ GDI sắc nét, thẻ Công cụ AI, bố cục không chồng chữ
    static void CollectAll(Control parent, System.Collections.Generic.List<Control> list)
    {
        foreach (Control c in parent.Controls) { list.Add(c); CollectAll(c, list); }
    }

    static void TestUiHygiene()
    {
        Form pane;
        using (DpiUtil.EnterPerMonitorV2()) { pane = new PaneForm(new Rec("app")); IntPtr h = pane.Handle; pane.PerformLayout(); }
        var all = new System.Collections.Generic.List<Control>();
        CollectAll(pane, all);
        int badMn = 0, badTr = 0;
        foreach (Control c in all)
        {
            var lb = c as Label; var bb = c as ButtonBase;
            if (lb != null && (lb.UseMnemonic || lb.UseCompatibleTextRendering)) { if ((lb.Text ?? "").Contains("&") || lb.UseCompatibleTextRendering) { if ((lb.Text ?? "").Contains("&") && lb.UseMnemonic) badMn++; if (lb.UseCompatibleTextRendering) badTr++; } }
            if (bb != null && (bb.UseMnemonic || bb.UseCompatibleTextRendering)) { if ((bb.Text ?? "").Contains("&") && bb.UseMnemonic) badMn++; if (bb.UseCompatibleTextRendering) badTr++; }
        }
        Eq(badMn, 0, "Bảng điều khiển: không điều khiển nào chứa & mà còn bật phím tắt");
        Eq(badTr, 0, "Bảng điều khiển: mọi Label và nút dùng GDI TextRenderer (UseCompatibleTextRendering = false)");

        // thẻ Công cụ AI
        TabControl tc = null;
        foreach (Control c in all) { var t = c as TabControl; if (t != null && t.TabPages.Count == 3) tc = t; }
        Tr(tc != null, "tìm thấy TabControl chính của Bảng điều khiển");
        if (tc != null)
        {
            bool hasAi = false, hasOld = false;
            foreach (TabPage tp in tc.TabPages) { if (tp.Text == "Công cụ AI") hasAi = true; if (tp.Text == "Viết lại" || tp.Text == "Văn phong" || tp.Text == "Ngữ pháp" || tp.Text == "Thuật ngữ") hasOld = true; }
            Tr(hasAi, "thẻ đã đổi tên thành Công cụ AI");
            Tr(!hasOld, "không còn thẻ độc lập Viết lại, Văn phong, Ngữ pháp");
        }
        CheckBox cw = null, cd = null; ComboBox cl = null, tools = null; GroupBox gm = null;
        foreach (Control c in all)
        {
            var cb = c as CheckBox;
            if (cb != null && cb.Text == "Viết lại") cw = cb;
            if (cb != null && cb.Text == "Dịch") cd = cb;
            var g = c as GroupBox; if (g != null && g.Text == "Vị trí chèn") gm = g;
            var co = c as ComboBox; if (co != null && co.AccessibleName == "Ngôn ngữ đích") cl = co;
            if (co != null && co.AccessibleName == "Chọn công cụ AI") tools = co;
        }
        Tr(cw != null && cd != null, "có hai ô chọn Viết lại và Dịch");
        if (cw != null && cd != null)
        {
            Tr(cw.Checked && !cd.Checked, "mặc định chọn Viết lại, chưa chọn Dịch");
        }
        Tr(cl != null, "có danh sách ngôn ngữ đích");
        if (cl != null)
        {
            Eq(cl.Items.Count, 10, "danh sách ngôn ngữ đích có 10 mục");
            Eq(cl.Items[0].ToString().StartsWith("Tiếng Việt"), true, "ngôn ngữ đầu tiên là tiếng Việt");
            Eq(cl.Items[1].ToString().StartsWith("Tiếng Anh"), true, "ngôn ngữ thứ hai là tiếng Anh");
        }
        Tr(gm == null, "bỏ nhóm Vị trí chèn để chỉ thay vùng chọn");
        Tr(tools != null && tools.Items.Count == 5 &&
            (string)tools.Items[0] == "Thuật ngữ" && (string)tools.Items[1] == "Văn phong" &&
            (string)tools.Items[2] == "Ngữ pháp" && (string)tools.Items[3] == "Từ đồng nghĩa" &&
            (string)tools.Items[4] == "Viết lại và Dịch", "một thẻ Công cụ AI có đủ năm nhóm chức năng");

        // Use a deterministic missing-key state without touching the user's settings.
        var settingsField = typeof(PaneForm).GetField("st", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        var refresh = typeof(PaneForm).GetMethod("RefreshAiNotice", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        pane.Show();
        Application.DoEvents();
        settingsField.SetValue(pane, new AppSettings());
        if (tc != null && tools != null)
        {
            tc.SelectedTab = tc.TabPages[2];
            var notice = (Label)pane.Controls.Find("AiKeyNotice", true)[0];
            var routes = new[] { "term", "style", "gram", "syn", "ai" };
            for (int i = 0; i < routes.Length; i++)
            {
                tools.SelectedIndex = i;
                Tr(notice.Visible && notice.ForeColor == Color.DarkRed && notice.Text.Contains("API key") &&
                    pane.Controls.Find("AiPanel_" + routes[i], true)[0].Visible,
                    "công cụ " + routes[i] + " hiện hướng dẫn API key bằng chữ đỏ");
            }
            settingsField.SetValue(pane, new AppSettings { ApiKeyEnc = Secret.Protect("test-only") });
            refresh.Invoke(pane, null);
            Tr(!notice.Visible, "ẩn hướng dẫn khi khóa tồn tại");
            tc.SelectedIndex = 0;
            var scanBox = (TableLayoutPanel)pane.Controls.Find("ReviewOptions", true)[0];
            var plus = (Button)pane.Controls.Find("ZoomIn", true)[0];
            var minus = (Button)pane.Controls.Find("ZoomOut", true)[0];
            var zoomPercent = (Label)pane.Controls.Find("ZoomPercent", true)[0];
            int baseDpi = DpiUtil.LayoutDpi(pane);
            foreach (int dpi in new[] { baseDpi, 192 })
            {
                Send(pane, dpi, pane.Bounds);
                for (int zoom = 50; zoom <= 130; zoom += 10)
                {
                    DpiUtil.SetZoom(pane, zoom);
                    Application.DoEvents();
                    int availableRows = scanBox.GetRowHeights().Sum();
                    Tr(scanBox.Height >= availableRows &&
                        scanBox.Controls.Cast<Control>().All(c => c.Height >= c.GetPreferredSize(Size.Empty).Height &&
                            c.Bottom + c.Margin.Bottom <= scanBox.ClientSize.Height),
                        "tùy chọn Quét không bị cắt chữ ở DPI " + dpi + ", zoom " + zoom);
                    Tr(plus.Size == minus.Size && plus.Top == minus.Top && plus.Top == zoomPercent.Top &&
                        plus.Left > zoomPercent.Left && zoomPercent.Left > minus.Left,
                        "cụm thu phóng cân đối ở DPI " + dpi + ", zoom " + zoom);
                }
            }
            pane.Hide();
        }
        pane.Dispose();

        // Cài đặt: không còn & mất chữ, bố cục Dàn trang không chồng chữ
        Form set;
        using (DpiUtil.EnterPerMonitorV2()) { set = new SettingsForm(); IntPtr h = set.Handle; set.PerformLayout(); }
        var all2 = new System.Collections.Generic.List<Control>();
        CollectAll(set, all2);
        int bad2 = 0, badTr2 = 0, overlaps = 0, tables = 0;
        foreach (Control c in all2)
        {
            var lb = c as Label; var bb = c as ButtonBase;
            if (lb != null) { if ((lb.Text ?? "").Contains("&") && lb.UseMnemonic) bad2++; if (lb.UseCompatibleTextRendering) badTr2++; }
            if (bb != null) { if ((bb.Text ?? "").Contains("&") && bb.UseMnemonic) bad2++; if (bb.UseCompatibleTextRendering) badTr2++; }
            var tlp = c as TableLayoutPanel;
            if (tlp != null && tlp.Controls.Count >= 4)
            {
                tables++;
                for (int i = 0; i < tlp.Controls.Count; i++)
                    for (int j = i + 1; j < tlp.Controls.Count; j++)
                        if (tlp.Controls[i].Visible && tlp.Controls[j].Visible && tlp.Controls[i].Bounds.IntersectsWith(tlp.Controls[j].Bounds)) overlaps++;
            }
        }
        Eq(bad2, 0, "Cài đặt: không điều khiển nào chứa & mà còn bật phím tắt (nút Kiểm tra & Lưu hiện đủ chữ)");
        Eq(badTr2, 0, "Cài đặt: mọi Label và nút dùng GDI TextRenderer");
        Tr(tables >= 4, "Cài đặt có các hàng nhãn và ô nhập (Dàn trang)");
        Eq(overlaps, 0, "Cài đặt: các hàng nhãn và ô nhập không chồng chữ");
        bool hasKeyBtn = false;
        foreach (Control c in all2) { var b = c as Button; if (b != null && b.Text == "Kiểm tra & Lưu" && !b.UseMnemonic) hasKeyBtn = true; }
        Tr(hasKeyBtn, "nút Kiểm tra & Lưu tắt phím tắt nên hiện đủ dấu &");
        set.Dispose();
    }

    static int Main()
    {
        Console.SetOut(new StreamWriter(Console.OpenStandardOutput(), new System.Text.UTF8Encoding(false)) { AutoFlush = true });
        using (DpiUtil.EnterPerMonitorV2()) { RunAllTests(); TestZoomButtons(); TestUiHygiene(); }
        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}
