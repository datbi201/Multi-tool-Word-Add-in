// Rev0.9.6: kiểm thử đổi ngôn ngữ giao diện Việt, Anh (lớp Lang, Ribbon chạy thật, hộp Cài đặt, bảng điều khiển, bảng dịch).
// Biên dịch cùng toàn bộ mã nguồn và MockWord.cs (bảng điều khiển cần một đối tượng Word giả).
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using TroLyKyThuat;

static class TestLang
{
    static int total, fails;
    static void Eq(object a, object b, string name)
    {
        total++; bool ok = Equals(a, b); if (!ok) fails++;
        Console.WriteLine((ok ? "PASS " : "FAIL ") + name + (ok ? "" : "  (nhận [" + a + "], mong [" + b + "])"));
    }
    static void Tr(bool c, string name) { Eq(c, true, name); }

    public class Ctl { public string Id; public Ctl(string id) { Id = id; } }

    // chuỗi chữ của một cây điều khiển: văn bản nhãn, nút, ô tích; chú thích; tên và mô tả truy cập; mục ComboBox
    static void Collect(Control c, ToolTip tip, List<string> sink)
    {
        if (c is Label || c is ButtonBase || c is GroupBox || c is TabPage || c is Form) sink.Add(c.Text);
        sink.Add(c.AccessibleName); sink.Add(c.AccessibleDescription);
        if (tip != null) sink.Add(tip.GetToolTip(c));
        TabPage tp = c as TabPage; if (tp != null) sink.Add(tp.ToolTipText);
        ComboBox cb = c as ComboBox; if (cb != null) foreach (object it in cb.Items) sink.Add(it.ToString());
        ListView lv = c as ListView; if (lv != null) foreach (ColumnHeader h in lv.Columns) sink.Add(h.Text);
        foreach (Control ch in c.Controls) Collect(ch, tip, sink);
    }
    static ToolTip TipOf(Form f)
    {
        FieldInfo fi = f.GetType().GetField("tip", BindingFlags.NonPublic | BindingFlags.Instance);
        return fi == null ? null : (ToolTip)fi.GetValue(f);
    }
    static List<string> Snapshot(Form f) { var l = new List<string>(); Collect(f, TipOf(f), l); return l; }

    [STAThread]
    static int Main()
    {
        string tmp = Path.Combine(Path.GetTempPath(), "trolytest_lang"); Environment.SetEnvironmentVariable("APPDATA", tmp);
        try { if (Directory.Exists(tmp)) Directory.Delete(tmp, true); } catch { }
        Application.EnableVisualStyles();

        // ---------------------------------------------------------------- A. mặc định và Cài đặt
        Eq(Lang.Current, "vi", "mặc định ngôn ngữ là tiếng Việt");
        Eq(new AppSettings().Language, "vi", "AppSettings.Language mặc định là vi (không đổi giao diện của người dùng cũ)");
        { var s = AppSettings.Load(); s.Language = "en"; s.Save(); Eq(AppSettings.Load().Language, "en", "Cài đặt lưu và nạp lại ngôn ngữ en"); s.Language = "vi"; s.Save(); Eq(AppSettings.Load().Language, "vi", "Cài đặt lưu và nạp lại ngôn ngữ vi"); }
        {
            Directory.CreateDirectory(AppSettings.Dir);
            File.WriteAllText(AppSettings.FilePath, "Language=fr\r\n", new UTF8Encoding(true));
            Eq(AppSettings.Load().Language, "vi", "giá trị ngôn ngữ lạ trong tệp thì dùng vi");
            File.Delete(AppSettings.FilePath);
        }
        Eq(Lang.Normalize("EN"), "vi", "Normalize chỉ nhận đúng en, vi");
        Eq(Lang.Normalize("en"), "en", "Normalize en");

        // ---------------------------------------------------------------- B. T, Msg, Doc
        Eq(Lang.T("Mọi bảng"), "Mọi bảng", "vi: T giữ nguyên tiếng Việt");
        Eq(Lang.Msg("Đã thêm 3 hàng phía trên mỗi hàng trong số: hàng 2. Tổng cộng 3 hàng mới."), "Đã thêm 3 hàng phía trên mỗi hàng trong số: hàng 2. Tổng cộng 3 hàng mới.", "vi: Msg giữ nguyên thông báo");
        int fired = 0; Action h = delegate { fired++; }; Lang.Changed += h;
        Tr(Lang.Set("en"), "Set en trả về true khi có đổi");
        Tr(!Lang.Set("en"), "Set en lần hai trả về false (không đổi)");
        Eq(fired, 1, "sự kiện Changed chỉ phát một lần khi đổi");
        Eq(Lang.IsEn, true, "IsEn sau khi đặt en");
        Eq(Lang.T("Mọi bảng"), "All tables", "en: T dịch nhãn Ribbon");
        Eq(Lang.T("Chuỗi chưa từng có trong bảng dịch"), "Chuỗi chưa từng có trong bảng dịch", "en: chuỗi lạ giữ nguyên tiếng Việt");
        Eq(Lang.T(""), "", "en: chuỗi rỗng"); Eq(Lang.T(null), null, "en: null");
        Eq(Lang.Msg("Đã thêm 3 hàng phía trên mỗi hàng trong số: hàng 2. Tổng cộng 3 hàng mới."),
            "Added 3 row(s) above each of the selected rows: 2. Total 3 new row(s).", "en: Msg dịch theo mảnh, giữ số ở nguyên vị trí");
        Eq(Lang.Msg("Đã chuẩn hóa 3 đoạn trong vùng chọn (bỏ qua 2 tiêu đề hoặc đoạn trong ô bảng) và căn giữa; giãn dòng 1,2, trước 6 pt, sau 6 pt."),
            "Standardized 3 paragraphs in the selection (skipped 2 headings or paragraphs in table cells) and centered; line spacing 1,2, before 6 pt, after 6 pt.", "en: thông báo Chuẩn hóa đoạn văn nhiều đoạn");
        Eq(Lang.Msg("Đã điền 4 kết quả (Tổng) vào hàng bên dưới. Bỏ qua 1 ô không phải số. (0,5 giây)"),
            "Filled 4 result(s) (Sum) into the row below. Skipped 1 non-number cells. (0,5 s)", "en: thông báo tính trong bảng kèm thời gian");
        Eq(Lang.Msg("Không có ô để điền kết quả (bảng có ô gộp, không thêm được hàng mới). Hãy chọn thêm một ô trống bên dưới rồi tính lại."),
            "There is no cell to put the result in (the table has merged cells and a new row could not be added). Select one more empty cell below and calculate again.", "en: thông báo bảng có ô gộp");
        Eq(Lang.Msg("Đã chuẩn hóa dàn trang cho 2 phần (1 dọc, 1 ngang): giấy A4, lề dọc 2-2-3-2 cm, lề ngang 3-2-2-2 cm."),
            "Standardized the page setup for 2 sections (1 portrait, 1 landscape): paper A4, portrait margins 2-2-3-2 cm, landscape margins 3-2-2-2 cm.", "en: thông báo dàn trang (có {n} trong mẫu)");
        Eq(Lang.Msg("Đã viết bằng chữ: một nghìn hai trăm kilôpascan"), "Written in words: một nghìn hai trăm kilôpascan", "en: mảnh ngắn (ô) không khớp giữa một từ khác (kilôpascan)");
        Eq(Lang.Msg("hàng 3 có 4 ô"), "row 3 có 4 cells", "en: mảnh ngắn vẫn khớp khi đứng thành từ riêng");
        Eq(Lang.Doc("QUICK\r\nBẢNG\r\nCả bốn nút giữ nguyên vùng bôi đen"), Lang.Doc("QUICK\r\nBẢNG\r\nCả bốn nút giữ nguyên vùng bôi đen"), "en: Doc ổn định");
        Tr(Lang.Doc("HƯỚNG DẪN NHANH\r\n\r\nBẢNG").StartsWith("QUICK GUIDE\r\n\r\nTABLES"), "en: Doc dịch từng dòng, giữ dòng trống");
        Lang.Set("vi");
        Eq(fired, 2, "Changed phát lần hai khi đổi về vi");
        Lang.Changed -= h;
        Eq(Lang.T("Mọi bảng"), "Mọi bảng", "đổi về vi: T trả lại tiếng Việt");

        // ---------------------------------------------------------------- C. bảng dịch: tính nhất quán
        {
            string[] ex = (string[])Type.GetType("TroLyKyThuat.LangData").GetField("Exact", BindingFlags.Public | BindingFlags.Static).GetValue(null);
            string[] fr = (string[])Type.GetType("TroLyKyThuat.LangData").GetField("Frag", BindingFlags.Public | BindingFlags.Static).GetValue(null);
            Tr(ex.Length >= 1500 && ex.Length % 2 == 0, "bảng Exact có cặp chẵn và đủ lớn (" + ex.Length / 2 + " cặp)");
            Tr(fr.Length >= 800 && fr.Length % 2 == 0, "bảng Frag có cặp chẵn và đủ lớn (" + fr.Length / 2 + " cặp)");
            int empty = 0, same = 0, badBrace = 0;
            for (int i = 0; i + 1 < ex.Length; i += 2)
            {
                if (string.IsNullOrWhiteSpace(ex[i + 1])) empty++;
                if (ex[i] == ex[i + 1]) same++;
                if (ex[i].Contains("{") != ex[i + 1].Contains("{") && !ex[i].Contains("\\")) badBrace++;
            }
            Eq(empty, 0, "mọi bản dịch đều không rỗng");
            Eq(same, 0, "không có bản dịch trùng nguyên văn tiếng Việt");
            Eq(badBrace, 0, "dấu { } của bản gốc và bản dịch khớp nhau");
            var keys = new HashSet<string>(); int dup = 0;
            for (int i = 0; i < ex.Length; i += 2) if (!keys.Add(ex[i])) dup++;
            Eq(dup, 0, "không có khóa tiếng Việt lặp trong bảng Exact");
        }

        // ---------------------------------------------------------------- D. Ribbon chạy thật
        {
            var sw = new XmlDocument(); sw.LoadXml(Connect.RibbonXmlForTest());
            var rt = new XmlDocument(); rt.LoadXml(Connect.RibbonXmlRuntimeForTest());
            int leftStatic = 0;
            foreach (XmlElement e in rt.SelectNodes("//*[@id]"))
                foreach (string a in new[] { "label", "screentip", "supertip" })
                    if (e.HasAttribute(a) && Lang.HasVietnamese(e.GetAttribute(a)) && e.LocalName != "item") { leftStatic++; Console.WriteLine("  còn chữ tĩnh: " + e.GetAttribute("id") + "|" + a); }
            Eq(leftStatic, 0, "Ribbon chạy thật: không còn nhãn, chú thích tĩnh tiếng Việt (đã chuyển thành lệnh gọi lại)");
            Eq(rt.SelectNodes("//*[@id]").Count, sw.SelectNodes("//*[@id]").Count, "Ribbon chạy thật: số phần tử có mã không đổi");
            int ids = 0, dupId = 0; var seen = new HashSet<string>();
            foreach (XmlElement e in rt.SelectNodes("//*[@id]")) { ids++; if (!seen.Add(e.GetAttribute("id"))) dupId++; }
            Eq(dupId, 0, "Ribbon chạy thật: mã duy nhất");
            // mọi tên lệnh gọi lại trong XML đều có hàm công khai tương ứng trên Connect
            var methods = new HashSet<string>(); foreach (MethodInfo m in typeof(Connect).GetMethods(BindingFlags.Public | BindingFlags.Instance)) methods.Add(m.Name);
            var missing = new List<string>();
            foreach (XmlElement e in rt.SelectNodes("//*")) foreach (XmlAttribute a in e.Attributes)
                if (a.Name.StartsWith("get") || a.Name.StartsWith("on")) if (a.Name != "onLoad" || true) if (!methods.Contains(a.Value)) { missing.Add(a.Name + "=" + a.Value); }
            Eq(string.Join(",", missing.ToArray()), "", "Ribbon chạy thật: mọi lệnh gọi lại có hàm trên Connect");

            // gọi lại như Ribbon: id điều khiển -> nhãn
            var conn = new Connect();
            Lang.Set("vi");
            Eq(conn.GetLbl(new Ctl("btnApplyAll")), "Mọi bảng", "vi: GetLbl trả về nhãn gốc");
            Eq(conn.GetScr(new Ctl("btnApplyAll")), "Áp dụng master style cho mọi bảng", "vi: GetScr trả về chú thích gốc");
            Lang.Set("en");
            Eq(conn.GetLbl(new Ctl("btnApplyAll")), "All tables", "en: GetLbl dịch nhãn");
            Eq(conn.GetScr(new Ctl("btnApplyAll")), "Apply the master style to all tables", "en: GetScr dịch chú thích");
            Tr(conn.GetSup(new Ctl("btnSum")).StartsWith("Select a horizontal range"), "en: GetSup dịch chú thích dài");
            Eq(conn.GetLbl(new Ctl("khong-co-ma")), "", "mã lạ trả về chuỗi rỗng, không lỗi");
            Eq(conn.GetLangLabel(null), "Language: EN", "nhãn nút ngôn ngữ ở en");
            Eq(conn.GetTabTableLabel(null), "My table", "nhãn thẻ My table giữ nguyên ở en");
            Eq(conn.GetLabel(null).StartsWith("Separator "), true, "en: nhãn nút Dấu dịch (Separator ...)");
            Eq(conn.GetSyncLabel(null).StartsWith("Sync: "), true, "en: nhãn nút Đồng bộ dịch (Sync: ...)");

            // phủ 100% nhãn, tên gọi, chú thích tĩnh của Ribbon: mọi chuỗi có chữ Việt đều có bản dịch khác bản gốc
            var untr = new List<string>(); int count = 0;
            foreach (XmlElement e in sw.SelectNodes("//*[@id]")) foreach (string a in new[] { "label", "screentip", "supertip" })
                if (e.HasAttribute(a) && Lang.HasVietnamese(e.GetAttribute(a))) { count++; string v = e.GetAttribute(a); if (Lang.T(v) == v) untr.Add(e.GetAttribute("id") + "|" + a + ": " + v); }
            foreach (string u in untr) Console.WriteLine("  chưa dịch: " + u);
            Eq(untr.Count, 0, "Ribbon: mọi nhãn, tên gọi, chú thích đều có bản dịch (" + count + " chuỗi)");
            // dịch rồi không còn chữ Việt có dấu trong Ribbon (trừ các từ được giữ nguyên có chủ ý)
            var viLeft = new List<string>();
            foreach (XmlElement e in sw.SelectNodes("//*[@id]")) foreach (string a in new[] { "label", "screentip", "supertip" })
                if (e.HasAttribute(a) && Lang.HasVietnamese(e.GetAttribute(a)) && Lang.HasVietnamese(Lang.T(e.GetAttribute(a)))) viLeft.Add(e.GetAttribute("id") + "|" + a);
            foreach (string u in viLeft) Console.WriteLine("  còn chữ Việt trong bản dịch: " + u);
            Tr(viLeft.Count <= 2, "Ribbon ở en gần như không còn chữ Việt có dấu (ví dụ Đọc số giữ chữ đọc bằng tiếng Việt): " + viLeft.Count);
            Lang.Set("vi");
        }

        // ---------------------------------------------------------------- E. hộp Cài đặt và bảng điều khiển
        foreach (string which in new[] { "settings", "pane" })
        {
            Lang.Set("vi");
            Form f; using (DpiUtil.EnterPerMonitorV2()) { if (which == "settings") f = new SettingsForm(); else f = new PaneForm(new Rec("app")); IntPtr hh = f.Handle; }
            List<string> vi = Snapshot(f);
            Tr(vi.Count > 100, which + ": thu được nhiều chuỗi chữ (" + vi.Count + ")");
            Lang.Set("en");
            Lang.Apply(f, TipOf(f));
            List<string> en = Snapshot(f);
            Eq(en.Count, vi.Count, which + ": số chuỗi không đổi sau khi dịch");
            int untranslated = 0, changed = 0; var examples = new List<string>();
            for (int i = 0; i < vi.Count && i < en.Count; i++)
            {
                if (string.IsNullOrEmpty(vi[i]) || !Lang.HasVietnamese(vi[i]) || vi[i] == "Tiếng Việt") continue;   // "Tiếng Việt" là tên ngôn ngữ giữ nguyên chữ gốc
                if (en[i] == vi[i]) { untranslated++; if (examples.Count < 12) examples.Add(vi[i].Length > 90 ? vi[i].Substring(0, 90) : vi[i]); } else changed++;
            }
            foreach (string x in examples) Console.WriteLine("  chưa dịch (" + which + "): " + x);
            Eq(untranslated, 0, which + ": mọi chữ có tiếng Việt trên giao diện đều đã được dịch (" + changed + " chuỗi dịch)");
            Lang.Set("vi");
            Lang.Apply(f, TipOf(f));
            List<string> back = Snapshot(f);
            int diff = 0; for (int i = 0; i < vi.Count && i < back.Count; i++) if (vi[i] != back[i]) { diff++; if (diff < 4) Console.WriteLine("  khác sau khi về vi: [" + vi[i] + "] -> [" + back[i] + "]"); }
            Eq(diff, 0, which + ": đổi en rồi về vi trả lại đúng từng chữ gốc");
            // đổi qua lại nhiều lần vẫn ổn định
            Lang.Set("en"); Lang.Apply(f, TipOf(f)); Lang.Apply(f, TipOf(f)); List<string> en2 = Snapshot(f);
            int d2 = 0; for (int i = 0; i < en.Count && i < en2.Count; i++) if (en[i] != en2[i]) d2++;
            Eq(d2, 0, which + ": áp dụng nhiều lần cho kết quả giống nhau");
            Lang.Set("vi"); f.Dispose();
        }

        // ---------------------------------------------------------------- F. chữ do người dùng gõ không bị dịch
        {
            Lang.Set("vi");
            Form f = new SettingsForm(); IntPtr hh = f.Handle;
            Lang.Set("en"); Lang.Apply(f, TipOf(f));
            TextBox key = null; foreach (Control c in f.Controls) { }
            var all = new List<Control>(); Action<Control> walk = null; walk = delegate(Control c) { all.Add(c); foreach (Control ch in c.Controls) walk(ch); }; walk(f);
            TextBox help = null; foreach (Control c in all) { TextBox t = c as TextBox; if (t != null && (t.Tag as string) == "doc") help = t; }
            Tr(help != null && help.Text.StartsWith("QUICK GUIDE"), "hộp Cài đặt: hướng dẫn được dịch từng dòng ở en");
            Tr(help != null && help.Text.Contains("EQUATION") == false && help.Text.Length > 3000, "hộp Cài đặt: hướng dẫn đầy đủ độ dài sau dịch");
            Lang.Set("vi"); Lang.Apply(f, TipOf(f));
            Tr(help != null && help.Text.StartsWith("HƯỚNG DẪN NHANH"), "hộp Cài đặt: về vi thì hướng dẫn trở lại tiếng Việt");
            // ô chọn ngôn ngữ trong Cài đặt: tên ngôn ngữ giữ nguyên chữ gốc, không bị dịch
            ComboBox cl = null; foreach (Control c in all) { ComboBox cb = c as ComboBox; if (cb != null && cb.Items.Count == 2 && cb.Items[0] is Item && ((Item)cb.Items[0]).Value == "vi") cl = cb; }
            Tr(cl != null, "hộp Cài đặt có ô chọn ngôn ngữ giao diện");
            Lang.Set("en"); Lang.Apply(f, TipOf(f));
            Eq(cl.Items[0].ToString() + "|" + cl.Items[1].ToString(), "Tiếng Việt|English", "ô chọn ngôn ngữ giữ nguyên tên gốc ở cả hai ngôn ngữ");
            Lang.Set("vi"); f.Dispose();
        }

        // ---------------------------------------------------------------- G. luồng lưu từ hộp Cài đặt đổi ngôn ngữ
        {
            Lang.Set("vi");
            var s = AppSettings.Load(); s.Language = "en"; s.Save();
            Lang.Set(AppSettings.Load().Language);
            Eq(Lang.Current, "en", "nạp Cài đặt có Language=en thì giao diện chuyển sang en");
            s.Language = "vi"; s.Save(); Lang.Set(AppSettings.Load().Language);
            Eq(Lang.Current, "vi", "nạp Cài đặt có Language=vi thì giao diện về vi");
        }

        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}
