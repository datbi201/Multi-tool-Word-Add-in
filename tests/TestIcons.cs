// Rev0.9.5.1: bộ kiểm thử icon Ribbon (phương án A) và tiêu đề phiên bản.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Xml;
using TroLyKyThuat;

static class TestIcons
{
    static int fails = 0, total = 0;
    static void Eq(object a, object b, string name)
    {
        total++;
        bool ok = false;
        if (a != null && b != null && (a is byte || a is int || a is short || a is long) && (b is byte || b is int || b is short || b is long))
            ok = Convert.ToInt64(a) == Convert.ToInt64(b);
        else
            ok = Equals(a, b);
        if (!ok) fails++;
        Console.WriteLine((ok ? "PASS " : "FAIL ") + name + (ok ? "" : "  (nhận " + a + ", mong " + b + ")"));
    }
    static void Tr(bool cond, string name) { Eq(cond, true, name); }

    static bool BitmapsEqual(Bitmap b1, Bitmap b2)
    {
        if (b1.Width != b2.Width || b1.Height != b2.Height) return false;
        for (int y = 0; y < b1.Height; y++)
        {
            for (int x = 0; x < b1.Width; x++)
            {
                if (b1.GetPixel(x, y).ToArgb() != b2.GetPixel(x, y).ToArgb()) return false;
            }
        }
        return true;
    }

    static void RunAllTests()
    {
        // 1. Phân tích XML Ribbon từ Connect để lấy toàn bộ id có getImage="GetImg"
        var doc = new XmlDocument();
        doc.LoadXml(Connect.BuildRibbonXml());
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("r", "http://schemas.microsoft.com/office/2006/01/customui");

        var nodes = doc.SelectNodes("//*[@getImage='GetImg']", ns);
        if (nodes == null || nodes.Count == 0) nodes = doc.SelectNodes("//*[@getImage='GetImg']");

        var ribbonIds = new List<string>();
        foreach (XmlNode n in nodes)
        {
            if (n.Attributes["id"] != null)
            {
                string id = n.Attributes["id"].Value;
                if (!string.IsNullOrEmpty(id) && !id.StartsWith("auto_") && !ribbonIds.Contains(id))
                    ribbonIds.Add(id);
            }
        }

        Tr(ribbonIds.Count >= 50, "Tìm thấy " + ribbonIds.Count + " mã điều khiển có getImage='GetImg' trong Ribbon");

        // (a) Mọi mã nút có getImage="GetImg" đều có case riêng trong Icons.cs
        string iconsSrc = File.ReadAllText(@"src\Icons.cs");
        foreach (string id in ribbonIds)
        {
            bool hasCase = iconsSrc.Contains("case \"" + id + "\":");
            Tr(hasCase, "(a) Icons.cs có case riêng cho '" + id + "'");
        }

        // Danh sách kiểm tra diện rộng gồm ribbonIds và các mã bảng/menu/màu
        var testIds = new List<string>(ribbonIds);
        string[] extra = { "mnuBorder", "mnuZebra", "mnuCalc", "mnuNum", "btnCopy", "btnCsv", "btnNInc", "btnNDec", "sw", "gs", "swA" };
        foreach (string x in extra) if (!testIds.Contains(x)) testIds.Add(x);

        // (b) & (c): Kích thước 32x32, diện tích không trong suốt trong [4%, 60%], 4 góc trong suốt
        foreach (string id in testIds)
        {
            using (Bitmap bmp = Icons.Draw(id))
            {
                Tr(bmp != null, "Icons.Draw('" + id + "') không null");
                Eq(bmp.Width, 32, id + " rộng 32 px");
                Eq(bmp.Height, 32, id + " cao 32 px");

                int nonTrans = 0;
                for (int y = 0; y < 32; y++)
                {
                    for (int x = 0; x < 32; x++)
                    {
                        if (bmp.GetPixel(x, y).A > 0) nonTrans++;
                    }
                }
                double pct = (double)nonTrans / (32 * 32) * 100.0;
                Tr(pct >= 4.0 && pct <= 60.0, "(b) " + id + " diện tích không trong suốt " + pct.ToString("0.0") + "% trong [4%, 60%]");

                bool cornersTrans = (bmp.GetPixel(0, 0).A == 0) &&
                                    (bmp.GetPixel(31, 0).A == 0) &&
                                    (bmp.GetPixel(0, 31).A == 0) &&
                                    (bmp.GetPixel(31, 31).A == 0);
                Tr(cornersTrans, "(c) " + id + " 4 góc đều trong suốt (alpha = 0)");
            }
        }

        // (d) Tỉ lệ điểm ảnh gần màu xanh đặc RGB(0,120,212) và alpha 255 dưới 30% diện tích
        foreach (string id in ribbonIds)
        {
            using (Bitmap bmp = Icons.Draw(id))
            {
                int solidBlue = 0;
                for (int y = 0; y < 32; y++)
                {
                    for (int x = 0; x < 32; x++)
                    {
                        Color c = bmp.GetPixel(x, y);
                        if (c.A == 255 && Math.Abs(c.R - 0) <= 30 && Math.Abs(c.G - 120) <= 30 && Math.Abs(c.B - 212) <= 30)
                            solidBlue++;
                    }
                }
                double pct = (double)solidBlue / (32 * 32) * 100.0;
                Tr(pct < 30.0, "(d) " + id + " tỷ lệ xanh đặc " + pct.ToString("0.0") + "% < 30%");
            }
        }

        // (e) Hai mã khác nhau mà không nằm trong bảng bí danh thì ảnh phải khác nhau
        var aliases = new Dictionary<string, bool>();
        Action<string, string> addAlias = delegate(string a, string b)
        {
            aliases[a + "|" + b] = true;
            aliases[b + "|" + a] = true;
        };
        addAlias("btnRowBelow", "mnuAddRow");
        addAlias("btnColRight", "mnuAddCol");
        addAlias("btnDelRow", "mnuDel");
        addAlias("btnSum", "mnuCalc");
        addAlias("btnNumStyle", "mnuNum");
        addAlias("btnHdrFg", "swA");
        // Rev0.9.5.14: nút trong thẻ Công cụ dùng cùng hình với nút gốc ở thẻ Chuẩn hóa
        addAlias("btnParaFormat", "tbParaFormat");
        addAlias("btnPageSetup", "tbPageSetup");
        addAlias("btnAddPort", "tbAddPort");
        addAlias("btnAddLand", "tbAddLand");

        int dupCount = 0;
        for (int i = 0; i < testIds.Count; i++)
        {
            using (Bitmap bi = Icons.Draw(testIds[i]))
            {
                for (int j = i + 1; j < testIds.Count; j++)
                {
                    using (Bitmap bj = Icons.Draw(testIds[j]))
                    {
                        if (BitmapsEqual(bi, bj))
                        {
                            string key = testIds[i] + "|" + testIds[j];
                            if (!aliases.ContainsKey(key))
                            {
                                dupCount++;
                                Console.WriteLine("  FAIL: Trùng ảnh không mong muốn: " + testIds[i] + " == " + testIds[j]);
                            }
                        }
                    }
                }
            }
        }
        Tr(dupCount == 0, "(e) Các icon không nằm trong bảng bí danh đều có hình vẽ khác nhau");

        // (e2) Rev0.9.5.1 (audit của Claude): "khác nhau" phải là khác rõ, không phải lệch 1 px. Hai icon không phải bí danh phải khác nhau ít nhất 25 điểm ảnh.
        int nearDup = 0;
        for (int i = 0; i < testIds.Count; i++)
        {
            using (Bitmap bi = Icons.Draw(testIds[i]))
            {
                for (int j = i + 1; j < testIds.Count; j++)
                {
                    if (aliases.ContainsKey(testIds[i] + "|" + testIds[j])) continue;
                    using (Bitmap bj = Icons.Draw(testIds[j]))
                    {
                        int diff = 0;
                        for (int y = 0; y < 32; y++)
                            for (int x = 0; x < 32; x++)
                                if (bi.GetPixel(x, y).ToArgb() != bj.GetPixel(x, y).ToArgb()) diff++;
                        if (diff < 25)
                        {
                            nearDup++;
                            Console.WriteLine("  FAIL: Hai icon gần như giống nhau (" + diff + " điểm ảnh khác): " + testIds[i] + " và " + testIds[j]);
                        }
                    }
                }
            }
        }
        Tr(nearDup == 0, "(e2) Mọi cặp icon không phải bí danh khác nhau từ 25 điểm ảnh trở lên");

        // (e3) Ba nút căn bảng trái, giữa, phải: bảng nhỏ màu xanh phải dạt trái, giữa, phải thật sự (trọng tâm ngang cách nhau từ 3 px)
        Func<string, double> blueCx = delegate(string id)
        {
            using (Bitmap bb = Icons.Draw(id))
            {
                double sx = 0; int n = 0;
                for (int y = 0; y < 32; y++)
                    for (int x = 0; x < 32; x++)
                    {
                        Color c = bb.GetPixel(x, y);
                        if (c.A >= 200 && c.R <= 40 && Math.Abs(c.G - 120) <= 40 && Math.Abs(c.B - 212) <= 40) { sx += x; n++; }
                    }
                return n == 0 ? -1 : sx / n;
            }
        };
        double cxL = blueCx("btnTblAlL"), cxC = blueCx("btnTblAlC"), cxR = blueCx("btnTblAlR");
        Tr(cxL > 0 && cxC - cxL >= 3.0 && cxR - cxC >= 3.0, "(e3) btnTblAlL/C/R: trọng tâm xanh " + cxL.ToString("0.0") + " < " + cxC.ToString("0.0") + " < " + cxR.ToString("0.0") + ", mỗi bước từ 3 px");

        // (f) Draw với mã lạ và id:kiểu lạ không ném ngoại lệ và trả icon dự phòng khác mọi icon có tên
        Bitmap bUnk1 = null, bUnk2 = null;
        try
        {
            bUnk1 = Icons.Draw("khong_co_ma_nay_xyz");
            bUnk2 = Icons.Draw("ma_la:kieu_la");
            Tr(bUnk1 != null, "(f) Draw mã lạ không ném ngoại lệ");
            Tr(bUnk2 != null, "(f) Draw id:kiểu lạ không ném ngoại lệ");
            Tr(BitmapsEqual(bUnk1, bUnk2), "(f) Hai mã lạ trả cùng icon dự phòng");

            bool diffNamed = true;
            foreach (string id in testIds)
            {
                using (Bitmap bn = Icons.Draw(id))
                {
                    if (BitmapsEqual(bUnk1, bn)) { diffNamed = false; break; }
                }
            }
            Tr(diffNamed, "(f) Icon dự phòng khác biệt với tất cả icon có tên");
        }
        finally
        {
            if (bUnk1 != null) bUnk1.Dispose();
            if (bUnk2 != null) bUnk2.Dispose();
        }

        // (g) sw:RRGGBB và gs:RRGGBB có đúng màu tại tâm ô
        using (Bitmap b = Icons.Draw("sw:0070C0"))
        {
            Color c = b.GetPixel(16, 16);
            Eq(c.R, 0, "(g) sw:0070C0 tâm R=0");
            Eq(c.G, 0x70, "(g) sw:0070C0 tâm G=0x70");
            Eq(c.B, 0xC0, "(g) sw:0070C0 tâm B=0xC0");
        }
        using (Bitmap b = Icons.Draw("gs:FF0000"))
        {
            Color c = b.GetPixel(16, 16);
            Eq(c.R, 255, "(g) gs:FF0000 tâm R=255");
            Eq(c.G, 0, "(g) gs:FF0000 tâm G=0");
            Eq(c.B, 0, "(g) gs:FF0000 tâm B=0");
        }
        using (Bitmap b = Icons.Draw("sw:123456"))
        {
            Color c = b.GetPixel(16, 16);
            Eq(c.R, 0x12, "(g) sw:123456 tâm R=0x12");
            Eq(c.G, 0x34, "(g) sw:123456 tâm G=0x34");
            Eq(c.B, 0x56, "(g) sw:123456 tâm B=0x56");
        }

        // (h) Tiêu đề phiên bản hiển thị đúng 0.9.5.2
        string curTitle = PaneForm.GetTitle();
        Version curV = typeof(PaneForm).Assembly.GetName().Version;
        string curVer = curV.Revision > 0 ? string.Format("{0}.{1}.{2}.{3}", curV.Major, curV.Minor, curV.Build, curV.Revision) : string.Format("{0}.{1}.{2}", curV.Major, curV.Minor, curV.Build);
        Eq(curTitle, "Trợ lý Kỹ thuật - Rev" + curVer, "(h) Tiêu đề PaneForm theo AssemblyVersion hiện hành (không ghi cứng số bản)");
        string rev0Title = PaneForm.GetTitle(new Version(0, 9, 5, 0));
        Eq(rev0Title, "Trợ lý Kỹ thuật - Rev0.9.5", "(h) GetTitle với revision=0 chỉ hiện ba số 'Trợ lý Kỹ thuật - Rev0.9.5'");
        string rev2Title = PaneForm.GetTitle(new Version(0, 9, 5, 2));
        Eq(rev2Title, "Trợ lý Kỹ thuật - Rev0.9.5.2", "(h) GetTitle với revision=2 hiện bốn số 'Trợ lý Kỹ thuật - Rev0.9.5.2'");
    }

    static int Main()
    {
        Console.SetOut(new StreamWriter(Console.OpenStandardOutput(), new System.Text.UTF8Encoding(false)) { AutoFlush = true });
        RunAllTests();
        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}
