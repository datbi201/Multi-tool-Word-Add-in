// Kiểm thử lõi tính toán (không cần Word). Chạy bằng run-tests.bat.
using System;
using System.Collections.Generic;
using TroLyKyThuat;

static class T
{
    static int fails = 0, total = 0;
    static void Eq(object a, object b, string name)
    {
        total++;
        bool ok = Equals(a, b);
        if (!ok) fails++;
        Console.WriteLine((ok ? "PASS " : "FAIL ") + name + (ok ? "" : "  (nhận " + a + ", mong " + b + ")"));
    }
    static int Main()
    {
        ParsedNumber p;
        p = NumberUtil.Parse("1,34"); Eq(p.V, 1.34, "1,34 = 1.34"); Eq(p.Style, "vn", "1,34 là kiểu Việt"); Eq(p.Decimals, 2, "1,34 có 2 thập phân");
        p = NumberUtil.Parse("3.6"); Eq(p.V, 3.6, "3.6"); Eq(p.Style, "us", "3.6 kiểu Mỹ");
        p = NumberUtil.Parse("12,500.5"); Eq(p.V, 12500.5, "12,500.5"); Eq(p.Decimals, 1, "12,500.5 có 1 thập phân");
        p = NumberUtil.Parse("12.500,5"); Eq(p.V, 12500.5, "12.500,5"); Eq(p.Style, "vn", "12.500,5 kiểu Việt");
        p = NumberUtil.Parse("1,234"); Eq(p.Amb, true, "1,234 mơ hồ"); Eq(p.V, 1.234, "1,234 mặc định thập phân");
        p = NumberUtil.Parse("1.234.567"); Eq(p.V, 1234567.0, "1.234.567 là hàng nghìn"); Eq(p.Amb, false, "1.234.567 không mơ hồ");
        p = NumberUtil.Parse("−5,5"); Eq(p.V, -5.5, "dấu trừ Unicode");
        p = NumberUtil.Parse(" 12 500,5 "); Eq(p.V, 12500.5, "dấu cách nghìn kiểu SI");
        p = NumberUtil.Parse("0,125"); Eq(p.Amb, false, "0,125 không mơ hồ (bắt đầu 0)");
        Eq(NumberUtil.Parse("abc") == null, true, "chữ không phải số");
        Eq(NumberUtil.Parse("1 2") == null, true, "1 2 không phải số (dấu cách không đúng nhóm 3)");
        Eq(NumberUtil.Parse("5\u00A0478,80").V, 5478.8, "SI: 5 478,80 với dấu cách không ngắt");
        Eq(NumberUtil.Parse("12 mm") == null, true, "12 mm không phải số");
        Eq(NumberUtil.Parse("") == null, true, "rỗng không phải số");
        Eq(NumberUtil.Format(1234567.891, 3, "us", true), "1,234,567.891", "format us");
        Eq(NumberUtil.Format(1234567.891, 3, "vn", true), "1.234.567,891", "format vn");
        Eq(NumberUtil.Format(1234567.891, 3, "si", true), "1 234 567,891", "format si");
        Eq(NumberUtil.Format(1234.5, 1, "si", true), "1\u00A0234,5", "SI: từ 4 chữ số cũng tách nghìn (nhất quán với us, vn)");
        Eq(NumberUtil.Format(999.5, 1, "si", true), "999,5", "SI: dưới 1000 không tách");
        Eq(NumberUtil.Format(999.5, 1, "us", true), "999.5", "không thêm nghìn khi < 1000");
        Eq(NumberUtil.Format(-0.001, 2, "us", true), "0.00", "không để -0.00");
        Eq(NumberUtil.Format(-12.5, 0, "us", false), "-13", "làm tròn số âm");
        Eq(NumberUtil.Reformat("1,34", "us", "keep", true, false), "1.34", "đổi 1,34 sang kiểu Mỹ");
        Eq(NumberUtil.Reformat("12500.5", "us", "keep", true, false), "12,500.5", "thêm dấu nghìn");
        Eq(NumberUtil.Reformat("3.6", "vn", "keep", true, false), "3,6", "3.6 sang kiểu Việt");
        Eq(NumberUtil.Reformat("3.6", "us", "2", true, false), "3.60", "2 chữ số thập phân");
        Eq(NumberUtil.Reformat("62", "us", "2", true, false), "62", "số nguyên không thêm thập phân");
        Eq(NumberUtil.Reformat("62", "us", "2", true, true), "62.00", "số nguyên thêm thập phân khi bật");
        Eq(NumberUtil.Reformat("abc", "us", "2", true, true), "abc", "chữ giữ nguyên");
        var nums = new List<double> { 12.7, 11.9, 6.3 };
        Eq(Math.Round(NumberUtil.Aggregate("sum", nums), 6), 30.9, "tổng");
        Eq(Math.Round(NumberUtil.Aggregate("avg", nums), 6), 10.3, "trung bình");
        Eq(NumberUtil.Aggregate("min", nums), 6.3, "nhỏ nhất");
        Eq(NumberUtil.Aggregate("max", nums), 12.7, "lớn nhất");
        Eq(NumberUtil.Aggregate("cnt", nums), 3.0, "đếm");
        Eq(NumberUtil.Aggregate("prod", new List<double> { 2.5, 4, 3 }), 30.0, "tích 2,5×4×3=30");
        Eq(NumberUtil.Aggregate("prod", new List<double> { -2, 3.5 }), -7.0, "tích (-2)×3,5=-7");
        Eq(NumberUtil.Aggregate("prod", new List<double> { 1.5, 1.5 }), 2.25, "tích 1,5×1,5=2,25");
        Eq(NumberUtil.Aggregate("prod", new List<double> { 42 }), 42.0, "tích một số duy nhất");
        Eq(NumberUtil.AutoDecimals("sum", 1), 1, "tự động thập phân: tổng");
        Eq(NumberUtil.AutoDecimals("avg", 0), 2, "tự động thập phân: trung bình tối thiểu 2");
        Eq(NumberUtil.AutoDecimals("avg", 6), 4, "tự động thập phân: tối đa 4");
        Eq(NumberUtil.AutoDecimals("cnt", 3), 0, "tự động thập phân: đếm");
        Eq(NumberUtil.AutoDecimals("prod", 2), 2, "tự động thập phân: tích giữ nguyên");
        Eq(NumberUtil.AutoDecimals("prod", 8), 6, "tự động thập phân: tích tối đa 6");
        Eq(NumberUtil.OpLabel("prod"), "Tích", "nhãn phép nhân");
        // hướng tính
        Eq(CalcPlanner.ResolveDirection("auto", 2, 2, 2, 4), "row", "dãy ngang -> điền cột phải");
        Eq(CalcPlanner.ResolveDirection("auto", 2, 3, 5, 3), "col", "dãy dọc -> điền hàng dưới");
        Eq(CalcPlanner.ResolveDirection("auto", 2, 2, 4, 4), "row", "khối vuông: bằng nhau thì theo hàng");
        Eq(CalcPlanner.ResolveDirection("auto", 2, 2, 6, 3), "col", "khối nhiều hàng hơn cột: theo cột");
        Eq(CalcPlanner.ResolveDirection("auto", 2, 2, 3, 6), "row", "khối nhiều cột hơn hàng: theo hàng");
        Eq(CalcPlanner.ResolveDirection("row", 2, 3, 5, 3), "col", "cài đặt cũ bị bỏ qua, một cột luôn theo cột");
        var jr = CalcPlanner.Plan("row", 2, 2, 3, 4);
        Eq(jr.Count, 2, "kế hoạch hàng: 2 việc"); Eq(jr[0].Sources.Count, 3, "mỗi việc 3 ô nguồn");
        Eq(jr[1].Target[0] + "," + jr[1].Target[1], "3,5", "đích hàng 3 cột 5");
        var jc = CalcPlanner.Plan("col", 2, 2, 4, 3);
        Eq(jc.Count, 2, "kế hoạch cột: 2 việc"); Eq(jc[1].Target[0] + "," + jc[1].Target[1], "5,3", "đích hàng 5 cột 3");
        // Word map
        Eq(WordMap.ColorToWord(0x15, 0x60, 0x82), 0x15 + 0x60 * 256 + 0x82 * 65536, "màu BGR");
        Eq(AppSettings.HexToWord("FF0000"), 255, "đỏ = 255");
        Eq(AppSettings.HexToWord("0000FF"), 16711680, "xanh = 16711680");
        Eq(WordMap.LineWidthEnum(1.0), 8, "1pt -> 8");
        Eq(WordMap.LineWidthEnum(0.25), 2, "0.25pt -> 2");
        Eq(WordMap.LineWidthEnum(1.25), 8, "1.25pt làm tròn về 1pt (gần nhất)");
        Eq(WordMap.LineWidthEnum(2.0), 18, "2pt -> 2.25pt");
        Eq(WordMap.LineWidthEnum(5), 36, "5pt -> 4.5pt");
        Eq(WordMap.LineStyleEnum(0), 1, "viền liền"); Eq(WordMap.LineStyleEnum(2), 2, "viền chấm");
        // CSV
        Eq(CsvUtil.Escape("a,b"), "\"a,b\"", "CSV có dấu phẩy");
        Eq(CsvUtil.Escape("say \"hi\""), "\"say \"\"hi\"\"\"", "CSV có nháy kép");
        Eq(CsvUtil.CellText("12.7\r\a"), "12.7", "bỏ ký tự cuối ô của Word");
        // cài đặt
        var s = new AppSettings(); s.HdrBg = "AABBCC"; s.BorderPt = 0.75; s.NumDec = "2"; s.Zebra = true;
        string tmp = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "trolytest"); Environment.SetEnvironmentVariable("APPDATA", tmp);
        s.Save(); var s2 = AppSettings.Load();
        Eq(s2.HdrBg, "AABBCC", "lưu/nạp màu"); Eq(s2.BorderPt, 0.75, "lưu/nạp độ dày"); Eq(s2.NumDec, "2", "lưu/nạp thập phân"); Eq(s2.Zebra, true, "lưu/nạp zebra");
        // === Rev0.5: thông số bảng đặt từ Ribbon
        {
            var q = new AppSettings(); string m;
            m = TableParams.Apply("hdrAlR", q); Eq(q.HdrAlign, 2, "căn tiêu đề phải"); Eq(m.Contains("Mọi bảng"), true, "nhắc bấm Mọi bảng để áp dụng");
            TableParams.Apply("bdW4", q); Eq(q.BorderPt, 1.5, "độ dày 1,5 pt"); TableParams.Apply("bdStDash", q); Eq(q.BorderStyle, 1, "viền đứt");
            q.Zebra = false; TableParams.Apply("zebraToggle", q); Eq(q.Zebra, true, "bật tô xen kẽ"); TableParams.Apply("zebraToggle", q); Eq(q.Zebra, false, "tắt tô xen kẽ");
            TableParams.Apply("txtAlJ", q); Eq(q.TextAlign, 3, "căn chữ hai bên"); TableParams.Apply("numAlC", q); Eq(q.NumAlign, 1, "căn số giữa");
            Eq(TableParams.Apply("khong-co", q), null, "mã lạ trả về null");
            TableParams.SetColor("hdrBg", q, "aabbcc"); Eq(q.HdrBg, "AABBCC", "đặt màu nền tiêu đề"); Eq(TableParams.GetColor("border", q), q.BorderColor, "đọc màu viền");
            TableParams.SetSize(q, "10,5"); Eq(q.FontSize, 10.5, "cỡ chữ 10,5 (dấu phẩy)"); Eq(TableParams.SetSize(q, "99").StartsWith("Đã"), false, "cỡ chữ 99 bị từ chối"); Eq(q.FontSize, 10.5, "cỡ chữ không đổi khi lỗi");
            TableParams.SetSpace(q, "3"); Eq(q.CellSpace, 3.0, "cách đoạn 3 pt"); TableParams.SetFont(q, " Arial "); Eq(q.FontName, "Arial", "đặt font");
            Eq(new AppSettings().RepeatHeader && new AppSettings().NoRowBreak, true, "mặc định lặp tiêu đề và không ngắt hàng");
            // Rev0.6: bước tăng giảm, độ dày viền gõ tay, bảng màu gợi ý
            q.FontSize = 11; TableParams.Step("btnSzInc", q); Eq(q.FontSize, 11.5, "mũi tên lên: cỡ chữ +0,5"); TableParams.Step("btnSzDec", q); TableParams.Step("btnSzDec", q); Eq(q.FontSize, 10.5, "mũi tên xuống: cỡ chữ -0,5");
            q.FontSize = 72; TableParams.Step("btnSzInc", q); Eq(q.FontSize, 72.0, "cỡ chữ không vượt 72");
            q.CellSpace = 0; TableParams.Step("btnSpDec", q); Eq(q.CellSpace, 0.0, "cách đoạn không âm"); TableParams.Step("btnSpInc", q); Eq(q.CellSpace, 1.0, "cách đoạn +1");
            q.BorderPt = 1; TableParams.Step("btnBdInc", q); Eq(q.BorderPt, 1.5, "viền dày hơn: bậc kế tiếp 1,5"); TableParams.Step("btnBdDec", q); TableParams.Step("btnBdDec", q); Eq(q.BorderPt, 0.75, "viền mỏng hơn hai bậc: 0,75");
            q.BorderPt = 3; TableParams.Step("btnBdInc", q); Eq(q.BorderPt, 3.0, "viền không dày quá 3 pt");
            string mb = TableParams.SetBorder(q, "0,6"); Eq(q.BorderPt, 0.5, "gõ 0,6: làm tròn về 0,5 (bậc gần nhất của Word)"); Eq(mb.Contains("làm tròn"), true, "báo đã làm tròn");
            Eq(TableParams.SetBorder(q, "abc").StartsWith("Đã"), false, "độ dày chữ bị từ chối");
            // Rev0.8: chuỗi nhóm nghìn sai quy tắc không phải số (lỗi đổi giá trị 48,5 thành 4.850.000)
            foreach (string bad in new[] { "48,5,000", "48,5,0", "1.2,3", "12,34,567", "1,5.000", "1.2.3", "4,850,00", "1234,567,890", "5,", ",5" })
                Eq(NumberUtil.Parse(bad), null, "không phải số: " + bad);
            foreach (string ok in new[] { "48,5", "1.234,56", "1,234.56", "1.234.567", "1,234,567", "12,345.6", "0,500", "4850000" })
                Eq(NumberUtil.Parse(ok) != null, true, "là số hợp lệ: " + ok);
            Eq(NumberUtil.Parse("1.234.567,89").D, 1234567.89m, "1.234.567,89");
            Eq(NumberUtil.Parse("1,234,567").D, 1234567m, "1,234,567");
            // Rev0.7: bảng màu giống Word
            Eq(WordPalette.Tint("4472C4", 0.2, 0.8), "DAE3F3", "Office 2013 Nhấn 1 sáng hơn 80%");
            Eq(WordPalette.Tint("4472C4", 0.4, 0.6), "B4C7E7", "Office 2013 Nhấn 1 sáng hơn 60%");
            Eq(WordPalette.Tint("4472C4", 0.6, 0.4), "8FAADC", "Office 2013 Nhấn 1 sáng hơn 40%");
            Eq(WordPalette.Tint("4472C4", 0.75, 0), "2F5597", "Office 2013 Nhấn 1 tối hơn 25%");
            Eq(WordPalette.Tint("4472C4", 0.5, 0), "203864", "Nhấn 1 tối hơn 50% (công thức HSL; Office cho 1F3864, lệch 1 đơn vị ở một kênh do cách làm tròn nội bộ)");
            Eq(WordPalette.Tint("ED7D31", 0.2, 0.8), "FBE5D6", "Office 2013 Nhấn 2 sáng hơn 80%");
            Eq(WordPalette.Tint("ED7D31", 0.75, 0), "C55A11", "Office 2013 Nhấn 2 tối hơn 25%");
            Eq(WordPalette.Tint("A5A5A5", 0.2, 0.8), "EDEDED", "Office 2013 Nhấn 3 sáng hơn 80%");
            Eq(WordPalette.Tint("FFC000", 0.5, 0), "7F6000", "Office 2013 Nhấn 4 tối hơn 50%");
            Eq(WordPalette.ThemeCount, 60, "Màu giao diện: 10 cột x 6 hàng");
            Eq(WordPalette.ThemeHex(0) + "," + WordPalette.ThemeHex(1) + "," + WordPalette.ThemeHex(4) + "," + WordPalette.ThemeHex(9), "FFFFFF,000000,156082,4EA72E", "hàng 1 là các màu gốc của theme mặc định");
            Eq(WordPalette.ThemeHex(10), "F2F2F2", "Nền 1 tối hơn 5%");
            Eq(WordPalette.ThemeHex(50), "7F7F7F", "Nền 1 tối hơn 50%");
            Eq(WordPalette.ThemeHex(11), "7F7F7F", "Chữ 1 sáng hơn 50%");
            Eq(WordPalette.ThemeHex(51), "0D0D0D", "Chữ 1 sáng hơn 5%");
            Eq(WordPalette.ThemeHex(60), null, "ô ngoài bảng");
            Eq(WordPalette.StdHex(0) + "," + WordPalette.StdHex(9), "C00000,7030A0", "Màu chuẩn: đỏ đậm đến tím");
            Eq(WordPalette.StdCount, 10, "Màu chuẩn có 10 màu");
            string tg; Eq(WordPalette.Resolve("gT_hdrBg", 4, out tg) + ":" + tg, "156082:hdrBg", "chọn Nhấn 1 cho màu nền tiêu đề");
            Eq(WordPalette.Resolve("gS_zebra", 3, out tg) + ":" + tg, "FFFF00:zebra", "chọn Vàng chuẩn cho dòng chẵn");
            Eq(WordPalette.Resolve("gS_zebra", 99, out tg), null, "ô vượt giới hạn");
            Eq(WordPalette.AutoHex("hdrFg") + WordPalette.AutoHex("hdrBg"), "000000FFFFFF", "Tự động: chữ đen, nền trắng");
            string mx = WordPalette.MenuXml("border", "clrBorder");
            Eq(mx.Split(new[] { "<gallery" }, StringSplitOptions.None).Length - 1, 2, "menu màu dự phòng có hai bảng chọn");
            Eq(mx.Contains("columns=\"10\" rows=\"6\"") && mx.Contains("id=\"clrBorder\"") && mx.Contains("Thêm màu..."), true, "bảng 10 x 6 và nút Thêm màu...");
            string bx = WordPalette.ButtonXml("btnHdrBg", "Màu nền", "Mở bảng màu");
            Eq(bx.Contains("<button") && bx.Contains("onAction=\"OnBtn\""), true, "F: nút màu thống nhất một bước bấm");
            Eq(WordPalette.ThemeSectionLabel + "|" + WordPalette.StdSectionLabel + "|" + WordPalette.MoreColorsLabel, "Màu giao diện|Màu chuẩn|Thêm màu...", "F: nhãn tiếng Việt chuẩn Word");
            { string[] t = (string[])WordPalette.DefaultTheme.Clone(); t[4] = "112233"; WordPalette.SetTheme(t); Eq(WordPalette.ThemeHex(4), "112233", "bộ màu theo tài liệu"); WordPalette.ResetTheme(); }
            { var z = new AppSettings(); Eq(z.SciDec, 2, "mặc định kiểu kỹ thuật 2 số lẻ"); Eq(z.ShowDialog, false, "mặc định không hiện hộp thoại"); Eq(z.SyncSci, false, "mặc định đồng bộ thường"); }
            Eq(new AppSettings().SyncDec, 2, "mặc định đồng bộ 2 chữ số thập phân");
            // Rev0.9.3 (A, B, C): danh sách font, cỡ chữ chuẩn Word, độ dày viền 0.25..6 bước 0.25, số lượng 0..3
            Eq(string.Join(",", TableParams.FontPresets), "Times New Roman,Arial,Calibri,Segoe UI", "B: 4 font chuẩn");
            Eq(string.Join(",", TableParams.SizePresets), "8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72", "B: 16 cỡ chữ chuẩn Word");
            Eq(TableParams.SetSize(q, "36").StartsWith("Đã"), true, "B: chấp nhận cỡ chữ 36 pt");
            Eq(TableParams.SetSize(q, "72").StartsWith("Đã"), true, "B: chấp nhận cỡ chữ 72 pt");
            Eq(TableParams.SetSize(q, "73").StartsWith("Đã"), false, "B: từ chối cỡ chữ > 72 pt");
            Eq(string.Join(",", TableParams.BorderPresets), "0.25,0.5,0.75,1.0,1.5,2.0,3.0", "A: danh sách độ dày viền có sẵn");
            Eq(TableParams.SetBorder(q, "2.0").StartsWith("Đã") && q.BorderPt == 2.0, true, "A: gõ 2.0 giữ đúng 2.0");
            Eq(TableParams.SetBorder(q, "4.75").StartsWith("Đã") && q.BorderPt == 4.75, true, "A: nhập tay 4.75 (bước 0.25)");
            Eq(TableParams.SetBorder(q, "6.0").StartsWith("Đã") && q.BorderPt == 6.0, true, "A: tối đa 6.0 pt");
            Eq(TableParams.SetBorder(q, "6.25").StartsWith("Đã"), false, "A: từ chối > 6.0 pt");
            Eq(string.Join(",", TableParams.NPresets), "0,1,2,3", "C: danh sách số lượng 0..3");
            { var s0 = new AppSettings(); s0.AddN = 0; s0.Save(); Eq(AppSettings.Load().AddN, 0, "C: lưu và nạp AddN = 0"); }
            // Rev0.9.3 (F Dàn trang): mặc định và lưu/nạp cấu hình Dàn trang
            {
                var sd = new AppSettings();
                Eq(sd.PagePaper, "A4", "F Dàn trang: mặc định A4");
                Eq(sd.PgPortTop + "," + sd.PgPortBot + "," + sd.PgPortLeft + "," + sd.PgPortRight, "2,2,3,2", "F Dàn trang: lề dọc 2-2-3-2");
                Eq(sd.PgLandTop + "," + sd.PgLandBot + "," + sd.PgLandLeft + "," + sd.PgLandRight, "3,2,2,2", "F Dàn trang: lề ngang 3-2-2-2");
                Eq(sd.PgHeaderDist + "," + sd.PgFooterDist, "1,1", "F Dàn trang: header/footer 1 cm");
                Eq(sd.ParaSizeTnr + "," + sd.ParaSizeArial + "," + sd.ParaLineSpace + "," + sd.ParaSpaceBefore + "," + sd.ParaSpaceAfter, "13,11,1.2,6,6", "F Dàn trang: thông số đoạn văn mặc định");
                sd.PagePaper = "Letter"; sd.PgPortLeft = 3.5; sd.ParaSizeTnr = 14; sd.ParaLineSpace = 1.5; sd.Save();
                var sd2 = AppSettings.Load();
                Eq(sd2.PagePaper, "Letter", "F Dàn trang: lưu/nạp khổ giấy");
                Eq(sd2.PgPortLeft, 3.5, "F Dàn trang: lưu/nạp lề trái");
                Eq(sd2.ParaSizeTnr, 14.0, "F Dàn trang: lưu/nạp cỡ chữ TNR");
                Eq(sd2.ParaLineSpace, 1.5, "F Dàn trang: lưu/nạp giãn dòng");
            }
        }
        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails == 0 ? 0 : 1;
    }
}
