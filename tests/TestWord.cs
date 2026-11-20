using System;
using System.Collections.Generic;
using TroLyKyThuat;
static class TW
{
    static int fails = 0, total = 0;
    static void Eq(object a, object b, string name) { total++; bool ok = Equals(a, b); if (!ok) fails++; Console.WriteLine((ok ? "PASS " : "FAIL ") + name + (ok ? "" : "  (nhan " + a + ", mong " + b + ")")); }
    static MTable Make()
    {
        return new MTable(new string[][] {
            new string[] { "Section", "Nominal", "Measured", "Loss" },
            new string[] { "S1", "12.7", "11.9", "6.3" },
            new string[] { "S2", "12.7", "11.2", "11.8" },
            new string[] { "S3", "12.7", "12.1", "4.7" } });
    }
    static MTable MakeBig()
    {
        var rows = new List<string[]>();
        rows.Add(new string[] { "Section", "Nominal", "Measured", "Loss" });
        for (int i = 1; i <= 7; i++) rows.Add(new string[] { "S" + i, "12.7", "11." + i, "6." + i });
        return new MTable(rows.ToArray());
    }
    static bool Hf(MTable x) { object o; return ((Rec)x.Rows.Item(1)).P.TryGetValue("HeadingFormat", out o) && o is bool && (bool)o; }
    static bool HfSet(MTable x) { object o; return ((Rec)x.Rows.Item(1)).P.TryGetValue("HeadingFormat", out o) && o is bool; }
    static string Row(MTable t, int r) { return string.Join("|", t.D[r - 1].ToArray()); }
    static int Main()
    {
        var st = new AppSettings(); st.CalcOp = "sum"; st.CalcDec = "auto";
        // 1. Tổng ngang -> cột mới
        var t = Make(); var app = new MApp(t, new MSel(t, 2, 2, 2, 4)); var w = new WordTables(app);
        var info = w.GetSelection(true); Eq(info.InTable, true, "nhận biết đang ở trong bảng"); Eq(info.R1 + "," + info.C1 + "," + info.R2 + "," + info.C2, "2,2,2,4", "vùng chọn 2,2 đến 2,4");
        Eq(info.NumericCells, 3, "3 ô có số");
        string msg = w.Calc(st);
        Eq(t.D[0].Count, 5, "thêm 1 cột bên phải"); Eq(t.D[0][4], "Tổng", "nhãn tiêu đề cột mới");
        Eq(t.D[1][4], "30.9", "tổng hàng 1 = 30.9"); Console.WriteLine("  -> " + msg);
        // 2. Tính lại sau khi sửa
        t.D[1][1] = "100"; w.Calc(st); Eq(t.D[1][4], "118.2", "tính lại sau khi sửa ô (100 + 11.9 + 6.3)");
        // 3. Trung bình dọc -> hàng mới
        t = Make(); app = new MApp(t, new MSel(t, 2, 3, 4, 3)); w = new WordTables(app); st.CalcOp = "avg";
        w.Calc(st);
        Eq(t.D.Count, 5, "thêm hàng dưới"); Eq(t.D[4][2], "11.73", "trung bình cột 3 = 11.73"); Eq(t.D[4][0], "Trung bình", "nhãn hàng mới ở cột 1");
        // 4. Chọn khối nhiều hàng, hướng ép
        t = Make(); app = new MApp(t, new MSel(t, 2, 2, 4, 4)); w = new WordTables(app); st.CalcOp = "sum";
        w.Calc(st); Eq(t.D[1][4], "30.9", "khối 3x3 tính theo hàng: hàng 2 = 30.9"); Eq(t.D[3][4], "29.5", "hàng 4 = 29.5");
        // 4b. Cột đứng 4 ô (lỗi người dùng báo ở Rev0.3): tổng điền xuống hàng dưới, thêm hàng nếu cần
        t = Make(); app = new MApp(t, new MSel(t, 2, 3, 4, 3)); w = new WordTables(app); st.CalcOp = "sum"; msg = w.Calc(st);
        Eq(t.D.Count, 5, "tổng theo cột thêm 1 hàng"); Eq(t.D[4][2], "35.2", "tổng cột 3 = 35.2"); Eq(msg.Contains("hàng bên dưới"), true, "báo điền xuống hàng dưới");
        t = Make(); app = new MApp(t, new MSel(t, 2, 2, 4, 3)); w = new WordTables(app); w.Calc(st);
        Eq(t.D.Count, 5, "khối 3 hàng 2 cột: nhiều hàng hơn cột nên tính theo cột"); Eq(t.D[4][1], "38.1", "tổng cột 2 = 38.1");
        // 5. Ô chữ lẫn trong dãy
        t = Make(); t.D[1][2] = "n/a"; app = new MApp(t, new MSel(t, 2, 2, 2, 4)); w = new WordTables(app); st.CalcOp = "sum";
        msg = w.Calc(st); Eq(t.D[1][4], "19.0", "bỏ qua ô chữ, cộng 12.7 + 6.3 = 19.0"); Eq(msg.Contains("Bỏ qua 1"), true, "báo bỏ qua ô không phải số");
        // 5b. Rev0.9.1A (thử nghiệm): Tổng theo định dạng chỉ cộng ô đậm
        {
            var boldFont = new Rec("font"); boldFont.P["Bold"] = -1;
            var stf = new AppSettings(); stf.CalcOp = "sum"; stf.CalcDec = "auto"; stf.CalcFmtKind = "bold";
            t = Make(); t.Cell(2, 2).Range.P["Font"] = boldFont; t.Cell(2, 4).Range.P["Font"] = boldFont;   // 12.7 và 6.3 đậm, 11.9 không đậm
            app = new MApp(t, new MSel(t, 2, 2, 2, 4)); w = new WordTables(app);
            stf.CalcFmtOn = true; msg = w.Calc(stf);
            Eq(t.D[1][4], "19.0", "tổng theo định dạng đậm: chỉ 12.7 + 6.3 = 19.0");
            Eq(msg.Contains("bỏ qua 1 ô khác"), true, "báo bỏ qua 1 ô không đậm");
            t = Make(); app = new MApp(t, new MSel(t, 2, 2, 2, 4)); w = new WordTables(app);
            msg = w.Calc(stf);
            Eq(msg.Contains("Không có ô nào khớp"), true, "không có ô đậm: báo rõ, không điền gì");
            Eq(t.D[0].Count, 4, "không có ô khớp thì không thêm cột");
            t = Make(); t.Cell(2, 2).Range.P["Font"] = boldFont; t.Cell(2, 4).Range.P["Font"] = boldFont;
            app = new MApp(t, new MSel(t, 2, 2, 2, 4)); w = new WordTables(app);
            stf.CalcFmtOn = false; msg = w.Calc(stf);
            Eq(t.D[1][4], "30.9", "tắt lọc: cộng cả ba ô như cũ (30.9)");
        }
        // 6. Thêm hàng/cột: mỗi hàng (cột) được chọn xử lý riêng
        t = Make(); app = new MApp(t, new MSel(t, 2, 2, 3, 3)); w = new WordTables(app);
        info = w.GetSelection(true); Eq(string.Join(",", info.Rows.ConvertAll(x => x.ToString()).ToArray()), "2,3", "danh sách hàng đã chọn 2,3");
        w.AddRows(2, true, ""); Eq(t.D.Count, 8, "chọn hàng 2-3, thêm 2 phía trên mỗi hàng: thêm 4 hàng"); Eq(Row(t, 2), "|||", "hàng 2 trống"); Eq(Row(t, 3), "|||", "hàng 3 trống");
        Eq(t.D[3][0], "S1", "S1 xuống hàng 4"); Eq(Row(t, 5), "|||", "hàng 5 trống (phía trên S2)"); Eq(Row(t, 6), "|||", "hàng 6 trống"); Eq(t.D[6][0], "S2", "S2 xuống hàng 7"); Eq(t.D[7][0], "S3", "S3 hàng 8");
        t = Make(); app = new MApp(t, new MSel(t, 2, 2, 3, 3)); w = new WordTables(app);
        w.AddRows(1, false, null); Eq(t.D.Count, 6, "thêm 1 hàng dưới mỗi hàng 2,3"); Eq(t.D[1][0], "S1", "S1 hàng 2"); Eq(Row(t, 3), "|||", "hàng 3 trống dưới S1"); Eq(t.D[3][0], "S2", "S2 hàng 4"); Eq(Row(t, 5), "|||", "hàng 5 trống dưới S2"); Eq(t.D[5][0], "S3", "S3 hàng 6");
        t = Make(); app = new MApp(t, new MSel(t, new int[] { 2, 4 }, 1, 1)); w = new WordTables(app);
        info = w.GetSelection(true); Eq(info.SelCells, 2, "chọn rời nhau: 2 ô"); string m2 = w.AddRows(1, false, ""); Eq(t.D.Count, 6, "chọn rời hàng 2 và 4: thêm 2 hàng");
        Eq(Row(t, 3), "|||", "hàng trống dưới hàng 2"); Eq(t.D[3][0], "S2", "S2 hàng 4"); Eq(t.D[4][0], "S3", "S3 hàng 5"); Eq(Row(t, 6), "|||", "hàng trống dưới hàng 4 cũ (cuối bảng)"); Console.WriteLine("  -> " + m2);
        t = Make(); app = new MApp(t, new MSel(t, 3, 1, 3, 1)); w = new WordTables(app);
        w.AddRows(1, true, "2,4"); Eq(t.D.Count, 6, "nhập tay 2,4: thêm phía trên hàng 2 và 4"); Eq(Row(t, 2), "|||", "hàng trống trên S1"); Eq(t.D[2][0], "S1", "S1 hàng 3"); Eq(t.D[3][0], "S2", "S2 hàng 4"); Eq(Row(t, 5), "|||", "hàng trống trên S3"); Eq(t.D[5][0], "S3", "S3 hàng 6");
        try { w.AddRows(1, true, "abc"); Eq(false, true, "danh sách sai phải báo lỗi"); } catch (UserMsg) { Eq(true, true, "danh sách sai báo lỗi"); }
        try { w.AddRows(1, true, "99"); Eq(false, true, "số ngoài bảng phải báo lỗi"); } catch (UserMsg) { Eq(true, true, "số ngoài bảng báo lỗi"); }
        t = Make(); app = new MApp(t, new MSel(t, 2, 2, 3, 3)); w = new WordTables(app);
        w.AddCols(1, true, ""); Eq(t.D[0].Count, 6, "chọn cột 2-3, thêm 1 cột trái mỗi cột: thêm 2 cột"); Eq(string.Join("|", t.D[0].ToArray()), "Section||Nominal||Measured|Loss", "thứ tự cột sau khi thêm bên trái");
        t = Make(); app = new MApp(t, new MSel(t, 2, 2, 3, 3)); w = new WordTables(app);
        w.AddCols(2, false, null); Eq(t.D[0].Count, 8, "thêm 2 cột phải mỗi cột 2,3: thêm 4 cột"); Eq(string.Join("|", t.D[0].ToArray()), "Section|Nominal|||Measured|||Loss", "thứ tự cột sau khi thêm bên phải");
        // 7. Xóa
        t = Make(); app = new MApp(t, new MSel(t, 2, 1, 3, 1)); w = new WordTables(app); w.DeleteRows(""); Eq(t.D.Count, 2, "xóa 2 hàng"); Eq(t.D[1][0], "S3", "còn S3");
        t = Make(); app = new MApp(t, new MSel(t, new int[] { 2, 4 }, 1, 1)); w = new WordTables(app); w.DeleteRows(null); Eq(t.D.Count, 2, "xóa hai hàng rời nhau"); Eq(t.D[1][0], "S2", "còn S2");
        t = Make(); app = new MApp(t, new MSel(t, 1, 2, 4, 3)); w = new WordTables(app); w.DeleteCols(""); Eq(t.D[0].Count, 2, "xóa 2 cột");
        t = Make(); app = new MApp(t, new MSel(t, 1, 1, 4, 4)); w = new WordTables(app);
        try { w.DeleteRows(""); Eq(false, true, "không được xóa hết hàng"); } catch (UserMsg) { Eq(true, true, "chặn xóa toàn bộ hàng"); }
        // 8. Định dạng số
        t = Make(); t.D[1][3] = "1,34"; app = new MApp(t, new MSel(t, 2, 2, 2, 4)); w = new WordTables(app);
        st.NumStyle = "us"; st.NumDec = "keep"; st.NumThou = true; st.NumAddDecToInt = false;
        w.FormatSelectedNumbers(st); Eq(t.D[1][3], "1.34", "định dạng 1,34 -> 1.34");
        // 9. CSV
        t = Make(); t.D[1][0] = "a,b"; app = new MApp(t, new MSel(t, 2, 2, 2, 2)); w = new WordTables(app);
        string csv = w.ToCsv(); Eq(csv.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length, 4, "CSV 4 dòng"); Eq(csv.Contains("\"a,b\""), true, "CSV bọc dấu phẩy");
        // 10. Master style: không lỗi khi gọi thuộc tính
        Rec.Log.Clear(); t = Make(); app = new MApp(t, new MSel(t, 2, 2, 2, 2)); w = new WordTables(app);
        try { w.ApplyToSelected(new AppSettings()); Eq(true, true, "áp master style không lỗi binder"); } catch (Exception ex) { Eq(ex.Message, "", "áp master style không lỗi binder"); }
        bool styleSet = false; foreach (string l in Rec.Log) if (l.Contains(".Style=TroLy_Master")) styleSet = true;
        Eq(styleSet, true, "gán Table Style TroLy_Master cho bảng");
        // 11. Master style: định dạng trực tiếp tiêu đề, font, zebra
        Rec.Log.Clear(); t = Make(); app = new MApp(t, new MSel(t, 2, 2, 2, 2)); w = new WordTables(app);
        var st2 = new AppSettings(); st2.Zebra = true; st2.FontSize = 10; st2.HdrBg = "156082"; st2.HdrFg = "FFFFFF"; st2.ZebraColor = "80FF80"; st2.BorderStyle = 1;
        w.ApplyToSelected(st2);
        int hdr = 0x15 + 0x60 * 256 + 0x82 * 65536, zeb = 0x80 + 0xFF * 256 + 0x80 * 65536;
        int cHdr = 0, cZeb = 0, cAuto = 0; bool size10 = false, white = false, dashed = false, bold = false;
        foreach (string l in Rec.Log)
        {
            if (l == "shading.BackgroundPatternColor=" + hdr) cHdr++;
            if (l == "shading.BackgroundPatternColor=" + zeb) cZeb++;
            if (l == "shading.BackgroundPatternColor=-16777216") cAuto++;
            if (l.EndsWith("Font.Size=10")) size10 = true;
            if (l.EndsWith("Font.Color=16777215")) white = true;
            if (l.EndsWith("Font.Bold=-1")) bold = true;
            if (l.EndsWith(".LineStyle=3")) dashed = true;
        }
        Eq(cHdr, 4, "tô nền tiêu đề: 4 ô hàng 1"); Eq(cZeb, 4, "zebra: hàng dữ liệu thứ 2 (hàng 3) và thứ 4 (hàng 5 không có) -> chỉ hàng 3 có 4 ô"); Eq(cAuto, 8, "hàng dữ liệu lẻ (hàng 2 và 4) không tô");
        Eq(size10, true, "cỡ chữ 10 được gán trực tiếp"); Eq(white, true, "chữ tiêu đề màu trắng"); Eq(bold, true, "tiêu đề đậm"); Eq(dashed, true, "viền đứt dùng LineStyle=3");
        Eq(WordMap.LineStyleEnum(0), 1, "viền liền = 1"); Eq(WordMap.LineStyleEnum(2), 2, "viền chấm = 2");

        // 12. Master style: áp dụng lần hai không ghi gì thêm (đọc, so sánh, chỉ ghi chỗ khác)
        Rec.Log.Clear(); t = Make(); app = new MApp(t, new MSel(t, 2, 2, 2, 2)); w = new WordTables(app);
        var st3 = new AppSettings(); st3.Zebra = true;
        string first = w.ApplyToSelected(st3);
        int writes1 = 0; foreach (string l in Rec.Log) if (l.Contains("=") && !l.Contains("(")) writes1++;
        Rec.Log.Clear();
        string second = w.ApplyToSelected(st3);
        int writes2 = 0; foreach (string l in Rec.Log) if (l.Contains("=") && !l.Contains("(")) writes2++;
        Eq(writes1 > 20, true, "lần đầu ghi nhiều thuộc tính (" + writes1 + ")"); Eq(writes2, 0, "lần hai không ghi thuộc tính nào"); Eq(second.Contains("Không có gì cần đổi"), true, "thông báo lần hai: không có gì cần đổi");
        Console.WriteLine("  -> lần 1: " + first); Console.WriteLine("  -> lần 2: " + second);
        // đổi một thông số thì chỉ ghi đúng phần đó
        Rec.Log.Clear(); st3.FontSize = 12; w.ApplyToSelected(st3);
        int writes3 = 0; foreach (string l in Rec.Log) if (l.Contains("=") && !l.Contains("(")) writes3++;
        Eq(writes3 <= 3, true, "đổi cỡ chữ: chỉ ghi vài thuộc tính (" + writes3 + ")");
        // 13. Định dạng số: các trường hợp trong ảnh (5478,80 / 7,900 / 0,000 / 21,000)
        var st4 = new AppSettings(); st4.NumStyle = "si"; st4.NumDec = "2"; st4.NumThou = true; st4.NumAddDecToInt = true;
        t = new MTable(new string[][] { new string[] { "Tiêu đề" }, new string[] { "5478,80" }, new string[] { "7,900" }, new string[] { "0,000" }, new string[] { "21,000" }, new string[] { "1234" }, new string[] { "chữ" } });
        app = new MApp(t, new MSel(t, 2, 1, 7, 1)); w = new WordTables(app);
        string nm = w.FormatSelectedNumbers(st4);
        Eq(string.Join("|", new string[] { t.D[1][0], t.D[2][0], t.D[3][0], t.D[4][0], t.D[5][0], t.D[6][0] }), "5\u00A0478,80|7,90|0,00|21,00|1\u00A0234,00|chữ", "SI, 2 thập phân, thêm thập phân cho số nguyên: mọi ô cùng một quy ước");
        Console.WriteLine("  -> " + nm);
        string nm2 = w.FormatSelectedNumbers(st4); Eq(nm2.Contains("0 ô") || nm2.Contains("Đã định dạng 0 ô"), true, "chạy lại không đổi ô nào"); Console.WriteLine("  -> " + nm2);
        // kiểu Mỹ, giữ nguyên thập phân, từ cùng cột: số 5478,80 cho biết dấu phẩy là thập phân nên 7,900 là 7.900
        st4.NumStyle = "us"; st4.NumDec = "keep"; st4.NumAddDecToInt = false;
        t = new MTable(new string[][] { new string[] { "H" }, new string[] { "5478,80" }, new string[] { "7,900" }, new string[] { "21,000" } });
        app = new MApp(t, new MSel(t, 2, 1, 4, 1)); w = new WordTables(app);
        w.FormatSelectedNumbers(st4);
        Eq(string.Join("|", new string[] { t.D[1][0], t.D[2][0], t.D[3][0] }), "5,478.80|7.900|21.000", "kiểu Mỹ: dùng bằng chứng của số 5478,80");
        // chọn nhiều hàng một lúc: mọi ô được xử lý (không bỏ ô nào)
        var rows = new List<string[]>(); rows.Add(new string[] { "H1", "H2" });
        for (int i = 1; i <= 30; i++) rows.Add(new string[] { (i * 1111).ToString(), (i * 12.5).ToString(System.Globalization.CultureInfo.InvariantCulture) });
        t = new MTable(rows.ToArray()); app = new MApp(t, new MSel(t, 2, 1, 31, 2)); w = new WordTables(app);
        st4.NumStyle = "vn"; st4.NumDec = "1"; st4.NumAddDecToInt = true; st4.NumThou = true;
        w.FormatSelectedNumbers(st4);
        int okc = 0; for (int i = 1; i <= 30; i++) { string a = t.D[i][0], b = t.D[i][1]; if (a.EndsWith(",0") && b.Contains(",")) okc++; }
        Eq(okc, 30, "30 hàng x 2 cột: không bỏ sót hàng nào");
        Eq(t.D[30][0], "33.330,0", "33330 kiểu Việt 1 thập phân"); Eq(t.D[2][1], "25,0", "25 -> 25,0");

        // 13b. Rev0.4: mô-đun Định dạng số (WordNumbers) trong bảng: vòng kiểu, tăng giảm thập phân, nhiều ô cùng lúc
        {
            var sn = new AppSettings(); sn.NumStyle = "us"; sn.NumThou = true;
            t = new MTable(new string[][] { new string[] { "Tiêu đề", "B" }, new string[] { "1,234.56", "0.5" }, new string[] { "7,890.1", "12" }, new string[] { "chữ", "3.25 mm" } });
            app = new MApp(t, new MSel(t, 2, 1, 4, 2)); var wn = new WordNumbers(app);
            string r1 = wn.Apply(sn, "cycle"); Console.WriteLine("  -> " + r1);
            Eq(sn.NumStyle, "si", "Mỹ sang SI, ghi nhớ kiểu mới");
            Eq(t.D[1][0].Replace("\u00A0", " ").Replace("\u202F", " "), "1 234,56", "ô 1,234.56 thành 1 234,56"); Eq(t.D[1][1], "0,5", "0.5 thành 0,5");
            Eq(t.D[3][1], "3,25 mm", "số kèm đơn vị trong ô vẫn đổi, giữ đơn vị"); Eq(t.D[3][0], "chữ", "ô chữ không đổi");
            wn.Apply(sn, "cycle"); Eq(sn.NumStyle, "vn", "SI sang Việt"); Eq(t.D[2][0], "7.890,1", "7,890.1 qua SI rồi Việt");
            wn.Apply(sn, "inc"); Eq(t.D[1][1], "0,50", "tăng thập phân 0,5 thành 0,50"); Eq(t.D[2][1], "12,0", "số nguyên 12 thành 12,0");
            wn.Apply(sn, "dec"); wn.Apply(sn, "dec"); Eq(t.D[2][1], "12", "giảm về số nguyên"); Eq(t.D[1][1], "1", "0,50 giảm hai lần: làm tròn 0,5 thành 1");
            // một cột
            t = new MTable(new string[][] { new string[] { "H" }, new string[] { "1.5" }, new string[] { "2.25" }, new string[] { "30" } });
            app = new MApp(t, new MSel(t, 2, 1, 4, 1)); wn = new WordNumbers(app); sn.NumStyle = "us";
            wn.Apply(sn, "inc"); Eq(t.D[1][0] + "|" + t.D[2][0] + "|" + t.D[3][0], "1.50|2.250|30.0", "cả cột tăng một chữ số thập phân");
            // chỉ đặt con trỏ trong một ô
            app = new MApp(t, new MSel(t, 3, 1, 3, 1)); wn = new WordNumbers(app); wn.Apply(sn, "dec");
            Eq(t.D[2][0], "2.25", "con trỏ trong một ô: chỉ ô đó giảm"); Eq(t.D[1][0], "1.50", "ô khác giữ nguyên");
            // ô trống, chữ: báo không có số
            t = new MTable(new string[][] { new string[] { "H" }, new string[] { "abc" }, new string[] { "" } });
            app = new MApp(t, new MSel(t, 2, 1, 3, 1)); wn = new WordNumbers(app);
            Eq(wn.Apply(sn, "inc").StartsWith("Không có số"), true, "ô không có số thì báo rõ");
        }
        // 13c. Rev0.6: Đồng bộ số lẻ, giữ vùng bôi đen, Lặp tiêu đề nhiều bảng, "Bảng này" không đụng style dùng chung
        {
            var sy = new AppSettings(); sy.NumStyle = "us"; sy.NumThou = true; sy.SyncDec = 2;
            t = new MTable(new string[][] { new string[] { "H", "K" }, new string[] { "1.5", "7" }, new string[] { "2.255", "30.1" }, new string[] { "1,234.5", "chữ" } });
            app = new MApp(t, new MSel(t, 2, 1, 4, 2)); var wn = new WordNumbers(app);
            Rec.Log.Clear();
            string rs = wn.Apply(sy, "sync"); Console.WriteLine("  -> " + rs);
            Eq(t.D[1][0] + "|" + t.D[1][1], "1.50|7.00", "đồng bộ 2 số lẻ: 1.5 thành 1.50, số nguyên 7 thành 7.00");
            Eq(t.D[2][0] + "|" + t.D[2][1], "2.26|30.10", "đồng bộ: 2.255 làm tròn 2.26, 30.1 thành 30.10");
            Eq(t.D[3][0] + "|" + t.D[3][1], "1,234.50|chữ", "đồng bộ giữ dấu hàng nghìn, ô chữ không đổi");
            Eq(rs.Contains("đúng 2 chữ số thập phân"), true, "báo đồng bộ 2 chữ số");
            Eq(Rec.Log.Contains("sel.SetRange(201,403)"), true, "giữ vùng bôi đen: chọn lại từ ô (2,1) đến hết ô (4,2)");
            sy.SyncDec = 3; wn.Apply(sy, "sync"); Eq(t.D[1][0] + "|" + t.D[1][1], "1.500|7.000", "đồng bộ 3 số lẻ"); Eq(t.D[2][0], "2.260", "2.26 thành 2.260");
            sy.SyncDec = 1; wn.Apply(sy, "sync"); Eq(t.D[1][0] + "|" + t.D[1][1] + "|" + t.D[2][1], "1.5|7.0|30.1", "đồng bộ 1 số lẻ (làm tròn bớt)");
            string again = wn.Apply(sy, "sync"); Eq(again.Contains("đã có đúng 1 chữ số"), true, "chạy lại không đổi gì");
            sy.NumStyle = "vn"; t = new MTable(new string[][] { new string[] { "H" }, new string[] { "1.234,5" }, new string[] { "7,25" }, new string[] { "40" } });
            app = new MApp(t, new MSel(t, 2, 1, 4, 1)); wn = new WordNumbers(app); sy.SyncDec = 3; wn.Apply(sy, "sync");
            Eq(t.D[1][0] + "|" + t.D[2][0] + "|" + t.D[3][0], "1.234,500|7,250|40,000", "kiểu Việt: đồng bộ 3 số lẻ giữ kiểu dấu");
            // Thêm, Bớt, Dấu cũng phải giữ vùng bôi đen
            Rec.Log.Clear(); sy.NumStyle = "us"; wn.Apply(sy, "inc"); Eq(Rec.Log.Contains("sel.SetRange(201,402)"), true, "Thêm số lẻ giữ vùng bôi đen");
            Rec.Log.Clear(); wn.Apply(sy, "dec"); Eq(Rec.Log.Contains("sel.SetRange(201,402)"), true, "Bớt số lẻ giữ vùng bôi đen");
            Rec.Log.Clear(); wn.Apply(sy, "cycle"); Eq(Rec.Log.Contains("sel.SetRange(201,402)"), true, "Đổi dấu giữ vùng bôi đen");
            // vùng chọn rời nhau: không dựng lại bằng hình chữ nhật (tránh chọn thừa ô)
            t = new MTable(new string[][] { new string[] { "H" }, new string[] { "1.5" }, new string[] { "2.5" }, new string[] { "3.5" } });
            app = new MApp(t, new MSel(t, new int[] { 2, 4 }, 1, 1)); wn = new WordNumbers(app); Rec.Log.Clear(); wn.Apply(sy, "inc");
            Eq(Rec.Log.Exists(x => x.StartsWith("sel.SetRange(")), false, "chọn rời nhau: không chọn lại hình chữ nhật bao ngoài");

            // Lặp tiêu đề cho nhiều bảng
            var t1 = Make(); var t2 = Make(); var t3 = Make();
            var apm = new MApp2(new[] { t1, t2, t3 }, new[] { t1, t2 });
            var wm = new WordTables(apm);
            string rh = wm.SetRepeatHeader("sel"); Console.WriteLine("  -> " + rh);
            Eq(Hf(t1) && Hf(t2), true, "bôi đen qua 2 bảng: cả hai bảng lặp tiêu đề");
            Eq(HfSet(t3), false, "bảng thứ 3 ngoài vùng chọn không bị đụng");
            rh = wm.SetRepeatHeader("sel"); Eq(Hf(t1) || Hf(t2), false, "bấm lần nữa: tắt cho cả hai bảng"); Eq(rh.Contains("Đã tắt"), true, "báo đã tắt");
            wm.SetRepeatHeader("allOn"); Eq(Hf(t1) && Hf(t2) && Hf(t3), true, "bật cho mọi bảng");
            wm.SetRepeatHeader("allOff"); Eq(Hf(t3), false, "tắt cho mọi bảng");
            ((Rec)t1.Rows.Item(1)).P["HeadingFormat"] = true; ((Rec)t2.Rows.Item(1)).P["HeadingFormat"] = false;
            wm.SetRepeatHeader("sel"); Eq(Hf(t1) && Hf(t2), true, "một bảng đã bật, một chưa: bấm thì bật hết");
            bool thrown = false; try { new WordTables(new MApp2(new[] { t1 }, new MTable[0])).SetRepeatHeader("sel"); } catch (UserMsg) { thrown = true; }
            Eq(thrown, true, "con trỏ ngoài mọi bảng: báo rõ cách dùng");

            // "Bảng này" không ghi vào Table Style dùng chung; "Mọi bảng" thì có
            var ts = new AppSettings(); ts.HdrBg = "C00000";
            t = Make(); app = new MApp(t, new MSel(t, 2, 1, 2, 2)); w = new WordTables(app);
            Rec.Log.Clear(); w.ApplyToSelected(ts);
            Eq(Rec.Log.Exists(x => x.Contains("Condition(0)")), false, "Bảng này: không đổi màu nền dòng tiêu đề của Table Style dùng chung");
            Rec.Log.Clear(); w.ApplyAll(ts);
            Eq(Rec.Log.Exists(x => x.Contains("Condition(0)")), true, "Mọi bảng: cập nhật Table Style");
        }
        // 14. Chọn rời nhau bằng Ctrl: Word chỉ cho vùng chọn cuối; bộ theo dõi cộng dồn các vùng
        {
            bool ctrl = false; long clock = 1000;
            Keyboard.CtrlDown = delegate { return ctrl; }; Keyboard.NowMs = delegate { return clock; };
            WordTables.Tracker.Reset();
            t = MakeBig(); var sl = new MSel(t, 2, 1, 2, 4); app = new MApp(t, sl); w = new WordTables(app);
            w.Poll();                                    // chọn thường hàng 2
            clock += 700; ctrl = true; clock += 60; sl.R1 = 4; sl.R2 = 4; w.Poll();     // Ctrl + chọn hàng 4
            clock += 60; sl.R1 = 6; sl.R2 = 6; w.Poll(); 
            ctrl = false; clock += 300;
            SelInfo si = w.GetSelection(false);
            Eq(IndexList.Format(si.Rows), "2, 4, 6", "Ctrl: cộng dồn hàng 2, 4, 6");
            Eq(si.SelCells, 12, "Ctrl: 3 hàng x 4 cột = 12 ô");
            string msg2 = w.AddRows(1, false, "");
            Eq(msg2.Contains("hàng 2,4,6") || msg2.Contains("2, 4, 6"), true, "Thêm hàng dưới mỗi hàng 2, 4, 6"); Console.WriteLine("  -> " + msg2);
            Eq(t.D.Count, 11, "8 hàng + 3 hàng mới = 11");
            // Ctrl nhả ra rồi chọn vùng khác: đặt lại
            WordTables.Tracker.Reset();
            t = MakeBig(); sl = new MSel(t, 2, 1, 2, 4); app = new MApp(t, sl); w = new WordTables(app);
            clock += 1000; w.Poll(); clock += 1000; sl.R1 = 4; sl.R2 = 4; w.Poll();       // không giữ Ctrl: thay thế
            Eq(IndexList.Format(w.GetSelection(false).Rows), "4", "không giữ Ctrl: vùng mới thay vùng cũ");
            // Ctrl vừa nhả trong 450 ms vẫn tính là Ctrl (nhấp nhanh)
            WordTables.Tracker.Reset();
            t = MakeBig(); sl = new MSel(t, 1, 1, 1, 4); app = new MApp(t, sl); w = new WordTables(app);
            clock += 1000; w.Poll();
            ctrl = true; clock += 30; w.Poll(); ctrl = false; clock += 40; sl.R1 = 3; sl.R2 = 3; w.Poll();
            Eq(IndexList.Format(w.GetSelection(false).Rows), "1, 3", "Ctrl nhả ngay sau khi nhấp vẫn cộng dồn");
            // chọn cột rời nhau 1, 2, 4
            WordTables.Tracker.Reset(); ctrl = false;
            t = MakeBig(); sl = new MSel(t, 1, 1, 6, 2); app = new MApp(t, sl); w = new WordTables(app);
            clock += 1000; w.Poll(); ctrl = true; clock += 60; sl.C1 = 4; sl.C2 = 4; w.Poll(); ctrl = false; clock += 500;
            SelInfo sc = w.GetSelection(false); Eq(IndexList.Format(sc.Cols), "1, 2, 4", "cột 1, 2, 4 rời nhau");
            string m3 = w.AddCols(1, false, ""); Console.WriteLine("  -> " + m3);
            Eq(m3.Contains("1, 2, 4") || m3.Contains("1,2,4"), true, "thêm cột bên phải mỗi cột 1, 2, 4");
            Keyboard.CtrlDown = delegate { return false; };
            WordTables.Tracker.Reset();
        }
        TestProdApplyNow();
        TestIndent();
        TestAlignMulti();
        TestMergedCells();
        TestCurrentParagraph();
        Eq(IndexList.Format(IndexList.Parse("8,3-5,4", 20)), "3-5, 8", "danh sách số 3-5, 8"); Eq(IndexList.Parse("0", 5) == null, true, "số 0 không hợp lệ"); Eq(IndexList.Parse("6", 5) == null, true, "số vượt giới hạn");
        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
    // Rev0.9.1C: kiểm thử Tích theo định dạng và áp dụng ngay (tách khỏi Main để không trùng tên biến)
    // Rev0.9.5.6: Chuẩn hóa đoạn văn có thụt dòng đầu (First line) và thụt treo (Hanging)
    static void TestCurrentParagraph()
    {
        var app = new MLayoutApp();
        var first = new MLayoutPara("Times New Roman", 10f, false);
        var current = new MLayoutPara("Arial", 14f, false);
        var last = new MLayoutPara("Calibri", 15f, false);
        app.ActiveDocument.Paragraphs.Items.AddRange(new[] { first, current, last });
        app.Selection.Paragraphs.Items.Add(current);
        var message = new WordLayout(app).StandardizeParagraphs(new AppSettings());
        Eq(current.Range.Font.Size, 11f, "đoạn chứa con trỏ được chỉnh cỡ Arial");
        Eq(current.Format.Alignment, 3, "đoạn chứa con trỏ được căn đều hai bên");
        Eq(first.Range.Font.Size, 10f, "đoạn trước con trỏ giữ nguyên");
        Eq(last.Format.LineSpacing, 0f, "đoạn sau con trỏ giữ nguyên");
        Eq(message.Contains("đoạn tại con trỏ"), true, "thông báo chỉ rõ phạm vi tại con trỏ");
        current.Range.Style.AutomaticallyUpdate = true;
        new WordLayout(app).StandardizeParagraphs(new AppSettings());
        Eq(current.Range.Style.AutomaticallyUpdate, true, "khôi phục AutomaticallyUpdate sau khi định dạng");
        current.Format.FailAlignment = true;
        try { new WordLayout(app).StandardizeParagraphs(new AppSettings()); Eq(false, true, "lỗi định dạng phải được báo"); }
        catch (InvalidOperationException) { Eq(true, true, "lỗi định dạng phải được báo"); }
        Eq(current.Range.Style.AutomaticallyUpdate, true, "khôi phục AutomaticallyUpdate cả khi định dạng lỗi");
        current.Format.FailAlignment = false;

        // Rev0.9.5.13: bôi đen nhiều đoạn thì áp cho TẤT CẢ các đoạn đó cùng lúc (không còn chỉ đoạn tại đầu hoạt động)
        app.Selection.Paragraphs.Items.Clear();
        app.Selection.Paragraphs.Items.Add(first);
        app.Selection.Paragraphs.Items.Add(last);
        app.Selection.Start = 1; app.Selection.End = 100;
        app.Selection.StartIsActive = false;
        var msgMulti = new WordLayout(app).StandardizeParagraphs(new AppSettings());
        Eq(last.Format.Alignment, 3, "vùng chọn nhiều đoạn: đoạn cuối được căn đều hai bên");
        Eq(first.Format.Alignment, 3, "vùng chọn nhiều đoạn: đoạn đầu cũng được căn đều hai bên");
        Eq(first.Range.Font.Size, 13f, "vùng chọn nhiều đoạn: đoạn Times New Roman đổi 13 pt");
        Eq(last.Range.Font.Size, 15f, "vùng chọn nhiều đoạn: đoạn Calibri giữ nguyên cỡ");
        Eq(current.Format.Alignment, 3, "đoạn không nằm trong vùng chọn giữ kết quả cũ");
        Eq(msgMulti.Contains("2 đoạn trong vùng chọn"), true, "thông báo nêu số đoạn trong vùng chọn");
        app.Selection.StartIsActive = true;
        new WordLayout(app).StandardizeParagraphs(new AppSettings());
        Eq(first.Format.Alignment, 3, "vùng chọn ngược cũng áp cho mọi đoạn");
        Eq(app.Selection.Start == 1 && app.Selection.End == 100, true, "giữ vùng chọn gốc sau chuẩn hóa");
    }

    // Rev0.9.5.13: bảng có ô gộp. Word đánh số ô theo vị trí trong từng hàng; hàng gộp A:C chỉ có 3 ô nên D nằm ở vị trí 3, không phải 5.
    static MTable MakeMerged()
    {
        var t = new MTable(new string[][] {
            new string[] { "", "A", "B", "C", "D" },
            new string[] { "1", "", "1", "2", "2" },
            new string[] { "2", "", "4", "5", "20" },
            new string[] { "3", "", "6", "8", "48" },
            new string[] { "4", "", "", "", "70" },
            new string[] { "5", "", "" } });                 // hàng 6: A6:C6 gộp, D6 là ô thứ 3
        t.Wd = new List<List<double>> {
            new List<double> { 40, 100, 100, 100, 100 }, new List<double> { 40, 100, 100, 100, 100 }, new List<double> { 40, 100, 100, 100, 100 },
            new List<double> { 40, 100, 100, 100, 100 }, new List<double> { 40, 100, 100, 100, 100 }, new List<double> { 40, 300, 100 } };
        return t;
    }

    static void TestMergedCells()
    {
        var st = new AppSettings(); st.CalcDec = "auto";
        // 0) mô hình Word giả đúng như Word: Columns.Count báo lỗi, Cell(6,5) không tồn tại
        var t0 = MakeMerged();
        Eq(t0.Uniform, false, "ô gộp: Table.Uniform = false như Word");
        Eq(t0.Columns.Count * t0.Rows.Count != t0.Range.Cells.Count, true, "ô gộp: số ô thật khác số hàng nhân số cột (Columns.Count vẫn trả 5 như Word thật)");
        bool cellThrow = false; try { t0.Cell(6, 5); } catch (ArgumentException) { cellThrow = true; }
        Eq(cellThrow, true, "ô gộp: hàng 6 chỉ có 3 ô nên Cell(6,5) không tồn tại như Word");

        // 1) Tổng D2:D5 (thẳng cột) ghi vào ô D6 nằm ở vị trí 3 của hàng gộp
        var t = MakeMerged();
        var app = new MApp(t, new MSel(t, 2, 5, 5, 5));
        st.CalcOp = "sum";
        string msg = new WordTables(app).Calc(st);
        Eq(t.D[5][2], "140", "ô gộp: tổng D2:D5 = 2+20+48+70 ghi vào D6 (ô thứ 3 của hàng gộp)");
        Eq(t.D[5][0] + "|" + t.D[5][1], "5|", "ô gộp: các ô khác của hàng gộp không bị đổi");
        Eq(msg.Contains("ô gộp"), true, "ô gộp: thông báo nêu bảng có ô gộp");
        Eq(t.D.Count, 6, "ô gộp: không thêm hàng thừa");

        // 2) Trung bình, nhỏ nhất, lớn nhất, đếm số cùng đúng
        foreach (var c in new[] { new object[] { "avg", "35.00" }, new object[] { "min", "2" }, new object[] { "max", "70" }, new object[] { "cnt", "4" } })
        {
            var tt = MakeMerged();
            st.CalcOp = (string)c[0];
            new WordTables(new MApp(tt, new MSel(tt, 2, 5, 5, 5))).Calc(st);
            Eq(tt.D[5][2], (string)c[1], "ô gộp: " + c[0] + " D2:D5 = " + c[1]);
        }

        // 3) Hàng: B3:C3 (4 và 5) ghi vào ô bên phải D3
        var t3 = MakeMerged(); st.CalcOp = "sum";
        new WordTables(new MApp(t3, new MSel(t3, 3, 3, 3, 4))).Calc(st);
        Eq(t3.D[2][4], "9", "ô gộp: tổng theo hàng B3:C3 ghi vào ô bên phải");

        // 4) Hết bảng: chọn cột cuối của bảng có hàng gộp phía trên thì thêm hàng mới và ghi vào ô thẳng cột
        var t4 = new MTable(new string[][] {
            new string[] { "", "A", "B", "C", "D" },
            new string[] { "5", "", "" },
            new string[] { "1", "", "1", "2", "3" },
            new string[] { "2", "", "4", "5", "4" } });
        t4.Wd = new List<List<double>> {
            new List<double> { 40, 100, 100, 100, 100 }, new List<double> { 40, 300, 100 },
            new List<double> { 40, 100, 100, 100, 100 }, new List<double> { 40, 100, 100, 100, 100 } };
        string m4 = new WordTables(new MApp(t4, new MSel(t4, 3, 5, 4, 5))).Calc(st);
        Eq(t4.D.Count, 5, "ô gộp: thêm một hàng mới khi hết bảng");
        Eq(t4.D[4][4], "7", "ô gộp: tổng D3:D4 ghi vào ô D của hàng mới thêm");

        // 5) Hàng cuối không còn chỗ ghi (không thêm được cột) thì báo rõ, không sửa bảng
        var t5 = MakeMerged();
        string m5 = new WordTables(new MApp(t5, new MSel(t5, 3, 3, 3, 5))).Calc(st);
        Eq(m5.Contains("không thêm được cột"), true, "ô gộp: hết cột và không thêm được thì báo rõ");
        Eq(t5.D[2][4], "20", "ô gộp: bảng không bị sửa khi không có ô để ghi");

        // 6) Tích theo định dạng nhiều hàng trong bảng có ô gộp: B, C đậm, D nhận tích (1x2, 4x5, 6x8)
        var t6 = MakeMerged();
        var bold = new Rec("font"); bold.P["Bold"] = -1;
        for (int r = 2; r <= 4; r++) { t6.Cell(r, 3).Range.P["Font"] = bold; t6.Cell(r, 4).Range.P["Font"] = bold; }
        var st6 = new AppSettings(); st6.CalcFmtKind = "bold";
        string m6 = new WordTables(new MApp(t6, new MSel(t6, 2, 5, 4, 5))).CalcProdFmt(st6);
        Eq(t6.D[1][4] + "|" + t6.D[2][4] + "|" + t6.D[3][4], "2|20|48", "ô gộp: Tích theo định dạng nhiều hàng ra 2, 20, 48");
        Eq(m6.Contains("3 hàng"), true, "ô gộp: thông báo Tích nêu 3 hàng");

        // 7) Tích theo định dạng ngay trên hàng gộp: các ô của hàng gộp chỉ có 3 ô
        var t7 = new MTable(new string[][] { new string[] { "5", "6", "" } });
        t7.Wd = new List<List<double>> { new List<double> { 40, 300, 100 } };
        t7.Cell(1, 1).Range.P["Font"] = bold; t7.Cell(1, 2).Range.P["Font"] = bold;
        var t7b = new MTable(new string[][] { new string[] { "a", "b", "c", "d" }, t7.D[0].ToArray() });
        t7b.Wd = new List<List<double>> { new List<double> { 100, 100, 100, 100 }, new List<double> { 40, 300, 100 } };
        t7b.Cell(2, 1).Range.P["Font"] = bold; t7b.Cell(2, 2).Range.P["Font"] = bold;
        new WordTables(new MApp(t7b, new MSel(t7b, 2, 3, 2, 3))).CalcProdFmt(st6);
        Eq(t7b.D[1][2], "30", "ô gộp: Tích 5 x 6 ghi vào ô thứ 3 của hàng gộp");
    }

    // Rev0.9.5.13: canh lề chọn được trong Cài đặt; bôi đen nhiều đoạn áp cùng lúc
    static void TestAlignMulti()
    {
        var st = new AppSettings();
        Eq(st.ParaAlign, "justify", "canh lề mặc định là hai bên");
        // 1) bốn kiểu canh lề cho một đoạn tại con trỏ
        foreach (var c in new[] { new object[] { "left", 0 }, new object[] { "center", 1 }, new object[] { "right", 2 }, new object[] { "justify", 3 } })
        {
            var app = new MLayoutApp();
            var p = new MLayoutPara("Times New Roman", 12f, false);
            app.ActiveDocument.Paragraphs.Items.Add(p);
            app.Selection.Paragraphs.Items.Add(p);
            var s = new AppSettings(); s.ParaAlign = (string)c[0];
            string m = new WordLayout(app).StandardizeParagraphs(s);
            Eq(p.Format.Alignment, (int)c[1], "canh lề " + c[0] + " áp cho đoạn tại con trỏ");
        }
        // 2) nhiều đoạn có tiêu đề, ô bảng, danh sách: chỉ đoạn thường và danh sách được định dạng; tiêu đề và ô bảng bị bỏ qua
        var app2 = new MLayoutApp();
        var body1 = new MLayoutPara("Times New Roman", 12f, false);
        var head = new MLayoutPara("Arial", 11f, false, 1, 0);
        var cell = new MLayoutPara("Times New Roman", 12f, true);
        var li = new MLayoutPara("Times New Roman", 12f, false, 10, 2);
        var body2 = new MLayoutPara("Arial", 14f, false);
        app2.ActiveDocument.Paragraphs.Items.AddRange(new[] { body1, head, cell, li, body2 });
        foreach (var x in new[] { body1, head, cell, li, body2 }) app2.Selection.Paragraphs.Items.Add(x);
        app2.Selection.Start = 0; app2.Selection.End = 500;
        var sc = new AppSettings(); sc.ParaAlign = "center";
        string m2 = new WordLayout(app2).StandardizeParagraphs(sc);
        Eq(body1.Format.Alignment + "|" + body2.Format.Alignment + "|" + li.Format.Alignment, "1|1|1", "nhiều đoạn: đoạn thường và danh sách được căn giữa");
        Eq(head.Format.Alignment + "|" + cell.Format.Alignment, "0|0", "nhiều đoạn: tiêu đề và ô bảng giữ nguyên căn lề");
        Eq(head.Range.Font.Size + "|" + cell.Range.Font.Size, "11|12", "nhiều đoạn: tiêu đề và ô bảng giữ nguyên cỡ chữ");
        Eq(Math.Abs(body1.Format.FirstLineIndent - 35.43f) < 0.05f && Math.Abs(body2.Format.FirstLineIndent - 35.43f) < 0.05f, true, "nhiều đoạn: đoạn thường thụt dòng đầu 1,25 cm");
        Eq(li.Format.FirstLineIndent, 0f, "nhiều đoạn: đoạn danh sách không thụt dòng đầu");
        Eq(body1.Range.Font.Size + "|" + body2.Range.Font.Size, "13|11", "nhiều đoạn: mỗi đoạn đổi cỡ chữ theo font của chính nó");
        Eq(m2.Contains("3 đoạn") && m2.Contains("bỏ qua 2"), true, "nhiều đoạn: thông báo 3 đoạn, bỏ qua 2 (tiêu đề, ô bảng)");
        Eq(m2.Contains("căn giữa"), true, "nhiều đoạn: thông báo nêu kiểu canh lề");

        // 3) mỗi style đều tạm tắt AutomaticallyUpdate rồi khôi phục (kể cả khi lỗi giữa chừng)
        var app3 = new MLayoutApp();
        var q1 = new MLayoutPara("Times New Roman", 12f, false); q1.Range.Style.NameLocal = "Normal"; q1.Range.Style.AutomaticallyUpdate = true;
        var q2 = new MLayoutPara("Times New Roman", 12f, false); q2.Range.Style.NameLocal = "Body"; q2.Range.Style.AutomaticallyUpdate = true;
        var q3 = new MLayoutPara("Times New Roman", 12f, false); q3.Range.Style.NameLocal = "Body2"; q3.Range.Style.AutomaticallyUpdate = false;
        app3.ActiveDocument.Paragraphs.Items.AddRange(new[] { q1, q2, q3 });
        foreach (var x in new[] { q1, q2, q3 }) app3.Selection.Paragraphs.Items.Add(x);
        app3.Selection.Start = 0; app3.Selection.End = 300;
        new WordLayout(app3).StandardizeParagraphs(new AppSettings());
        Eq(q1.Range.Style.AutomaticallyUpdate && q2.Range.Style.AutomaticallyUpdate && !q3.Range.Style.AutomaticallyUpdate, true, "nhiều đoạn: AutomaticallyUpdate các style được khôi phục đúng trạng thái cũ");
        q2.Format.FailAlignment = true;
        try { new WordLayout(app3).StandardizeParagraphs(new AppSettings()); Eq(false, true, "lỗi giữa chừng phải được báo"); }
        catch (InvalidOperationException) { Eq(true, true, "lỗi giữa chừng phải được báo"); }
        Eq(q1.Range.Style.AutomaticallyUpdate && q2.Range.Style.AutomaticallyUpdate, true, "nhiều đoạn: khôi phục AutomaticallyUpdate cả khi lỗi giữa chừng");

        // 4) nhiều đoạn lớn (800 đoạn) vẫn xong và đúng
        var app4 = new MLayoutApp();
        var all = new List<MLayoutPara>();
        for (int i = 0; i < 800; i++) { var q = new MLayoutPara(i % 2 == 0 ? "Times New Roman" : "Arial", 12f, false); all.Add(q); app4.ActiveDocument.Paragraphs.Items.Add(q); app4.Selection.Paragraphs.Items.Add(q); }
        app4.Selection.Start = 0; app4.Selection.End = 99999;
        string m4 = new WordLayout(app4).StandardizeParagraphs(new AppSettings());
        int okCount = 0; foreach (var q in all) if (q.Format.Alignment == 3 && (q.Range.Font.Size == 13f || q.Range.Font.Size == 11f)) okCount++;
        Eq(okCount, 800, "800 đoạn: tất cả được chuẩn hóa");
        Eq(m4.Contains("800 đoạn"), true, "800 đoạn: thông báo đúng số đoạn");
    }


    static void TestIndent()
    {
        // 1) Mặc định: thụt dòng đầu 1,25 cm = 35,43 pt cho đoạn thường; tiêu đề, bảng, danh sách giữ nguyên
        var app = new MLayoutApp();
        app.ActiveDocument.Sections.Items.Add(new MLayoutSection(1, 0));
        app.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Times New Roman", 12f, false));            // 0: thường
        app.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Arial", 11f, false, 1, 0));                  // 1: tiêu đề (OutlineLevel 1)
        app.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Times New Roman", 12f, true));              // 2: trong bảng
        app.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Times New Roman", 12f, false, 10, 2));      // 3: danh sách
        app.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Calibri", 11f, false));                      // 4: thường, font khác
        var st = new AppSettings();
        Eq(Math.Abs(st.ParaFirstLineCm - 1.25) < 1e-9 && st.ParaHangingCm == 0, true, "thụt lề: mặc định First line 1,25 cm, Hanging 0 cm");
        var wl = new WordLayout(app);
        string msg = wl.StandardizeParagraphs(st);
        var P = app.ActiveDocument.Paragraphs.Items;
        Eq(Math.Abs(P[0].Format.FirstLineIndent - 35.43f) < 0.05f, true, "thụt dòng đầu 1,25 cm = 35,43 pt cho đoạn thường");
        Eq(P[4].Format.FirstLineIndent, 0f, "đoạn font khác ngoài con trỏ giữ nguyên");
        Eq(P[1].Format.FirstLineIndent, 0f, "tiêu đề không thụt dòng đầu");
        Eq(P[2].Format.FirstLineIndent, 0f, "đoạn trong bảng không thụt dòng đầu");
        Eq(P[3].Format.FirstLineIndent, 0f, "đoạn danh sách không thụt dòng đầu");
        Eq(P[0].Format.LeftIndent, 0f, "thụt dòng đầu không đổi thụt trái");
        Eq(msg.Contains("đoạn tại con trỏ"), true, "thông báo phạm vi tại con trỏ");
        foreach (int index in new[] { 1, 2, 3, 4 })
        {
            app.Selection.Paragraphs.Items.Clear();
            app.Selection.Paragraphs.Items.Add(P[index]);
            wl.StandardizeParagraphs(st);
            Eq(P[index].Format.Alignment, 3, "căn đều khi con trỏ tại đoạn " + index);
        }
        Eq(P[4].Range.Font.Size, 11f, "font khác giữ nguyên cỡ khi chuẩn hóa đoạn hiện tại");
        Eq(Math.Abs(P[4].Format.FirstLineIndent - 35.43f) < 0.05f, true, "font khác vẫn có thụt đầu dòng");
        Eq(P[1].Format.FirstLineIndent + P[2].Format.FirstLineIndent + P[3].Format.FirstLineIndent, 0f,
            "đầu dòng của tiêu đề, bảng và danh sách được giữ");

        // 2) Thụt dòng đầu tùy chỉnh 0 cm: không thụt
        var app2 = new MLayoutApp();
        app2.ActiveDocument.Sections.Items.Add(new MLayoutSection(1, 0));
        app2.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Times New Roman", 12f, false));
        var st2 = new AppSettings(); st2.ParaFirstLineCm = 0;
        new WordLayout(app2).StandardizeParagraphs(st2);
        Eq(app2.ActiveDocument.Paragraphs.Items[0].Format.FirstLineIndent, 0f, "First line 0 cm: không thụt");

        // 3) Thụt treo 1 cm: LeftIndent 28,35 và FirstLineIndent -28,35
        var app3 = new MLayoutApp();
        app3.ActiveDocument.Sections.Items.Add(new MLayoutSection(1, 0));
        app3.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Times New Roman", 12f, false));
        var st3 = new AppSettings(); st3.ParaHangingCm = 1.0;
        new WordLayout(app3).StandardizeParagraphs(st3);
        var f3 = app3.ActiveDocument.Paragraphs.Items[0].Format;
        Eq(Math.Abs(f3.LeftIndent - 28.35f) < 0.05f && Math.Abs(f3.FirstLineIndent + 28.35f) < 0.05f, true, "Hanging 1 cm: thụt trái 28,35 pt và dòng đầu -28,35 pt");

        // 4) Hanging ưu tiên khi cả hai > 0
        var app4 = new MLayoutApp();
        app4.ActiveDocument.Sections.Items.Add(new MLayoutSection(1, 0));
        app4.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Times New Roman", 12f, false));
        var st4 = new AppSettings(); st4.ParaFirstLineCm = 1.25; st4.ParaHangingCm = 0.5;
        new WordLayout(app4).StandardizeParagraphs(st4);
        var f4 = app4.ActiveDocument.Paragraphs.Items[0].Format;
        Eq(Math.Abs(f4.FirstLineIndent + 14.17f) < 0.05f, true, "cả hai > 0: ưu tiên thụt treo (dòng đầu -14,17 pt)");

        // 5) Cài đặt: giới hạn 0 đến 5 khi đọc từ tệp văn bản
        var st5 = new AppSettings(); st5.ParaFirstLineCm = 1.5; st5.ParaHangingCm = 0.75;
        Eq(st5.ParaFirstLineCm, 1.5, "Settings giữ First line 1,5");
        Eq(st5.ParaHangingCm, 0.75, "Settings giữ Hanging 0,75");
    }


    static void TestProdApplyNow()
    {

        // 15. Rev0.9.1C: Tích theo định dạng (CalcProdFmt)
        {
            var boldFont = new Rec("font"); boldFont.P["Bold"] = -1;
            var stp = new AppSettings(); stp.CalcFmtKind = "bold";

            // 15a. Bảng 1 hàng có các ô số, vài ô in đậm; con trỏ ở cột giữa -> nhận tích đúng (ví dụ 2,5 và 4 và 3 -> 30), không thêm hàng/cột
            var t = new MTable(new string[][] { new string[] { "2,5", "4", "", "3", "99" } });
            t.Cell(1, 1).Range.P["Font"] = boldFont;
            t.Cell(1, 2).Range.P["Font"] = boldFont;
            t.Cell(1, 4).Range.P["Font"] = boldFont;
            // Cột 3 là con trỏ ở giữa; cột 5 ("99") không in đậm
            var app = new MApp(t, new MSel(t, 1, 3, 1, 3));
            var w = new WordTables(app);
            string msg = w.CalcProdFmt(stp);
            Eq(t.D.Count, 1, "CalcProdFmt: bảng 1 hàng không thêm hàng");
            Eq(t.D[0].Count, 5, "CalcProdFmt: không thêm cột");
            Eq(t.D[0][2], "30,0", "con trỏ ở cột giữa nhận tích đúng của các ô đậm (2,5 x 4 x 3 -> 30,0)");
            Eq(msg.Contains("Tích của 3 số"), true, "thông báo tích của 3 số");

            // Kiểm tra thêm với các số nguyên: 2 x 5 x 3 = 30
            var tInt = new MTable(new string[][] { new string[] { "2", "5", "", "3", "99" } });
            tInt.Cell(1, 1).Range.P["Font"] = boldFont;
            tInt.Cell(1, 2).Range.P["Font"] = boldFont;
            tInt.Cell(1, 4).Range.P["Font"] = boldFont;
            var wInt = new WordTables(new MApp(tInt, new MSel(tInt, 1, 3, 1, 3)));
            wInt.CalcProdFmt(stp);
            Eq(tInt.D[0][2], "30", "con trỏ ở cột giữa nhận tích số nguyên 2 x 5 x 3 -> 30");

            // 15b. Số thập phân: 1,5 và 1,5 -> 2,25
            t = new MTable(new string[][] { new string[] { "1,5", "", "1,5" } });
            t.Cell(1, 1).Range.P["Font"] = boldFont;
            t.Cell(1, 3).Range.P["Font"] = boldFont;
            app = new MApp(t, new MSel(t, 1, 2, 1, 2));
            w = new WordTables(app);
            msg = w.CalcProdFmt(stp);
            Eq(t.D[0][1], "2,25", "số thập phân 1,5 và 1,5 -> 2,25");
            Eq(t.D.Count, 1, "số thập phân: không thêm hàng");
            Eq(t.D[0].Count, 3, "số thập phân: không thêm cột");

            // 15c. Số âm
            t = new MTable(new string[][] { new string[] { "-2", "", "3" } });
            t.Cell(1, 1).Range.P["Font"] = boldFont;
            t.Cell(1, 3).Range.P["Font"] = boldFont;
            app = new MApp(t, new MSel(t, 1, 2, 1, 2));
            w = new WordTables(app);
            msg = w.CalcProdFmt(stp);
            Eq(t.D[0][1], "-6", "số âm: -2 x 3 = -6");
            // Số âm có phần thập phân: -1,5 và 2 -> -3,0
            t = new MTable(new string[][] { new string[] { "-1,5", "", "2" } });
            t.Cell(1, 1).Range.P["Font"] = boldFont;
            t.Cell(1, 3).Range.P["Font"] = boldFont;
            app = new MApp(t, new MSel(t, 1, 2, 1, 2));
            w = new WordTables(app);
            w.CalcProdFmt(stp);
            Eq(t.D[0][1], "-3,0", "số âm thập phân: -1,5 x 2 = -3,0");

            // 15d. Ô con trỏ cũng in đậm vẫn bị loại
            t = new MTable(new string[][] { new string[] { "2", "99", "3" } });
            t.Cell(1, 1).Range.P["Font"] = boldFont;
            t.Cell(1, 2).Range.P["Font"] = boldFont; // ô con trỏ (cột 2) cũng in đậm
            t.Cell(1, 3).Range.P["Font"] = boldFont;
            app = new MApp(t, new MSel(t, 1, 2, 1, 2));
            w = new WordTables(app);
            w.CalcProdFmt(stp);
            Eq(t.D[0][1], "6", "ô con trỏ cũng in đậm vẫn bị loại khỏi phép tính (2 x 3 = 6)");

            // 15e. Con trỏ ở cột đầu và cột cuối
            // Cột đầu (cột 1)
            t = new MTable(new string[][] { new string[] { "", "3", "5" } });
            t.Cell(1, 2).Range.P["Font"] = boldFont;
            t.Cell(1, 3).Range.P["Font"] = boldFont;
            app = new MApp(t, new MSel(t, 1, 1, 1, 1));
            w = new WordTables(app);
            w.CalcProdFmt(stp);
            Eq(t.D[0][0], "15", "con trỏ ở cột đầu: nhận tích 3 x 5 = 15");

            // Cột cuối (cột 3)
            t = new MTable(new string[][] { new string[] { "3", "5", "" } });
            t.Cell(1, 1).Range.P["Font"] = boldFont;
            t.Cell(1, 2).Range.P["Font"] = boldFont;
            app = new MApp(t, new MSel(t, 1, 3, 1, 3));
            w = new WordTables(app);
            w.CalcProdFmt(stp);
            Eq(t.D[0][2], "15", "con trỏ ở cột cuối: nhận tích 3 x 5 = 15");

            // 15f. Không ô nào khớp -> bảng không đổi và thông báo chứa 'Không có ô nào khớp'
            t = new MTable(new string[][] { new string[] { "2", "3", "" } });
            app = new MApp(t, new MSel(t, 1, 3, 1, 3));
            w = new WordTables(app);
            msg = w.CalcProdFmt(stp);
            Eq(msg.Contains("Không có ô nào khớp"), true, "không ô nào khớp: thông báo chứa 'Không có ô nào khớp'");
            Eq(Row(t, 1), "2|3|", "không ô nào khớp: bảng không đổi");
            Eq(t.D.Count, 1, "không ô nào khớp: không thêm hàng");
            Eq(t.D[0].Count, 3, "không ô nào khớp: không thêm cột");

            // 15g. Vùng chọn nhiều ô -> thông báo 'đặt con trỏ' và bảng không đổi
            t = new MTable(new string[][] { new string[] { "2", "3", "4" } });
            t.Cell(1, 1).Range.P["Font"] = boldFont;
            t.Cell(1, 2).Range.P["Font"] = boldFont;
            app = new MApp(t, new MSel(t, 1, 2, 1, 3)); // chọn 2 ô (cột 2 đến cột 3)
            w = new WordTables(app);
            msg = w.CalcProdFmt(stp);
            Eq(msg.Contains("đặt con trỏ"), true, "vùng chọn nhiều ô: thông báo yêu cầu đặt con trỏ");
            Eq(Row(t, 1), "2|3|4", "vùng chọn nhiều ô: bảng không đổi");
        }

        // 16. Rev0.9.1C: Áp dụng từng nhóm (ApplyNow) và Bảng này (ApplyToSelected)
        {
            // 16a. group='font': chỉ font và cỡ chữ đổi, màu nền tiêu đề và viền KHÔNG đổi
            Rec.Log.Clear();
            var t = Make();
            var app = new MApp(t, new MSel(t, 2, 2, 2, 2));
            var w = new WordTables(app);
            var stFont = new AppSettings();
            stFont.FontName = "Arial";
            stFont.FontSize = 13;
            stFont.HdrBg = "FF0000";
            stFont.BorderStyle = 1;
            w.ApplyNow(stFont, "font");

            bool fontNameSet = false, fontSizeSet = false;
            bool hdrBgSet = false, borderSet = false;
            foreach (string l in Rec.Log)
            {
                if (l.EndsWith("Font.Name=Arial")) fontNameSet = true;
                if (l.EndsWith("Font.Size=13")) fontSizeSet = true;
                if (l.Contains("BackgroundPatternColor")) hdrBgSet = true;
                if (l.Contains("LineStyle")) borderSet = true;
            }
            Eq(fontNameSet && fontSizeSet, true, "group='font': font và cỡ chữ của bảng đổi");
            Eq(hdrBgSet, false, "group='font': màu nền tiêu đề KHÔNG đổi");
            Eq(borderSet, false, "group='font': viền KHÔNG đổi");

            // 16b. group='hdrBg': chỉ nền dòng tiêu đề đổi
            Rec.Log.Clear();
            t = Make();
            app = new MApp(t, new MSel(t, 2, 2, 2, 2));
            w = new WordTables(app);
            var stHdr = new AppSettings();
            stHdr.HdrBg = "336699";
            stHdr.FontName = "Calibri";
            stHdr.FontSize = 14;
            stHdr.BorderStyle = 2;
            int hdrCol = AppSettings.HexToWord("336699");
            w.ApplyNow(stHdr, "hdrBg");

            int cHdr = 0, cOtherBg = 0;
            bool fontChanged = false, borderChanged = false;
            foreach (string l in Rec.Log)
            {
                if (l == "shading.BackgroundPatternColor=" + hdrCol) cHdr++;
                else if (l.Contains("shading.BackgroundPatternColor")) cOtherBg++;
                if (l.Contains("Font.Name=Calibri") || l.Contains("Font.Size=14")) fontChanged = true;
                if (l.Contains("LineStyle")) borderChanged = true;
            }
            Eq(cHdr, 4, "group='hdrBg': chỉ nền dòng tiêu đề đổi (4 ô hàng 1)");
            Eq(cOtherBg, 0, "group='hdrBg': các ô khác không đổi nền");
            Eq(fontChanged, false, "group='hdrBg': font KHÔNG đổi");
            Eq(borderChanged, false, "group='hdrBg': viền KHÔNG đổi");

            // 16c. group='borderSt': chỉ kiểu nét đổi sang nét đứt
            Rec.Log.Clear();
            t = Make();
            app = new MApp(t, new MSel(t, 2, 2, 2, 2));
            w = new WordTables(app);
            var stBorderSt = new AppSettings();
            stBorderSt.BorderStyle = 1; // nét đứt (LineStyle = 3)
            stBorderSt.BorderColor = "FF0000";
            stBorderSt.BorderPt = 2.25;
            stBorderSt.FontName = "Courier New";
            stBorderSt.FontSize = 16;
            stBorderSt.HdrBg = "00FF00";
            w.ApplyNow(stBorderSt, "borderSt");

            bool dashed = false, fontOrBgChanged = false;
            bool stLineWidthWritten = false, stColorWritten = false;
            foreach (string l in Rec.Log)
            {
                if (l.EndsWith(".LineStyle=3")) dashed = true;
                if (l.Contains("table.Borders") && l.Contains(".LineWidth=")) stLineWidthWritten = true;
                if (l.Contains("table.Borders") && l.Contains(".Color=")) stColorWritten = true;
                if (l.Contains("Font.Name=Courier New") || l.Contains("Font.Size=16") || l.Contains("BackgroundPatternColor")) fontOrBgChanged = true;
            }
            Eq(dashed, true, "group='borderSt': chỉ viền đổi sang nét đứt");
            Eq(stLineWidthWritten, false, "group='borderSt': LineWidth KHÔNG bị ghi");
            Eq(stColorWritten, false, "group='borderSt': Color KHÔNG bị ghi");
            Eq(fontOrBgChanged, false, "group='borderSt': font và màu nền KHÔNG đổi");

            // 16c2. group='borderClr': CHỈ Color các cạnh đổi, còn LineStyle và LineWidth KHÔNG bị ghi
            Rec.Log.Clear();
            t = Make();
            app = new MApp(t, new MSel(t, 2, 2, 2, 2));
            w = new WordTables(app);
            var stBorderClr = new AppSettings();
            stBorderClr.BorderColor = "FF0000";
            stBorderClr.BorderStyle = 1;
            stBorderClr.BorderPt = 2.25;
            stBorderClr.FontName = "Courier New";
            stBorderClr.FontSize = 16;
            stBorderClr.HdrBg = "00FF00";
            int clrTarget = AppSettings.HexToWord("FF0000");
            w.ApplyNow(stBorderClr, "borderClr");

            bool clrColorWritten = false, clrLineStyleWritten = false, clrLineWidthWritten = false, clrFontOrBgChanged = false;
            foreach (string l in Rec.Log)
            {
                if (l.Contains("table.Borders") && l.EndsWith(".Color=" + clrTarget)) clrColorWritten = true;
                if (l.Contains("table.Borders") && l.Contains(".LineStyle=")) clrLineStyleWritten = true;
                if (l.Contains("table.Borders") && l.Contains(".LineWidth=")) clrLineWidthWritten = true;
                if (l.Contains("Font.Name=Courier New") || l.Contains("Font.Size=16") || l.Contains("BackgroundPatternColor")) clrFontOrBgChanged = true;
            }
            Eq(clrColorWritten, true, "group='borderClr': CHỈ Color các cạnh đổi");
            Eq(clrLineStyleWritten, false, "group='borderClr': LineStyle KHÔNG bị ghi");
            Eq(clrLineWidthWritten, false, "group='borderClr': LineWidth KHÔNG bị ghi");
            Eq(clrFontOrBgChanged, false, "group='borderClr': font và màu nền KHÔNG đổi");

            // 16c3. group='borderW': CHỈ LineWidth các cạnh đổi, còn LineStyle và Color KHÔNG bị ghi
            Rec.Log.Clear();
            t = Make();
            app = new MApp(t, new MSel(t, 2, 2, 2, 2));
            w = new WordTables(app);
            var stBorderW = new AppSettings();
            stBorderW.BorderPt = 1.5; // WordMap.LineWidthEnum(1.5) = 12
            stBorderW.BorderStyle = 2;
            stBorderW.BorderColor = "0000FF";
            stBorderW.FontName = "Courier New";
            stBorderW.FontSize = 16;
            stBorderW.HdrBg = "00FF00";
            w.ApplyNow(stBorderW, "borderW");

            bool wWidthWritten = false, wStyleWritten = false, wColorWritten = false, wFontOrBgChanged = false;
            foreach (string l in Rec.Log)
            {
                if (l.Contains("table.Borders") && l.EndsWith(".LineWidth=12")) wWidthWritten = true;
                if (l.Contains("table.Borders") && l.Contains(".LineStyle=")) wStyleWritten = true;
                if (l.Contains("table.Borders") && l.Contains(".Color=")) wColorWritten = true;
                if (l.Contains("Font.Name=Courier New") || l.Contains("Font.Size=16") || l.Contains("BackgroundPatternColor")) wFontOrBgChanged = true;
            }
            Eq(wWidthWritten, true, "group='borderW': CHỈ LineWidth các cạnh đổi");
            Eq(wStyleWritten, false, "group='borderW': LineStyle KHÔNG bị ghi");
            Eq(wColorWritten, false, "group='borderW': Color KHÔNG bị ghi");
            Eq(wFontOrBgChanged, false, "group='borderW': font và màu nền KHÔNG đổi");

            // 16c4. group='border' cũ không còn tác dụng
            Rec.Log.Clear();
            t = Make();
            app = new MApp(t, new MSel(t, 2, 2, 2, 2));
            w = new WordTables(app);
            var stDefunct = new AppSettings();
            stDefunct.BorderStyle = 1;
            stDefunct.BorderColor = "FF0000";
            stDefunct.BorderPt = 1.5;
            w.ApplyNow(stDefunct, "border");
            bool defunctBorderWritten = false;
            foreach (string l in Rec.Log) if (l.Contains("table.Borders")) defunctBorderWritten = true;
            Eq(defunctBorderWritten, false, "nhóm 'border' cũ không còn tác dụng (không ghi gì vào viền)");

            // 16d. group='zebra' với st.Zebra=false thì xóa nền dòng chẵn
            Rec.Log.Clear();
            t = Make();
            app = new MApp(t, new MSel(t, 2, 2, 2, 2));
            w = new WordTables(app);
            var stZeb = new AppSettings();
            stZeb.Zebra = true;
            stZeb.ZebraColor = "80FF80";
            w.ApplyNow(stZeb, "zebra");
            int zebCol = AppSettings.HexToWord("80FF80");
            int cZebBefore = 0;
            foreach (string l in Rec.Log) if (l == "shading.BackgroundPatternColor=" + zebCol) cZebBefore++;
            Eq(cZebBefore, 4, "bật zebra: 4 ô dòng chẵn (hàng 3) nhận màu zebra");

            Rec.Log.Clear();
            stZeb.Zebra = false;
            w.ApplyNow(stZeb, "zebra");
            int cAutoAfter = 0;
            foreach (string l in Rec.Log) if (l == "shading.BackgroundPatternColor=-16777216") cAutoAfter++;
            Eq(cAutoAfter, 4, "group='zebra' với st.Zebra=false: xóa nền dòng chẵn (-16777216)");

            // 16e. selection không trong bảng nào -> thông báo chứa 'Đã lưu thông số' và không ghi gì
            Rec.Log.Clear();
            t = Make();
            var apmOutside = new MApp2(new[] { t }, new MTable[0]);
            var wOutside = new WordTables(apmOutside);
            var stOutside = new AppSettings();
            string msgOutside = wOutside.ApplyNow(stOutside, "font");
            Eq(msgOutside.Contains("Đã lưu thông số"), true, "không trong bảng nào: thông báo chứa 'Đã lưu thông số'");
            int writesOutside = 0;
            foreach (string l in Rec.Log) if (l.Contains("=") && !l.Contains("(")) writesOutside++;
            Eq(writesOutside, 0, "không trong bảng nào: không ghi gì");

            // 16f. ApplyToSelected (Bảng này) vẫn ghi đủ mọi nhóm
            Rec.Log.Clear();
            t = Make();
            app = new MApp(t, new MSel(t, 2, 2, 2, 2));
            w = new WordTables(app);
            var stFull = new AppSettings();
            stFull.FontName = "Segoe UI";
            stFull.FontSize = 10;
            stFull.HdrBg = "156082";
            stFull.BorderStyle = 1;
            stFull.BorderPt = 1.5;
            stFull.BorderColor = "FF0000";
            stFull.Zebra = true;
            stFull.ZebraColor = "80FF80";
            string msgFull = w.ApplyToSelected(stFull);

            bool fullFont = false, fullBorderLs = false, fullBorderLw = false, fullBorderClr = false;
            bool fullHdrBg = false, fullZebra = false, fullStyle = false;
            int hdrFullCol = AppSettings.HexToWord("156082");
            int zebFullCol = AppSettings.HexToWord("80FF80");
            int borderFullCol = AppSettings.HexToWord("FF0000");
            foreach (string l in Rec.Log)
            {
                if (l.EndsWith("Font.Name=Segoe UI") || l.EndsWith("Font.Size=10")) fullFont = true;
                if (l.Contains("table.Borders") && l.EndsWith(".LineStyle=3")) fullBorderLs = true;
                if (l.Contains("table.Borders") && l.EndsWith(".LineWidth=12")) fullBorderLw = true;
                if (l.Contains("table.Borders") && l.EndsWith(".Color=" + borderFullCol)) fullBorderClr = true;
                if (l == "shading.BackgroundPatternColor=" + hdrFullCol) fullHdrBg = true;
                if (l == "shading.BackgroundPatternColor=" + zebFullCol) fullZebra = true;
                if (l.Contains(".Style=TroLy_Master")) fullStyle = true;
            }
            bool fullBorderAll = fullBorderLs && fullBorderLw && fullBorderClr;
            Eq(fullBorderAll, true, "ApplyToSelected (Bảng này) vẫn ghi đủ cả ba (LineStyle, LineWidth, Color)");
            Eq(fullFont && fullBorderAll && fullHdrBg && fullZebra && fullStyle, true, "ApplyToSelected (Bảng này) vẫn ghi đủ mọi nhóm (font, viền, nền tiêu đề, zebra, style)");
        }

        // 17. Rev0.9.1D: Tích theo định dạng cho nhiều hàng (bôi đen nhiều ô cùng một cột)
        {
            WordTables.Tracker.Reset();
            var boldFont = new Rec("font"); boldFont.P["Bold"] = -1;
            var stMulti = new AppSettings();
            stMulti.CalcFmtKind = "bold";
            stMulti.NumThou = false;

            // 17a. Bảng 3 hàng, mỗi hàng có số lượng và đơn giá in đậm, ô thành tiền ở cột cuối; bôi đen cột cuối 3 ô -> mỗi ô nhận đúng tích của hàng mình, không thêm hàng cột
            var tMulti = new MTable(new string[][] {
                new string[] { "2", "1500", "" },
                new string[] { "4", "250", "" },
                new string[] { "3", "10,5", "" }
            });
            tMulti.Cell(1, 1).Range.P["Font"] = boldFont;
            tMulti.Cell(1, 2).Range.P["Font"] = boldFont;
            tMulti.Cell(2, 1).Range.P["Font"] = boldFont;
            tMulti.Cell(2, 2).Range.P["Font"] = boldFont;
            tMulti.Cell(3, 1).Range.P["Font"] = boldFont;
            tMulti.Cell(3, 2).Range.P["Font"] = boldFont;

            var appMulti = new MApp(tMulti, new MSel(tMulti, 1, 3, 3, 3));
            var wMulti = new WordTables(appMulti);
            string msgMulti = wMulti.CalcProdFmt(stMulti);

            Eq(tMulti.D.Count, 3, "Tích nhiều hàng: bảng 3 hàng không thêm hàng");
            Eq(tMulti.D[0].Count, 3, "Tích nhiều hàng: không thêm cột");
            Eq(tMulti.D[0][2], "3000", "hàng 1 nhận đúng tích của hàng mình (2 x 1500 = 3000)");
            Eq(tMulti.D[1][2], "1000", "hàng 2 nhận đúng tích của hàng mình (4 x 250 = 1000)");
            Eq(tMulti.D[2][2], "31,5", "hàng 3 nhận đúng tích của hàng mình (3 x 10,5 = 31,5)");
            Eq(msgMulti.Contains("Đã tính tích cho 3 hàng"), true, "thông báo đã tính tích cho 3 hàng");

            // 17b. Một hàng không có ô đậm: ô đích hàng đó giữ nguyên và thông báo nêu số hàng bỏ qua
            WordTables.Tracker.Reset();
            var tSkip = new MTable(new string[][] {
                new string[] { "2", "1500", "" },
                new string[] { "4", "250", "GiaCu" },
                new string[] { "3", "10,5", "" }
            });
            tSkip.Cell(1, 1).Range.P["Font"] = boldFont;
            tSkip.Cell(1, 2).Range.P["Font"] = boldFont;
            // Hàng 2 không có ô đậm
            tSkip.Cell(3, 1).Range.P["Font"] = boldFont;
            tSkip.Cell(3, 2).Range.P["Font"] = boldFont;

            var appSkip = new MApp(tSkip, new MSel(tSkip, 1, 3, 3, 3));
            var wSkip = new WordTables(appSkip);
            string msgSkip = wSkip.CalcProdFmt(stMulti);

            Eq(tSkip.D[0][2], "3000", "hàng 1: tính đúng 3000");
            Eq(tSkip.D[1][2], "GiaCu", "hàng 2: không có ô đậm thì ô đích hàng đó giữ nguyên");
            Eq(tSkip.D[2][2], "31,5", "hàng 3: tính đúng 31,5");
            Eq(msgSkip.Contains("bỏ qua 1 hàng"), true, "thông báo nêu rõ số hàng bỏ qua");
            Eq(msgSkip.Contains("Đã tính tích cho 2 hàng"), true, "thông báo nêu số hàng đã tính");

            // 17c. Các ô được bôi đen không cùng một cột -> báo lỗi và bảng không đổi
            WordTables.Tracker.Reset();
            var tDiffCols = new MTable(new string[][] {
                new string[] { "2", "1500", "A" },
                new string[] { "4", "250", "B" },
                new string[] { "3", "10,5", "C" }
            });
            tDiffCols.Cell(1, 1).Range.P["Font"] = boldFont;
            tDiffCols.Cell(1, 2).Range.P["Font"] = boldFont;
            tDiffCols.Cell(2, 1).Range.P["Font"] = boldFont;
            tDiffCols.Cell(2, 2).Range.P["Font"] = boldFont;
            tDiffCols.Cell(3, 1).Range.P["Font"] = boldFont;
            tDiffCols.Cell(3, 2).Range.P["Font"] = boldFont;

            var appDiffCols = new MApp(tDiffCols, new MSel(tDiffCols, 1, 2, 2, 3));
            var wDiffCols = new WordTables(appDiffCols);
            string msgDiffCols = wDiffCols.CalcProdFmt(stMulti);

            Eq(msgDiffCols.Contains("cùng một cột"), true, "các ô được bôi đen không cùng một cột: báo lỗi");
            Eq(Row(tDiffCols, 1), "2|1500|A", "không cùng cột: hàng 1 không đổi");
            Eq(Row(tDiffCols, 2), "4|250|B", "không cùng cột: hàng 2 không đổi");
            Eq(Row(tDiffCols, 3), "3|10,5|C", "không cùng cột: hàng 3 không đổi");

            // 17d. Không hàng nào khớp -> bảng không đổi
            WordTables.Tracker.Reset();
            var tNone = new MTable(new string[][] {
                new string[] { "2", "1500", "X" },
                new string[] { "4", "250", "Y" },
                new string[] { "3", "10,5", "Z" }
            });
            var appNone = new MApp(tNone, new MSel(tNone, 1, 3, 3, 3));
            var wNone = new WordTables(appNone);
            string msgNone = wNone.CalcProdFmt(stMulti);

            Eq(msgNone.Contains("Không có ô nào khớp"), true, "không hàng nào khớp: thông báo không có ô khớp");
            Eq(Row(tNone, 1), "2|1500|X", "không hàng nào khớp: hàng 1 không đổi");
            Eq(Row(tNone, 2), "4|250|Y", "không hàng nào khớp: hàng 2 không đổi");
            Eq(Row(tNone, 3), "3|10,5|Z", "không hàng nào khớp: hàng 3 không đổi");
            Eq(tNone.D.Count, 3, "không hàng nào khớp: không thêm hàng");
            Eq(tNone.D[0].Count, 3, "không hàng nào khớp: không thêm cột");

            // 17e. Kết quả hàng này không ảnh hưởng hàng khác: ô thành tiền ở hàng 1 cũng in đậm vẫn không bị nhân vào hàng 2
            WordTables.Tracker.Reset();
            var tCross = new MTable(new string[][] {
                new string[] { "2", "1500", "999" },
                new string[] { "4", "250", "" },
                new string[] { "3", "10,5", "" }
            });
            tCross.Cell(1, 1).Range.P["Font"] = boldFont;
            tCross.Cell(1, 2).Range.P["Font"] = boldFont;
            tCross.Cell(1, 3).Range.P["Font"] = boldFont;
            tCross.Cell(2, 1).Range.P["Font"] = boldFont;
            tCross.Cell(2, 2).Range.P["Font"] = boldFont;
            tCross.Cell(3, 1).Range.P["Font"] = boldFont;
            tCross.Cell(3, 2).Range.P["Font"] = boldFont;

            var appCross = new MApp(tCross, new MSel(tCross, 1, 3, 3, 3));
            var wCross = new WordTables(appCross);
            wCross.CalcProdFmt(stMulti);

            Eq(tCross.D[0][2], "3000", "hàng 1: ô thành tiền ở hàng 1 cũng in đậm vẫn không bị nhân vào chính nó");
            Eq(tCross.D[1][2], "1000", "hàng 2: không bị nhân ô đậm của hàng 1 (kết quả vẫn là 1000)");
            Eq(tCross.D[2][2], "31,5", "hàng 3: không bị ảnh hưởng bởi hàng khác (kết quả vẫn là 31,5)");

            // 18. Rev0.9.3 (A, C): Canh toàn bảng trái/giữa/phải và thêm 0 hàng/cột
            Rec.Log.Clear();
            var tAl = Make();
            var appAl = new MApp(tAl, new MSel(tAl, 2, 2, 2, 2));
            var wAl = new WordTables(appAl);
            string alMsg = wAl.AlignTables(1);
            Eq(alMsg.Contains("giữa"), true, "A: canh toàn bảng giữa");
            Eq(Rec.Log.Contains("rows.Alignment=1"), true, "A: ghi Rows.Alignment=1");
            Rec.Log.Clear(); wAl.AlignTables(2);
            Eq(Rec.Log.Contains("rows.Alignment=2"), true, "A: ghi Rows.Alignment=2 (phải)");
            Rec.Log.Clear(); wAl.AlignTables(0);
            Eq(Rec.Log.Contains("rows.Alignment=0"), true, "A: ghi Rows.Alignment=0 (trái)");
            bool alErr = false;
            try { new WordTables(new MApp2(new[] { tAl }, new MTable[0])).AlignTables(1); } catch (UserMsg) { alErr = true; }
            Eq(alErr, true, "A: ngoài bảng báo UserMsg");
            int rowsBefore = tAl.D.Count, colsBefore = tAl.D[0].Count;
            string add0r = wAl.AddRows(0, false, "");
            string add0c = wAl.AddCols(0, false, "");
            Eq(tAl.D.Count, rowsBefore, "C: AddRows(0) không thêm hàng");
            Eq(tAl.D[0].Count, colsBefore, "C: AddCols(0) không thêm cột");
            Eq(add0r.Contains("0") && add0c.Contains("0"), true, "C: thông báo số lượng 0");

            // 19. Rev0.9.3 (F Dàn trang): chuẩn hóa trang, chuẩn hóa đoạn văn, thêm 1 trang ngang/dọc
            {
                var lapp = new MLayoutApp();
                lapp.ActiveDocument.Sections.Items.Add(new MLayoutSection(1, 0));
                lapp.ActiveDocument.Sections.Items.Add(new MLayoutSection(2, 1));
                var lst = new AppSettings();
                var wl = new WordLayout(lapp);
                string psMsg = wl.StandardizePageSetup(lst);
                Eq(psMsg.Contains("A4"), true, "F: thông báo chuẩn hóa A4");
                var p1 = lapp.ActiveDocument.Sections.Items[0].PageSetup;
                var p2 = lapp.ActiveDocument.Sections.Items[1].PageSetup;
                Eq(p1.PaperSize, 7, "F: section 1 đặt A4 (7)");
                Eq(Math.Abs(p1.LeftMargin - (float)WordLayout.CmToPt(3.0)) < 0.05f, true, "F: trang dọc lề trái 3 cm");
                Eq(Math.Abs(p1.TopMargin - (float)WordLayout.CmToPt(2.0)) < 0.05f, true, "F: trang dọc lề trên 2 cm");
                Eq(Math.Abs(p2.TopMargin - (float)WordLayout.CmToPt(3.0)) < 0.05f, true, "F: trang ngang lề trên 3 cm");
                Eq(Math.Abs(p2.LeftMargin - (float)WordLayout.CmToPt(2.0)) < 0.05f, true, "F: trang ngang lề trái 2 cm");
                Eq(Math.Abs(p1.HeaderDistance - (float)WordLayout.CmToPt(1.0)) < 0.05f, true, "F: header 1 cm");

                // Chuẩn hóa đoạn văn: TNR->13, Arial->11, Calibri->giữ nguyên
                lapp.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Times New Roman", 12f, false));
                lapp.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Arial", 14f, false));
                lapp.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Calibri", 15f, false));
                lapp.ActiveDocument.Paragraphs.Items.Add(new MLayoutPara("Times New Roman", 10f, true)); // trong bảng: bỏ qua
                string pfMsg = wl.StandardizeParagraphs(lst);
                Eq(pfMsg.Contains("đoạn tại con trỏ"), true, "F: chuẩn hóa riêng một đoạn");
                Eq(lapp.ActiveDocument.Paragraphs.Items[1].Range.Font.Size, 14f, "F: đoạn Arial khác giữ nguyên");
                Eq(lapp.ActiveDocument.Paragraphs.Items[2].Format.LineSpacing, 0f, "F: đoạn Calibri khác giữ nguyên");
                for (int index = 1; index <= 2; index++)
                {
                    lapp.Selection.Paragraphs.Items.Clear();
                    lapp.Selection.Paragraphs.Items.Add(lapp.ActiveDocument.Paragraphs.Items[index]);
                    wl.StandardizeParagraphs(lst);
                }
                Eq(lapp.ActiveDocument.Paragraphs.Items[0].Range.Font.Size, 13f, "F: Times New Roman -> 13 pt");
                Eq(lapp.ActiveDocument.Paragraphs.Items[1].Range.Font.Size, 11f, "F: Arial -> 11 pt");
                Eq(lapp.ActiveDocument.Paragraphs.Items[2].Range.Font.Size, 15f, "F: font khác (Calibri) giữ nguyên cỡ 15 pt");
                Eq(lapp.ActiveDocument.Paragraphs.Items[3].Range.Font.Size, 10f, "F: ô trong bảng giữ nguyên 10 pt");
                Eq(lapp.ActiveDocument.Paragraphs.Items[2].Format.LineSpacingRule, 5, "F: giãn dòng bội số");
                Eq(Math.Abs(lapp.ActiveDocument.Paragraphs.Items[2].Format.LineSpacing - 14.4f) < 0.05f, true, "F: giãn dòng 1.2 (14.4 pt)");
                Eq(lapp.ActiveDocument.Paragraphs.Items[2].Format.SpaceBefore, 6f, "F: khoảng cách trên 6 pt");
                Eq(lapp.ActiveDocument.Paragraphs.Items[2].Format.SpaceAfter, 6f, "F: khoảng cách dưới 6 pt");

                // Thêm 1 trang ngang giữa văn bản khổ dọc (trang 1 trên 3 trang)
                var lapp2 = new MLayoutApp();
                lapp2.ActiveDocument.Sections.Items.Add(new MLayoutSection(1, 0));
                lapp2.Selection.CurPage = 1; lapp2.Selection.TotalPages = 3; lapp2.Selection.CurSecIdx = 1;
                var wl2 = new WordLayout(lapp2);
                string landMsg = wl2.InsertOrientedPage(lst, true);
                Eq(landMsg.Contains("khổ ngang"), true, "F: thông báo thêm 1 trang ngang");
                Eq(lapp2.ActiveDocument.Sections.Count, 3, "F: tách thành 3 section (dọc - ngang - dọc)");
                Eq(lapp2.ActiveDocument.Sections.Items[0].PageSetup.Orientation, 0, "F: section đầu vẫn dọc");
                Eq(lapp2.ActiveDocument.Sections.Items[1].PageSetup.Orientation, 1, "F: section mới thêm quay ngang");
                Eq(lapp2.ActiveDocument.Sections.Items[2].PageSetup.Orientation, 0, "F: section ngay dưới giữ nguyên khổ dọc");

                // Thêm 1 trang dọc khi đang ở trang ngang
                var lapp3 = new MLayoutApp();
                lapp3.ActiveDocument.Sections.Items.Add(new MLayoutSection(1, 1));
                lapp3.Selection.CurPage = 1; lapp3.Selection.TotalPages = 2; lapp3.Selection.CurSecIdx = 1;
                var wl3 = new WordLayout(lapp3);
                wl3.InsertOrientedPage(lst, false);
                Eq(lapp3.ActiveDocument.Sections.Count, 3, "F: tách thành 3 section (ngang - dọc - ngang)");
                Eq(lapp3.ActiveDocument.Sections.Items[1].PageSetup.Orientation, 0, "F: section mới thêm khổ dọc");
                Eq(lapp3.ActiveDocument.Sections.Items[2].PageSetup.Orientation, 1, "F: section ngay dưới giữ nguyên khổ ngang");
            }
        }
    }
}
