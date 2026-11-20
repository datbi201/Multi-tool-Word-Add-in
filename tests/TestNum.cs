// Ma trận kiểm thử định dạng số: mọi tổ hợp kiểu (us, vn, si) x chữ số thập phân x dấu nghìn x thêm thập phân cho số nguyên.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using TroLyKyThuat;

static class TN
{
    static int fails = 0, total = 0;
    static void Eq(object a, object b, string name)
    {
        total++; bool ok = Equals(a, b); if (!ok) { fails++; if (fails < 40) Console.WriteLine("FAIL " + name + "  (nhan [" + a + "], mong [" + b + "])"); }
    }
    static NumberFormatInfo Nfi(string style, bool thousands)
    {
        var n = new NumberFormatInfo();
        n.NumberDecimalSeparator = style == "us" ? "." : ",";
        n.NumberGroupSeparator = style == "vn" ? "." : (style == "si" ? " " : ",");
        n.NumberGroupSizes = new int[] { 3 };
        n.NegativeSign = "-";
        return n;
    }
    // đáp án độc lập, dựa trên decimal.ToString của .NET (làm tròn xa số 0)
    static string Oracle(decimal v, int dec, string style, bool thousands)
    {
        decimal r = Math.Round(v, dec, MidpointRounding.AwayFromZero);
        string s = r.ToString((thousands ? "N" : "F") + dec, Nfi(style, thousands));
        if (r == 0) s = s.TrimStart('-');
        return s;
    }
    static int Main()
    {
        decimal[] values = { 0m, 0.5m, 7.9m, 12.7m, 999m, 999.995m, 1000m, 1234m, 1234.5m, 5478.8m, 21000m, 12500.5m, 123456.789m, 1234567.891m, -0.001m, -12.5m, -1234.567m, 0.125m, 100m, 62m };
        string[] styles = { "us", "vn", "si" };
        string[] decs = { "keep", "0", "1", "2", "3", "4" };
        int combos = 0;
        foreach (string style in styles)
            foreach (string d in decs)
                foreach (bool thou in new[] { true, false })
                    foreach (bool addDec in new[] { true, false })
                    {
                        foreach (decimal v in values)
                        {
                            combos++;
                            // viết số gốc theo ba kiểu, rồi đưa về kiểu đích
                            int srcDec = (int)((decimal.GetBits(v)[3] >> 16) & 0xFF);
                            foreach (string src in new[] { "us", "vn", "si" })
                            {
                                string text = Oracle(v, srcDec, src, true);
                                char hint = NumberUtil.HintForStyle(style);
                                // với số gốc mơ hồ (3 chữ số sau dấu và phần nguyên 1-3 chữ số) cần bằng chứng: dùng đúng dấu của kiểu nguồn
                                char srcHint = src == "us" ? '.' : ',';
                                string got = NumberUtil.Reformat(text, style, d, thou, addDec, srcHint);
                                int dec = d == "keep" ? srcDec : int.Parse(d);
                                if (d != "keep" && srcDec == 0 && !addDec) dec = 0;
                                string want = Oracle(v, dec, style, thou);
                                Eq(got, want, string.Format("{0}->{1} dec={2} thou={3} add={4} v={5} text=[{6}]", src, style, d, thou, addDec, v, text));
                                // lặp lại phải ra đúng kết quả cũ (idempotent), khi dùng bằng chứng của kiểu đích
                                string again = NumberUtil.Reformat(got, style, d, thou, addDec, NumberUtil.HintForStyle(style));
                                Eq(again, got, string.Format("idempotent {0} dec={1} thou={2} add={3} v={4} got=[{5}]", style, d, thou, addDec, v, got));
                            }
                        }
                    }
        Console.WriteLine("  đã thử " + combos + " tổ hợp");
        // --- đúng các trường hợp trong ảnh người dùng gửi
        var col = new[] { "5478,80", "7,900", "0,000", "21,000", "1234" };
        char ev = NumberUtil.DetectHint(col);
        Eq(ev, ',', "bằng chứng từ 5478,80 là dấu phẩy thập phân");
        foreach (string style in styles)
        {
            var res = col.Select(x => NumberUtil.Reformat(x, style, "2", true, true, ev)).ToArray();
            string ds = style == "us" ? "." : ",";
            string gs = style == "vn" ? "." : (style == "si" ? " " : ",");
            Eq(string.Join("|", res), "5" + gs + "478" + ds + "80|7" + ds + "90|0" + ds + "00|21" + ds + "00|1" + gs + "234" + ds + "00", "ảnh người dùng, kiểu " + style);
        }
        // không có bằng chứng: theo kiểu đích
        Eq(NumberUtil.DetectHint(new[] { "7,900", "21,000", "1234" }), '\0', "không có bằng chứng khi toàn số mơ hồ");
        Eq(NumberUtil.Reformat("7,900", "us", "keep", true, false, NumberUtil.HintForStyle("us")), "7,900", "kiểu Mỹ: 7,900 là bảy nghìn chín trăm");
        Eq(NumberUtil.Reformat("7,900", "vn", "keep", true, false, NumberUtil.HintForStyle("vn")), "7,900", "kiểu Việt: 7,900 là 7,9 giữ ba chữ số thập phân");
        Eq(NumberUtil.Reformat("7,900", "vn", "2", true, false, NumberUtil.HintForStyle("vn")), "7,90", "kiểu Việt, 2 thập phân: 7,90");
        Eq(NumberUtil.Reformat("7,900", "si", "2", true, true, NumberUtil.HintForStyle("si")), "7,90", "SI, 2 thập phân: 7,90");
        Eq(NumberUtil.Reformat("1.234.567", "us", "keep", true, false), "1,234,567", "1.234.567 là hàng nghìn");
        Eq(NumberUtil.Reformat("-0,001", "vn", "2", true, true, ','), "0,00", "số âm làm tròn về 0 không để -0,00");
        Eq(NumberUtil.Reformat("1234,5", "si", "1", true, false, ','), "1 234,5", "SI nhóm nghìn từ 4 chữ số");
        Eq(NumberUtil.Reformat("1 234,5", "si", "1", true, false, ','), "1 234,5", "SI đọc lại được dấu cách không ngắt");
        Eq(NumberUtil.Reformat("1 234,5", "us", "keep", true, false, ','), "1,234.5", "đổi SI sang Mỹ");
        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}
