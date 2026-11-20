// Kiểm thử Rev0.7: đọc số bằng chữ tiếng Việt và viết số kiểu kỹ thuật (a × 10ⁿ).
using System;
using System.Collections.Generic;
using TroLyKyThuat;

static class TNW
{
    static int fails = 0, total = 0;
    static void Eq(object a, object b, string name)
    {
        total++; bool ok = Equals(a, b); if (!ok) { fails++; if (fails < 60) Console.WriteLine("FAIL " + name + "  (nhan [" + a + "], mong [" + b + "])"); }
    }
    static string W(string raw) { return NumberWords.Read(raw, ','); }
    static string Sci(string text, int dec, string style)
    {
        var o = new NumEditOpts { Sci = true, SciDec = dec, FallbackStyle = style, Style = null };
        return NumberEdit.Plan(text, o, true).Text;
    }
    static string Dec(string text, int dec, string style)
    {
        var o = new NumEditOpts { SetDec = dec, FallbackStyle = style };
        return NumberEdit.Plan(text, o, true).Text;
    }

    static int Main()
    {
        // --- đọc số: hàng đơn vị, chục, trăm
        Eq(W("0"), "không", "0");
        Eq(W("5"), "năm", "5");
        Eq(W("10"), "mười", "10");
        Eq(W("11"), "mười một", "11");
        Eq(W("15"), "mười lăm", "15");
        Eq(W("20"), "hai mươi", "20");
        Eq(W("21"), "hai mươi mốt", "21");
        Eq(W("24"), "hai mươi bốn", "24");
        Eq(W("25"), "hai mươi lăm", "25");
        Eq(W("99"), "chín mươi chín", "99");
        Eq(W("100"), "một trăm", "100");
        Eq(W("101"), "một trăm linh một", "101");
        Eq(W("105"), "một trăm linh năm", "105");
        Eq(W("110"), "một trăm mười", "110");
        Eq(W("115"), "một trăm mười lăm", "115");
        Eq(W("121"), "một trăm hai mươi mốt", "121");
        Eq(W("999"), "chín trăm chín mươi chín", "999");
        // --- hàng nghìn, triệu, tỷ
        Eq(W("1000"), "một nghìn", "1000");
        Eq(W("1.234"), "một nghìn hai trăm ba mươi bốn", "1.234 (nhóm nghìn kiểu Việt)");
        Eq(W("1.005"), "một nghìn không trăm linh năm", "1.005");
        Eq(W("1.050"), "một nghìn không trăm năm mươi", "1.050");
        Eq(W("2.015"), "hai nghìn không trăm mười lăm", "2.015");
        Eq(W("10.001"), "mười nghìn không trăm linh một", "10.001");
        Eq(W("100.000"), "một trăm nghìn", "100.000");
        Eq(W("1.000.000"), "một triệu", "1 triệu");
        Eq(W("1.000.005"), "một triệu không trăm linh năm", "1.000.005");
        Eq(W("1.001.000"), "một triệu không trăm linh một nghìn", "1.001.000");
        Eq(W("12.345.678"), "mười hai triệu ba trăm bốn mươi lăm nghìn sáu trăm bảy mươi tám", "12.345.678");
        Eq(W("1.000.000.000"), "một tỷ", "1 tỷ");
        Eq(W("1.000.000.005"), "một tỷ không trăm linh năm", "1 tỷ lẻ 5");
        Eq(W("2.500.000.000"), "hai tỷ năm trăm triệu", "2,5 tỷ");
        Eq(W("1.000.000.000.000"), "một nghìn tỷ", "1 nghìn tỷ");
        Eq(W("1.001.000.000.000"), "một nghìn không trăm linh một tỷ", "1001 tỷ");
        Eq(W("1.000.000.000.000.000.000"), "một tỷ tỷ", "10^18");
        Eq(NumberWords.Integer("1" + new string('0', 31)), null, "quá 30 chữ số thì không đọc");
        // --- thập phân và dấu
        Eq(W("3,14"), "ba phẩy một bốn", "3,14");
        Eq(W("3,05"), "ba phẩy không năm", "3,05");
        Eq(W("0,5"), "không phẩy năm", "0,5");
        Eq(W("12,50"), "mười hai phẩy năm không", "12,50 giữ số 0 cuối");
        Eq(W("1.234,56"), "một nghìn hai trăm ba mươi bốn phẩy năm sáu", "1.234,56");
        Eq(NumberWords.Read("1,234.56", '.'), "một nghìn hai trăm ba mươi bốn phẩy năm sáu", "1,234.56 kiểu Mỹ");
        Eq(NumberWords.Read("1.234", '.'), "một phẩy hai ba bốn", "1.234 với gợi ý dấu thập phân '.' thì đọc là 1,234");
        Eq(NumberWords.Read("1 234,5", ','), "một nghìn hai trăm ba mươi bốn phẩy năm", "SI dấu cách không ngắt");
        Eq(NumberWords.Read("1.234", ','), "một nghìn hai trăm ba mươi bốn", "1.234 với gợi ý dấu thập phân ',' thì '.' là dấu nghìn");
        Eq(W("-7"), "âm bảy", "âm");
        Eq(W("−12,5"), "âm mười hai phẩy năm", "dấu trừ Unicode");
        Eq(W("-0"), "không", "âm không là không");
        Eq(W("abc"), null, "không phải số");
        // --- số kỹ thuật
        Eq(W("2,3 × 10⁶"), "hai phẩy ba nhân mười mũ sáu", "2,3 × 10⁶");
        Eq(W("3.67 x 10^3"), "ba phẩy sáu bảy nhân mười mũ ba", "3.67 x 10^3");
        Eq(W("1,5 × 10⁻⁴"), "một phẩy năm nhân mười mũ âm bốn", "mũ âm");

        // --- viết kiểu kỹ thuật
        Eq(Sci("2300000", 1, "vn"), "2,3 × 10⁶", "2300000 → 2,3 × 10⁶");
        Eq(Sci("3670", 2, "us"), "3.67 × 10³", "3670 → 3.67 × 10³ kiểu Mỹ");
        Eq(Sci("0,00367", 2, "vn"), "3,67 × 10⁻³", "0,00367 → 3,67 × 10⁻³");
        Eq(Sci("12", 0, "vn"), "1 × 10¹", "12, 0 số lẻ");
        Eq(Sci("9,96", 1, "vn"), "1,0 × 10¹", "làm tròn vượt 10 thì chuẩn hóa lại");
        Eq(Sci("-4500", 1, "vn"), "-4,5 × 10³", "số âm");
        Eq(Sci("1", 2, "vn"), "1,00 × 10⁰", "số 1");
        Eq(Sci("0", 2, "vn"), "0", "số 0 giữ nguyên");
        Eq(Sci("1.234.567,8", 3, "vn"), "1,235 × 10⁶", "nhóm nghìn kiểu Việt");
        Eq(Sci("Áp suất 250 bar và 0,5 MPa", 1, "vn"), "Áp suất 2,5 × 10² bar và 5,0 × 10⁻¹ MPa", "nhiều số trong chữ");
        // chạy lần hai không lồng thêm: số kiểu kỹ thuật có sẵn được viết lại, không thêm × 10 lần nữa
        Eq(Sci("2,3 × 10⁶", 1, "vn"), "2,3 × 10⁶", "chạy lại cùng số lẻ không đổi");
        Eq(Sci("2,3 × 10⁶", 3, "vn"), "2,300 × 10⁶", "đổi số lẻ của phần a");
        Eq(Sci("23 × 10^5", 1, "vn"), "2,3 × 10⁶", "phần a ngoài 1..10 được chuẩn hóa: 23 × 10^5 = 2,3 × 10⁶");
        Eq(Sci("1,5 x 10^-4", 2, "vn"), "1,50 × 10⁻⁴", "x và ^ được nhận");
        // đồng bộ số lẻ thường không đụng vào cơ số 10 của số kỹ thuật
        Eq(Dec("2,3 × 10⁶", 2, "vn"), "2,30 × 10⁶", "đồng bộ 2 số lẻ chỉ sửa phần a");
        Eq(Dec("3 x 10^4", 1, "vn"), "3,0 x 10^4", "giữ nguyên phần đuôi x 10^4");
        // đồng bộ 0 đến 5 chữ số thập phân
        for (int d = 0; d <= 5; d++)
        {
            string[] ex = { "3", "3,1", "3,14", "3,142", "3,1416", "3,14159" }; string expect = ex[d];
            Eq(Dec("3,14159", d, "vn"), expect, "đồng bộ " + d + " số lẻ");
        }
        Eq(Dec("2,5", 0, "vn"), "3", "0 số lẻ làm tròn nửa lên: 2,5 → 3");
        Eq(Dec("12 và 3,5", 2, "vn"), "12,00 và 3,50", "đồng bộ thêm số 0");
        Eq(Dec("1,23456789", 5, "vn"), "1,23457", "5 số lẻ");
        // === Rev0.9: đọc đơn vị theo Nghị định 134/2007/NĐ-CP (law) và tên thông dụng (common)
        Func<string, bool, string> U = delegate(string t, bool law) { UnitSpan u = UnitWords.Scan(t, 0, law); return u == null ? null : u.Words; };
        Eq(U(" kPa", true), "kilôpascan", "kPa luật"); Eq(U(" kPa", false), "kilopascal", "kPa thông dụng");
        Eq(U(" N", true), "niutơn", "N luật"); Eq(U(" N", false), "newton", "N thông dụng");
        Eq(U(" MPa", true), "megapascan", "MPa"); Eq(U(" GPa", true), "gigapascan", "GPa");
        Eq(U(" kg", true), "kilôgam", "kg"); Eq(U(" mg", true), "miligam", "mg"); Eq(U(" µm", true), "micrômét", "µm");
        Eq(U(" W", true), "oát", "W"); Eq(U(" kW", true), "kilôoát", "kW"); Eq(U(" V", true), "vôn", "V"); Eq(U(" Hz", true), "héc", "Hz");
        Eq(U(" J", true), "jun", "J"); Eq(U(" Pa", true), "pascan", "Pa"); Eq(U(" °C", true), "độ Celsius", "°C"); Eq(U(" K", true), "kenvin", "K");
        Eq(U(" m", true), "mét", "m"); Eq(U(" mm", true), "milimét", "mm"); Eq(U(" km", true), "kilômét", "km"); Eq(U(" s", true), "giây", "s");
        Eq(U(" m²", true), "mét vuông", "m²"); Eq(U(" m³", true), "mét khối", "m³"); Eq(U(" s²", true), "giây bình phương", "s²");
        Eq(U(" m/s", true), "mét trên giây", "m/s"); Eq(U(" m/s²", true), "mét trên giây bình phương", "m/s²");
        Eq(U(" kg/m³", true), "kilôgam trên mét khối", "kg/m³"); Eq(U(" N·m", true), "niutơn mét", "N·m");
        Eq(U(" W/(m·K)", true), "oát trên mét kenvin", "W/(m·K)"); Eq(U(" %", true), "phần trăm", "%");
        Eq(U("%", true), "phần trăm", "% liền số"); Eq(U(" VNĐ", true), "Việt Nam đồng", "VNĐ"); Eq(U(" đồng", true), "đồng", "đồng");
        Eq(U(" USD", true), "đô la Mỹ", "USD"); Eq(U(" bar", true), "bar", "bar"); Eq(U(" psi", true), "pao trên inch vuông", "psi");
        Eq(U(" người", true), null, "'người' không phải đơn vị"); Eq(U(" cái", true), null, "'cái' không phải đơn vị");
        Eq(U(" kpa", true), null, "kpa viết thường không nhận"); Eq(U(" Phạm", true), null, "Phạm không bị nhận là Pa hay P");
        Eq(U(" mét", true), null, "đã viết chữ thì không đọc lại");
        Func<string, bool, string> R = delegate(string t, bool law)
        {
            var its = NumberEdit.Items(t, true, -1, -1); if (its.Count == 0) return null; int ue; return NumberWords.ReadItem(t, its[0], ',', law, out ue);
        };
        Eq(R("20 kPa", true), "hai mươi kilôpascan", "20 kPa");
        Eq(R("20 kPa", false), "hai mươi kilopascal", "20 kPa thông dụng");
        Eq(R("1.000.000 VNĐ", true), "một triệu Việt Nam đồng", "1.000.000 VNĐ");
        Eq(R("2,5 MPa", true), "hai phẩy năm megapascan", "2,5 MPa");
        Eq(R("-15 °C", true), "âm mười lăm độ Celsius", "-15 °C");
        Eq(R("$100", true), "một trăm đô la Mỹ", "$100");
        Eq(R("100", true), "một trăm", "không đơn vị");
        Eq(R("5.000đ", true), "năm nghìn đồng", "5.000đ"); Eq(R("5.000 đ", true), "năm nghìn đồng", "5.000 đ");
        Eq(R("3 cSt", true), "ba centistốc", "cSt"); Eq(R("10 mD", true), "mười milidarcy", "mD");
        Eq(R("9,81 m/s²", true), "chín phẩy tám một mét trên giây bình phương", "9,81 m/s²");
        Eq(R("2 Nm³/h", true), "hai mét khối chuẩn trên giờ", "Nm³/h");
        Eq(R("25 máy", true), "hai mươi lăm", "chữ thường sau số");
        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}
