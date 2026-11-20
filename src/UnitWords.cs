// Trợ lý Kỹ thuật - Rev0.9: đọc đơn vị đo lường bằng chữ tiếng Việt (cho nút Đọc số): "20 kPa" thành "hai mươi kilôpascan", "1.000.000 VNĐ" thành "một triệu Việt Nam đồng".
// Tên gọi theo Nghị định 134/2007/NĐ-CP về đơn vị đo lường chính thức: kilôgam, niutơn, pascan, jun, oát, culông, vôn, fara, ôm, simen, vebe, becơren, sivơ...
// và tên các ước, bội thập phân: kilô, mega, giga, tera, mili, micrô, nanô, picô... Ghép tiền tố với đơn vị thành một từ (kilôpascan, milimét), mỗi đơn vị chỉ ghép một tiền tố.
// Chế độ "common" dùng tên quốc tế quen đọc (kilopascal, newton, joule, watt). Làm việc trên chuỗi, không gọi Word.
using System;
using System.Collections.Generic;
using System.Text;

namespace TroLyKyThuat
{
    public sealed class UnitSpan { public int Start, End; public string Words; }

    public static class UnitWords
    {
        sealed class U { public string Law, Com; public bool Pre, Len; }
        static readonly Dictionary<string, U> T = new Dictionary<string, U>();
        static readonly Dictionary<string, string[]> Pfx = new Dictionary<string, string[]>();   // ký hiệu tiền tố → { tên theo luật, tên quen dùng }
        static readonly Dictionary<string, string> Cur = new Dictionary<string, string>();       // tiền tệ (không phân biệt hoa thường): khóa viết hoa

        static void A(string sym, string law, string com, bool pre, bool len) { U u = new U(); u.Law = law; u.Com = com; u.Pre = pre; u.Len = len; T[sym] = u; }
        static void A(string sym, string both, bool pre) { A(sym, both, both, pre, false); }
        static void P(string sym, string law, string com) { Pfx[sym] = new[] { law, com }; }

        static UnitWords()
        {
            // --- tiền tố (Nghị định 134/2007/NĐ-CP, bảng 3)
            P("Y", "yôtta", "yotta"); P("Z", "zetta", "zetta"); P("E", "exa", "exa"); P("P", "peta", "peta"); P("T", "tera", "tera"); P("G", "giga", "giga"); P("M", "mega", "mega");
            P("k", "kilô", "kilo"); P("h", "hectô", "hecto"); P("da", "deca", "deca"); P("d", "deci", "deci"); P("c", "centi", "centi"); P("m", "mili", "mili");
            P("µ", "micrô", "micro"); P("μ", "micrô", "micro"); P("n", "nanô", "nano"); P("p", "picô", "pico"); P("f", "femtô", "femto"); P("a", "attô", "atto"); P("z", "zeptô", "zepto"); P("y", "yoctô", "yocto");
            // --- đơn vị cơ bản SI (bảng 1)
            A("m", "mét", "mét", true, true); A("g", "gam", "gam", true, false); A("s", "giây", true); A("A", "ampe", true); A("K", "kenvin", "kelvin", true, false);
            A("mol", "mol", true); A("cd", "candela", true);
            // --- đơn vị dẫn xuất có tên riêng (bảng 2)
            A("Hz", "héc", "hertz", true, false); A("N", "niutơn", "newton", true, false); A("Pa", "pascan", "pascal", true, false); A("J", "jun", "joule", true, false);
            A("W", "oát", "watt", true, false); A("C", "culông", "coulomb", true, false); A("V", "vôn", "volt", true, false); A("F", "fara", "farad", true, false);
            A("Ω", "ôm", "ohm", true, false); A("S", "simen", "siemens", true, false); A("Wb", "vebe", "weber", true, false); A("T", "tesla", true); A("H", "henry", true);
            A("lm", "lumen", true); A("lx", "lux", true); A("Bq", "becơren", "becquerel", true, false); A("Gy", "gray", true); A("Sv", "sivơ", "sievert", true, false); A("kat", "katal", true);
            // --- đơn vị ngoài SI được dùng (bảng 4, bảng 5 và phụ lục)
            A("L", "lít", true); A("l", "lít", true); A("t", "tấn", true); A("min", "phút", false); A("h", "giờ", false); A("d", "ngày", false);
            A("u", "đơn vị khối lượng nguyên tử", false); A("bar", "bar", true); A("eV", "electronvôn", "electronvolt", true, false); A("Wh", "oát giờ", "watt giờ", true, false);
            A("dB", "đềxiben", "đềxiben", false, false); A("B", "ben", true); A("Np", "nepe", false); A("ha", "hécta", "hécta", false, false); A("b", "barn", false);
            A("r", "vòng", false); A("mmHg", "milimét thuỷ ngân", "milimét thủy ngân", false, false); A("cal", "calo", true); A("ct", "carat", false);
            A("°", "độ", false); A("°C", "độ Celsius", "độ C", false, false); A("°F", "độ Fahrenheit", "độ F", false, false); A("℃", "độ Celsius", "độ C", false, false); A("℉", "độ Fahrenheit", "độ F", false, false);
            A("%", "phần trăm", false); A("‰", "phần nghìn", false); A("ppm", "phần triệu", false); A("ppb", "phần tỷ", false); A("cSt", "centistốc", false); A("cP", "centipoa", false); A("mD", "milidarcy", false); A("D", "darcy", false);
            A("wt%", "phần trăm khối lượng", false); A("vol%", "phần trăm thể tích", false); A("mol%", "phần trăm mol", false); A("at%", "phần trăm nguyên tử", false);
            A("rpm", "vòng trên phút", false); A("Å", "ăngstrom", "ăngstrom", false, true);
            // --- ngoài danh mục luật nhưng gặp nhiều trong báo cáo kỹ thuật
            A("atm", "atmôtphe", false); A("psi", "pao trên inch vuông", false); A("psig", "pao trên inch vuông áp kế", false); A("psia", "pao trên inch vuông tuyệt đối", false); A("barg", "bar áp kế", false); A("bara", "bar tuyệt đối", false); A("in", "inch", "inch", false, true); A("ft", "foot", "foot", false, true); A("lb", "pound", false);
            A("gal", "gallon", false); A("bbl", "thùng", "thùng", false, false); A("bpd", "thùng trên ngày", false); A("bopd", "thùng dầu trên ngày", false);
            A("Mbbl", "nghìn thùng", false); A("MMbbl", "triệu thùng", false); A("scf", "bộ khối chuẩn", false); A("Mscf", "nghìn bộ khối chuẩn", false);
            A("MMscf", "triệu bộ khối chuẩn", false); A("MMscfd", "triệu bộ khối chuẩn trên ngày", false); A("Bscf", "tỷ bộ khối chuẩn", false);
            A("Nm", "niutơn mét", "newton mét", false, false); A("Nm3", "mét khối chuẩn", false); A("Nm³", "mét khối chuẩn", false); A("Sm3", "mét khối tiêu chuẩn", false); A("Sm³", "mét khối tiêu chuẩn", false);
            A("VA", "vôn ampe", false); A("pH", "pH", false); A("y", "năm", false); A("yr", "năm", false); A("year", "năm", false); A("mpy", "mil trên năm", false);
            // --- tiền tệ
            Cur["VND"] = "Việt Nam đồng"; Cur["VNĐ"] = "Việt Nam đồng"; Cur["Đ"] = "đồng"; Cur["₫"] = "đồng"; Cur["USD"] = "đô la Mỹ"; Cur["EUR"] = "ơ-rô"; Cur["GBP"] = "bảng Anh";
            Cur["JPY"] = "yên Nhật"; Cur["CNY"] = "nhân dân tệ"; Cur["KRW"] = "won Hàn Quốc"; Cur["AUD"] = "đô la Úc"; Cur["SGD"] = "đô la Singapore";
        }

        static bool IsSym(char c) { return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == 'µ' || c == 'μ' || c == 'Ω' || c == 'Å' || c == 'Đ' || c == 'đ'; }
        const string SupDigits = "⁰¹²³⁴⁵⁶⁷⁸⁹";

        sealed class Term { public string Words; public bool Over; }

        /// <summary>Tên đọc của một ký hiệu (không có số mũ): "kPa" thành "kilôpascan". Trả null nếu không nhận ra. len = đơn vị chiều dài (để đọc vuông, khối).</summary>
        public static string Name(string sym, bool law, out bool len)
        {
            len = false; U u;
            if (T.TryGetValue(sym, out u)) { len = u.Len; return law ? u.Law : u.Com; }
            string up = sym.ToUpperInvariant(); string c;
            if (Cur.TryGetValue(up, out c) || Cur.TryGetValue(sym, out c)) return c;
            // tiền tố + đơn vị: thử tiền tố dài trước (da trước d)
            for (int pl = Math.Min(2, sym.Length - 1); pl >= 1; pl--)
            {
                string[] pf; if (!Pfx.TryGetValue(sym.Substring(0, pl), out pf)) continue;
                if (!T.TryGetValue(sym.Substring(pl), out u) || !u.Pre) continue;
                len = u.Len;
                return (law ? pf[0] : pf[1]) + (law ? u.Law : u.Com);
            }
            return null;
        }

        static int ParseExp(string s)
        {
            bool neg = false; int v = 0;
            foreach (char c in s)
            {
                if (c == '-' || c == '−' || c == '⁻') neg = true;
                else if (c >= '0' && c <= '9') v = v * 10 + (c - '0');
                else { int i = SupDigits.IndexOf(c); if (i >= 0) v = v * 10 + i; }
                if (v > 99) break;
            }
            return neg ? -v : v;
        }

        // Đọc một số hạng: ký hiệu và số mũ tùy chọn. Trả null (i không đổi) nếu không phải đơn vị.
        static Term ParseTerm(string text, ref int i, bool law)
        {
            int n = text.Length, j = i;
            if (j >= n) return null;
            string sym; string words; bool len;
            if (text[j] == '%' || text[j] == '‰' || text[j] == '₫' || text[j] == '℃' || text[j] == '℉') { sym = text[j].ToString(); j++; }
            else
            {
                int st = j;
                if (text[j] == '°') j++;
                while (j < n && IsSym(text[j])) j++;
                if (j < n && text[j] == '%') j++;                    // wt%, vol%
                if (j == st) return null;
                sym = text.Substring(st, j - st);
                if (j < n && char.IsLetter(text[j])) return null;    // phải là trọn một từ: "Ph" trong "Phạm" không phải đơn vị
            }
            int expVal = 1; int k = j;
            if (k < n)
            {
                int e0 = k;
                if (text[k] == '^') { k++; if (k < n && text[k] == '{') k++; int s1 = k; if (k < n && (text[k] == '-' || text[k] == '−' || text[k] == '+')) k++; while (k < n && char.IsDigit(text[k])) k++; if (k > s1 && k < n && text[k] == '}') { expVal = ParseExp(text.Substring(s1, k - s1)); k++; } else if (k > s1) expVal = ParseExp(text.Substring(s1, k - s1)); else k = e0; }
                else if (SupDigits.IndexOf(text[k]) >= 0 || ((text[k] == '⁻' || text[k] == '⁺') && k + 1 < n && SupDigits.IndexOf(text[k + 1]) >= 0))
                {
                    int s1 = k; if (text[k] == '⁻' || text[k] == '⁺') k++;
                    while (k < n && SupDigits.IndexOf(text[k]) >= 0) k++;
                    expVal = ParseExp(text.Substring(s1, k - s1));
                    string comb2 = sym + text.Substring(s1, k - s1);
                    if (T.ContainsKey(comb2)) { sym = comb2; expVal = 1; }                    // Nm³, Sm³
                }
                else if (char.IsDigit(text[k]) || ((text[k] == '-' || text[k] == '−') && k + 1 < n && char.IsDigit(text[k + 1])))
                {
                    // số mũ viết liền: m3, s-1. Nhận một chữ số và không để dính sang số khác ("m3,5" không phải m³)
                    int s1 = k; if (!char.IsDigit(text[k])) k++;
                    k++;
                    bool glued = k < n && (char.IsLetterOrDigit(text[k]) || ((text[k] == ',' || text[k] == '.') && k + 1 < n && char.IsDigit(text[k + 1])));
                    string comb = sym + text.Substring(s1, k - s1);
                    if (!glued && T.ContainsKey(comb)) { sym = comb; expVal = 1; }              // Nm3, Sm3
                    else if (!glued) expVal = ParseExp(text.Substring(s1, k - s1));
                    else k = e0;
                }
            }
            words = Name(sym, law, out len);
            if (words == null) return null;
            var t = new Term(); i = k;
            if (expVal < 0) { t.Over = true; expVal = -expVal; }
            if (expVal == 0) expVal = 1;
            if (expVal == 2) words += len ? " vuông" : " bình phương";
            else if (expVal == 3) words += len ? " khối" : " lập phương";
            else if (expVal >= 4) words += " mũ " + NumberWords.Integer(expVal.ToString());
            t.Words = words; return t;
        }

        /// <summary>Tìm đơn vị đứng ngay sau một số, tại vị trí pos (cho phép một dấu cách). Nhận ký hiệu đơn giản (kPa, °C, %), có số mũ (m², s⁻¹, m^3), đơn vị ghép (N·m, kg/m³, W/(m·K), mg/L) và tiền tệ (VNĐ).
        /// Trả null nếu sau số không phải đơn vị. law = true: tên gọi theo Nghị định 134/2007/NĐ-CP.</summary>
        public static UnitSpan Scan(string text, int pos, bool law)
        {
            if (text == null || pos < 0 || pos >= text.Length) return null;
            int i = pos;
            if (text[i] == ' ' || text[i] == ' ' || text[i] == ' ') i++;
            int start = i;
            if (string.Compare(text, i, "đồng", 0, 4, StringComparison.OrdinalIgnoreCase) == 0 && (i + 4 >= text.Length || !char.IsLetter(text[i + 4])))
            { var dg = new UnitSpan(); dg.Start = start; dg.End = i + 4; dg.Words = "đồng"; return dg; }
            Term first = ParseTerm(text, ref i, law);
            if (first == null) return null;
            var terms = new List<Term>(); terms.Add(first);
            while (i < text.Length)
            {
                char c = text[i]; int save = i;
                if (c == '/')
                {
                    i++;
                    if (i < text.Length && text[i] == '(')
                    {
                        i++; int cnt = 0; bool ok = false;
                        while (true)
                        {
                            Term t = ParseTerm(text, ref i, law);
                            if (t == null) break;
                            if (cnt == 0) t.Over = true; cnt++; terms.Add(t);
                            if (i < text.Length && (text[i] == '·' || text[i] == '⋅' || text[i] == '*' || text[i] == '.')) { i++; continue; }
                            if (i < text.Length && text[i] == ')') { i++; ok = true; }
                            break;
                        }
                        if (!ok) { i = save; while (cnt-- > 0) terms.RemoveAt(terms.Count - 1); break; }
                    }
                    else
                    {
                        Term t = ParseTerm(text, ref i, law);
                        if (t == null) { i = save; break; }
                        t.Over = true; terms.Add(t);
                    }
                }
                else if (c == '·' || c == '⋅' || c == '*' || (c == '.' && i + 1 < text.Length && IsSym(text[i + 1])))
                {
                    i++;
                    Term t = ParseTerm(text, ref i, law);
                    if (t == null) { i = save; break; }
                    terms.Add(t);
                }
                else break;
            }
            var sb = new StringBuilder();
            foreach (Term t in terms) { if (sb.Length > 0) sb.Append(' '); if (t.Over) sb.Append("trên "); sb.Append(t.Words); }
            string w = sb.ToString();
            var u = new UnitSpan(); u.Start = start; u.End = i; u.Words = w; return u;
        }

        /// <summary>Ký hiệu tiền tệ đứng trước số: $100, US$100, €50, £20. Trả tên đọc hoặc null.</summary>
        public static string Prefix(string text, int numStart)
        {
            if (text == null || numStart <= 0) return null;
            int i = numStart - 1;
            if (i >= 0 && (text[i] == '-' || text[i] == '−' || text[i] == '+')) i--;
            if (i < 0) return null;
            if (text[i] == '$') return "đô la Mỹ";
            if (text[i] == '€') return "ơ-rô";
            if (text[i] == '£') return "bảng Anh";
            return null;
        }
    }
}
