// Trợ lý Kỹ thuật - Rev0.1
// Lõi tính toán thuần (không phụ thuộc Word): đọc số, định dạng số, tính tổng/trung bình, ánh xạ màu và viền.
// Chỉ dùng cú pháp C# 5 để biên dịch được bằng csc.exe có sẵn trong Windows.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    public sealed class ParsedNumber
    {
        public double V;
        public decimal D;       // giá trị chính xác dùng để làm tròn và định dạng
        public int Decimals;
        public char Dec;        // '.' hoặc ','
        public char Thou;       // '\0' nếu không có dấu nghìn
        public bool Amb;        // mơ hồ kiểu 1,234
        public string Style;    // "us" hoặc "vn"
        public string Norm;     // chữ số không dấu, thập phân là '.', để đọc số bằng chữ
    }

    public static class NumberUtil
    {
        static readonly Regex NumRe = new Regex(@"^[-+]?\d[\d.,]*$", RegexOptions.Compiled);
        static readonly Regex SpaceRe = new Regex(@"^[-+\u2212]?\d{1,3}(?:[ \u00A0\u202F]\d{3})+(?:[.,]\d+)?$", RegexOptions.Compiled);

        public static ParsedNumber Parse(string s) { return Parse(s, '\0'); }

        /// <summary>
        /// Đọc số từ chữ trong ô. Trả về null nếu không phải số.
        /// hintDec là dấu thập phân đang dùng trong vùng (từ các số không mơ hồ): dùng để quyết định số mơ hồ như 1,234.
        /// </summary>
        public static ParsedNumber Parse(string s, char hintDec)
        {
            if (s == null) return null;
            s = s.Trim();
            if (s.Length == 0) return null;
            // Dấu cách chỉ được là dấu phân cách nhóm 3 chữ số (kiểu SI: 12 500,5). Còn lại, chữ như "1 2" không phải số.
            if (SpaceRe.IsMatch(s)) { /* hợp lệ, bỏ khoảng trắng bên dưới */ }
            else
            {
                foreach (char ch0 in s) if (char.IsWhiteSpace(ch0)) return null;
            }
            var sb = new StringBuilder();
            foreach (char ch in s)
            {
                if (char.IsWhiteSpace(ch)) continue;      // gồm cả dấu cách thường, không ngắt, hẹp
                sb.Append(ch == '\u2212' ? '-' : ch);
            }
            s = sb.ToString();
            if (!NumRe.IsMatch(s)) return null;
            bool neg = s[0] == '-';
            string x = (s[0] == '-' || s[0] == '+') ? s.Substring(1) : s;
            int lc = x.LastIndexOf(','), ld = x.LastIndexOf('.');
            char dec = '.', thou = '\0';
            bool amb = false;
            if (lc >= 0 && ld >= 0)
            {
                dec = lc > ld ? ',' : '.';
                thou = dec == ',' ? '.' : ',';
            }
            else if (lc >= 0 || ld >= 0)
            {
                char sep = lc >= 0 ? ',' : '.';
                string[] parts = x.Split(sep);
                if (parts.Length > 2)
                {
                    thou = sep;
                    dec = sep == ',' ? '.' : ',';
                }
                else if (parts[1].Length == 3 && parts[0].Length <= 3 && parts[0] != "0")
                {
                    if (hintDec == '\0') { amb = true; dec = sep; }
                    else if (sep == hintDec) dec = sep;                 // dấu này là dấu thập phân của vùng
                    else { thou = sep; dec = sep == ',' ? '.' : ','; }  // dấu này là dấu nghìn
                }
                else dec = sep;
            }
            // Rev0.8: dấu nhóm nghìn phải đúng quy tắc (nhóm đầu 1 đến 3 chữ số, các nhóm sau đúng 3 chữ số), dấu thập phân chỉ xuất hiện một lần ở cuối.
            // Trước đây "48,5,000" bị đọc thành 485000 nên một lần bấm sai đã làm đổi giá trị số. Chuỗi không hợp lệ như vậy không phải số.
            if (thou != '\0')
            {
                string ipart = x, fpart = null;
                if (lc >= 0 && ld >= 0) { int dp = x.LastIndexOf(dec); ipart = x.Substring(0, dp); fpart = x.Substring(dp + 1); }
                if (fpart != null && (fpart.Length == 0 || fpart.IndexOf(',') >= 0 || fpart.IndexOf('.') >= 0)) return null;
                if (ipart.IndexOf(dec) >= 0) return null;
                string[] gs = ipart.Split(thou);
                if (gs.Length < 2 || gs[0].Length < 1 || gs[0].Length > 3) return null;
                for (int gi = 1; gi < gs.Length; gi++) if (gs[gi].Length != 3) return null;
            }
            else if (dec != '.' || lc >= 0 || ld >= 0)
            {
                // một dấu duy nhất, nằm ở giữa các chữ số
                string[] ps = x.Split(dec);
                if (ps.Length != 2 || ps[0].Length == 0 || ps[1].Length == 0) return null;
            }
            string t = thou != '\0' ? x.Replace(thou.ToString(), "") : x;
            t = t.Replace(dec, '.');
            double v; decimal d;
            if (!double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return null;
            if (!decimal.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) d = (decimal)0;
            int di = t.IndexOf('.');
            var p = new ParsedNumber();
            p.V = neg ? -v : v;
            p.D = neg ? -d : d;
            p.Decimals = di >= 0 ? t.Length - di - 1 : 0;
            p.Dec = dec; p.Thou = thou; p.Amb = amb;
            p.Style = dec == ',' ? "vn" : "us";
            p.Norm = t;
            return p;
        }

        /// <summary>
        /// Tìm dấu thập phân đang dùng trong một nhóm ô: lấy phiếu từ các số không mơ hồ (12,5 hoặc 1.234,5 hoặc 1.234.567).
        /// Trả về '.' hoặc ',' hoặc '\0' nếu chưa đủ bằng chứng.
        /// </summary>
        public static char DetectHint(IEnumerable<string> texts)
        {
            int dot = 0, comma = 0;
            foreach (string tx in texts)
            {
                if (tx == null || (tx.IndexOf(',') < 0 && tx.IndexOf('.') < 0)) continue;
                ParsedNumber p = Parse(tx, '\0');
                if (p == null || p.Amb) continue;
                if (p.Dec == ',') comma++; else dot++;
            }
            if (dot > comma) return '.';
            if (comma > dot) return ',';
            return '\0';
        }

        /// <summary>Khi không có bằng chứng từ các số khác, số mơ hồ được hiểu theo kiểu đích.</summary>
        public static char HintForStyle(string style) { return style == "us" ? '.' : ','; }

        public static string Format(double v, int decimals, string style, bool thousands)
        {
            try { return Format((decimal)v, decimals, style, thousands); }
            catch (OverflowException) { return v.ToString("F" + decimals, CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Định dạng số: style = us (1,234.56), vn (1.234,56), si (1 234,56 với dấu cách không ngắt).
        /// Nhóm hàng nghìn từ 4 chữ số trở lên ở mọi kiểu. Làm tròn nửa lên, tính bằng decimal nên không lệch như double.
        /// </summary>
        public static string Format(decimal v, int decimals, string style, bool thousands)
        {
            if (decimals < 0) decimals = 0;
            if (decimals > 15) decimals = 15;
            decimal r = Math.Round(Math.Abs(v), decimals, MidpointRounding.AwayFromZero);
            string s = r.ToString("F" + decimals, CultureInfo.InvariantCulture);
            string ip = s, fp = "";
            int dot = s.IndexOf('.');
            if (dot >= 0) { ip = s.Substring(0, dot); fp = s.Substring(dot + 1); }
            string gs = style == "vn" ? "." : (style == "si" ? "\u00A0" : ",");
            string ds = style == "us" ? "." : ",";
            if (thousands && ip.Length >= 4)
            {
                var sb = new StringBuilder();
                int first = ip.Length % 3;
                if (first > 0) sb.Append(ip.Substring(0, first));
                for (int i = first; i < ip.Length; i += 3)
                {
                    if (sb.Length > 0) sb.Append(gs);
                    sb.Append(ip.Substring(i, 3));
                }
                ip = sb.ToString();
            }
            bool allZero = true;
            foreach (char c in (ip + fp)) { if (c >= '1' && c <= '9') { allZero = false; break; } }
            string sign = (v < 0 && !allZero) ? "-" : "";
            return sign + ip + (fp.Length > 0 ? ds + fp : "");
        }

        /// <summary>Định dạng lại một chuỗi số theo quy ước. decSetting = "keep" hoặc "0".."4".</summary>
        public static string Reformat(string text, string style, string decSetting, bool thousands, bool addDecToInt)
        {
            return Reformat(text, style, decSetting, thousands, addDecToInt, NumberUtil.HintForStyle(style));
        }

        public static string Reformat(string text, string style, string decSetting, bool thousands, bool addDecToInt, char hint)
        {
            ParsedNumber p = Parse(text, hint);
            if (p == null) return text;
            int dec = decSetting == "keep" ? p.Decimals : int.Parse(decSetting, CultureInfo.InvariantCulture);
            if (decSetting != "keep" && p.Decimals == 0 && !addDecToInt) dec = 0;
            return Format(p.D, dec, style, thousands);
        }

        public static string OpLabel(string op)
        {
            switch (op)
            {
                case "sum": return "Tổng";
                case "avg": return "Trung bình";
                case "min": return "Nhỏ nhất";
                case "max": return "Lớn nhất";
                case "prod": return "Tích";
                default: return "Số ô";
            }
        }

        public static double Aggregate(string op, IList<double> nums)
        {
            double s = 0, mn = double.MaxValue, mx = double.MinValue;
            double prod = 1;
            foreach (double d in nums) { s += d; prod *= d; if (d < mn) mn = d; if (d > mx) mx = d; }
            switch (op)
            {
                case "sum": return s;
                case "avg": return s / nums.Count;
                case "min": return mn;
                case "max": return mx;
                case "prod": return prod;
                default: return nums.Count;
            }
        }

        /// <summary>Số chữ số thập phân tự động cho kết quả.</summary>
        public static int AutoDecimals(string op, int maxInputDecimals)
        {
            if (op == "cnt") return 0;
            if (op == "avg") return Math.Min(4, Math.Max(maxInputDecimals, 2));
            if (op == "prod") return Math.Min(6, maxInputDecimals);
            return Math.Min(4, maxInputDecimals);
        }
    }

    /// <summary>Kế hoạch tính toán trên vùng chọn: mỗi việc gồm danh sách ô nguồn và ô đích.</summary>
    public sealed class CalcJob
    {
        public List<int[]> Sources = new List<int[]>();   // {hàng, cột} (đánh số từ 1)
        public int[] Target;                              // {hàng, cột}
    }

    public static class CalcPlanner
    {
        /// <summary>Hướng tự động (Rev0.4, không còn cài đặt tay). Một hàng: "row" (điền cột bên phải). Một cột: "col" (điền hàng bên dưới).
        /// Vùng nhiều hàng nhiều cột: nhiều hàng hơn cột thì tính theo cột, ngược lại tính theo hàng. Tham số dir chỉ để tương thích, bị bỏ qua.</summary>
        public static string ResolveDirection(string dir, int r1, int c1, int r2, int c2)
        {
            int nr = r2 - r1 + 1, nc = c2 - c1 + 1;
            if (nr == 1) return "row";
            if (nc == 1) return "col";
            return nr > nc ? "col" : "row";
        }

        public static List<CalcJob> Plan(string dir, int r1, int c1, int r2, int c2)
        {
            var jobs = new List<CalcJob>();
            if (dir == "row")
            {
                for (int r = r1; r <= r2; r++)
                {
                    var j = new CalcJob();
                    for (int c = c1; c <= c2; c++) j.Sources.Add(new int[] { r, c });
                    j.Target = new int[] { r, c2 + 1 };
                    jobs.Add(j);
                }
            }
            else
            {
                for (int c = c1; c <= c2; c++)
                {
                    var j = new CalcJob();
                    for (int r = r1; r <= r2; r++) j.Sources.Add(new int[] { r, c });
                    j.Target = new int[] { r2 + 1, c };
                    jobs.Add(j);
                }
            }
            return jobs;
        }
    }

    public static class WordMap
    {
        /// <summary>Màu RGB sang số màu của Word (BGR).</summary>
        public static int ColorToWord(int r, int g, int b) { return r + g * 256 + b * 65536; }

        public static readonly double[] BorderPts = { 0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6 };
        public static readonly int[] BorderEnum = { 2, 4, 6, 8, 12, 18, 24, 36, 48 };

        /// <summary>Word chỉ có vài độ dày viền cố định, chọn giá trị gần nhất.</summary>
        public static int NearestIndex(double pt)
        {
            int best = 0;
            double bd = double.MaxValue;
            for (int i = 0; i < BorderPts.Length; i++)
            {
                double d = Math.Abs(BorderPts[i] - pt);
                if (d < bd) { bd = d; best = i; }
            }
            return best;
        }
        public static int LineWidthEnum(double pt) { return BorderEnum[NearestIndex(pt)]; }
        public static double SnapPt(double pt) { return BorderPts[NearestIndex(pt)]; }

        /// <summary>0 = liền, 1 = đứt, 2 = chấm. Giá trị theo WdLineStyle: Single 1, Dot 2, DashSmallGap 3.</summary>
        public static int LineStyleEnum(int style)
        {
            if (style == 1) return 3;   // wdLineStyleDashSmallGap (Rev0.1 dùng nhầm 15 = ThinThickLargeGap, nét đôi)
            if (style == 2) return 2;   // wdLineStyleDot
            return 1;                   // wdLineStyleSingle
        }
    }

    /// <summary>Danh sách số hàng/cột kiểu "2,4,6-8".</summary>
    public static class IndexList
    {
        /// <summary>Trả về danh sách đã sắp xếp, không trùng. Trả về null nếu sai cú pháp hoặc ngoài khoảng 1..max.</summary>
        public static List<int> Parse(string text, int max)
        {
            var set = new SortedDictionary<int, bool>();
            if (text == null) return new List<int>();
            foreach (string part in text.Split(new char[] { ',', ';', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int a, b;
                int dash = part.IndexOf('-');
                if (dash > 0)
                {
                    if (!int.TryParse(part.Substring(0, dash), NumberStyles.None, CultureInfo.InvariantCulture, out a)) return null;
                    if (!int.TryParse(part.Substring(dash + 1), NumberStyles.None, CultureInfo.InvariantCulture, out b)) return null;
                    if (b < a) { int t = a; a = b; b = t; }
                }
                else
                {
                    if (!int.TryParse(part, NumberStyles.None, CultureInfo.InvariantCulture, out a)) return null;
                    b = a;
                }
                if (a < 1 || b > max) return null;
                for (int i = a; i <= b; i++) set[i] = true;
            }
            return new List<int>(set.Keys);
        }

        /// <summary>{3,4,5,8} sang "3-5, 8".</summary>
        public static string Format(IList<int> l)
        {
            var sb = new StringBuilder();
            int i = 0;
            while (i < l.Count)
            {
                int j = i;
                while (j + 1 < l.Count && l[j + 1] == l[j] + 1) j++;
                if (sb.Length > 0) sb.Append(", ");
                if (j - i >= 2) sb.Append(l[i]).Append('-').Append(l[j]);
                else { sb.Append(l[i]); if (j > i) sb.Append(", ").Append(l[j]); }
                i = j + 1;
            }
            return sb.ToString();
        }
    }

    public static class CsvUtil
    {
        public static string Escape(string v)
        {
            if (v == null) return "";
            return (v.IndexOf(',') >= 0 || v.IndexOf('"') >= 0 || v.IndexOf('\n') >= 0 || v.IndexOf('\r') >= 0)
                ? "\"" + v.Replace("\"", "\"\"") + "\"" : v;
        }
        public static string CellText(string raw)
        {
            if (raw == null) return "";
            return raw.Replace("\r\a", "").Replace("\a", "").Replace("\r", " ").Replace("\u000B", " ").Trim();
        }
    }
}
