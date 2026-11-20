// Trợ lý Kỹ thuật - Rev0.4: bộ đổi kiểu số cho mô-đun Định dạng số. Làm việc trên chuỗi bất kỳ (ô bảng hoặc đoạn văn), không gọi Word.
// Hai thao tác: luân chuyển kiểu dấu phân cách (Việt, Mỹ, SI) và tăng, giảm số chữ số thập phân (0 đến 5) như Excel.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    public sealed class NumEditOpts
    {
        public string Style;              // kiểu đích: "vn", "us", "si"; null = giữ kiểu của từng số
        public int DecDelta;              // +1 tăng, -1 giảm một chữ số thập phân; 0 giữ nguyên
        public int SetDec = -1;           // ≥ 0: đồng bộ, mọi số có đúng chừng này chữ số thập phân (làm tròn hoặc thêm số 0)
        public int DecMax = 5;            // tối đa chữ số thập phân khi tăng
        public bool Thousands = true;     // nhóm hàng nghìn khi đổi kiểu
        public string FallbackStyle = "us";   // kiểu hiện hành, dùng để đoán số mơ hồ và kiểu của số nguyên
        public bool Sci;                  // true: viết kiểu kỹ thuật a × 10ⁿ (a có SciDec chữ số thập phân)
        public int SciDec = 2;
    }

    public sealed class NumItem { public int S, E, MS, Exp, UEnd; public string Raw, Tail, UWords; public bool IsSci; }

    public sealed class NumChange { public int S, E; public string Old, New; }

    public sealed class NumEditResult
    {
        public List<NumChange> Changes = new List<NumChange>();
        public int Total, Same, Skipped, Ambiguous;
        public string Text;               // văn bản sau khi sửa
        public char Hint;                 // dấu thập phân đã dùng để đọc số mơ hồ
    }

    public static class NumberEdit
    {
        // Số: có thể có dấu, nhóm nghìn bằng dấu cách không ngắt (hoặc dấu cách thường khi allowSpace), phần thập phân hoặc nhóm bằng . hoặc ,
        const string Body = @"\d+(?:[  ]\d{3}(?![0-9]))*(?:[.,]\d+)*";
        const string BodySp = @"\d+(?:[   ]\d{3}(?![0-9]))*(?:[.,]\d+)*";
        static readonly Regex TokRe = new Regex(@"(?<![A-Za-z0-9_.,])(?:(?<=^|[\s(\[;:=])[-+−](?=\d))?" + Body + @"(?![0-9_])", RegexOptions.Compiled);
        static readonly Regex TokSpRe = new Regex(@"(?<![A-Za-z0-9_.,])(?:(?<=^|[\s(\[;:=])[-+−](?=\d))?" + BodySp + @"(?![0-9_])", RegexOptions.Compiled);

        // Số kiểu kỹ thuật đã có sẵn: 2,3 × 10⁶, 3.67 x 10^3, 1,5·10⁻⁴ (dấu nhân × x X · * ⋅; mũ bằng ký tự Unicode hoặc ^)
        const string SupDigits = "⁰¹²³⁴⁵⁶⁷⁸⁹";
        static readonly Regex SciRe = new Regex(@"(?<![A-Za-z0-9_.,])(?<m>(?:(?<=^|[\s(\[;:=])[-+−](?=\d))?" + Body + @")\s*[×xX·*⋅]\s*10(?<e>[⁻⁺]?[⁰¹²³⁴⁵⁶⁷⁸⁹]+|\^\{?[-+−]?\d+\}?)", RegexOptions.Compiled);

        public static int ParseExp(string e)
        {
            bool neg = false; int v = 0;
            foreach (char c in e)
            {
                if (c == '⁻' || c == '-' || c == '−') neg = true;
                else if (c == '⁺' || c == '+' || c == '^' || c == '{' || c == '}') { }
                else { int i = SupDigits.IndexOf(c); if (i >= 0) v = v * 10 + i; else if (c >= '0' && c <= '9') v = v * 10 + (c - '0'); }
                if (v > 400) break;
            }
            return neg ? -v : v;
        }

        public static string Sup(int e)
        {
            var sb = new StringBuilder();
            if (e < 0) sb.Append('⁻');
            foreach (char c in Math.Abs(e).ToString(CultureInfo.InvariantCulture)) sb.Append(SupDigits[c - '0']);
            return sb.ToString();
        }

        /// <summary>Viết kiểu kỹ thuật: mantissa m (đã có phần nguyên 1 đến 9) × 10 mũ n. m = giá trị, e0 = số mũ sẵn có (0 với số thường). Trả null nếu m = 0.
        /// Ví dụ 2300000 → "2,3 × 10⁶" (dec = 1, kiểu vn); 0,00367 → "3,67 × 10⁻³". Dấu cách quanh × là dấu cách không ngắt để không xuống dòng ở giữa.</summary>
        public static string FormatSci(decimal m, int e0, int dec, string style)
        {
            if (m == 0) return null;
            if (dec < 0) dec = 0; if (dec > 15) dec = 15;
            decimal a = Math.Abs(m); int k = 0;
            while (a >= 10m) { a /= 10m; k++; }
            while (a < 1m) { a *= 10m; k--; }
            decimal r = Math.Round(a, dec, MidpointRounding.AwayFromZero);
            if (r >= 10m) { r /= 10m; k++; r = Math.Round(r, dec, MidpointRounding.AwayFromZero); }
            string mant = NumberUtil.Format(m < 0 ? -r : r, dec, style, false);
            return mant + "\u00A0×\u00A010" + Sup(e0 + k);
        }

        /// <summary>Số kiểu kỹ thuật a × 10ⁿ có sẵn trong chữ: trả danh sách đoạn (vị trí, độ dài, phần a, số mũ).</summary>
        public static List<Match> FindSci(string text)
        {
            var l = new List<Match>();
            foreach (Match m in SciRe.Matches(text)) l.Add(m);
            return l;
        }

        public static string NextStyle(string cur) { return cur == "vn" ? "us" : (cur == "us" ? "si" : "vn"); }     // Việt, Mỹ, SI, rồi quay lại Việt
        public static string StyleName(string s)
        {
            return s == "vn" ? "kiểu Việt (1.234,56)" : (s == "us" ? "kiểu Mỹ (1,234.56)" : "kiểu SI (1 234,56)");
        }
        public static string StyleShort(string s) { return s == "vn" ? "VN" : (s == "us" ? "US" : "SI"); }
        public static string StyleSample(string s) { return s == "vn" ? "1.234,56" : (s == "us" ? "1,234.56" : "1 234,56"); }

        /// <summary>Chữ chỉ gồm các số, dấu phân cách và đơn vị (một ô số, một dãy số được bôi đen): mọi số đều được đổi. Ngược lại chỉ đổi số mang tính kỹ thuật.</summary>
        public static bool IsPure(string text)
        {
            string t = (text ?? "").Replace("\r", " ").Replace("\a", " ").Replace("\n", " ");
            t = SciRe.Replace(t, " ");
            string rest = TokSpRe.Replace(t, " ");
            rest = Regex.Replace(rest, @"[\s;,:()\[\]\-–−/|+=*%°]+", " ").Trim();
            if (rest.Length == 0) return true;
            // còn lại chỉ là đơn vị đã biết thì vẫn coi là thuần số
            foreach (string w in rest.Split(' ')) if (w.Length > 0 && Units.Scan(" " + w) == null) return false;
            return true;
        }

        public static List<Match> Find(string text, bool allowSpace)
        {
            var l = new List<Match>();
            foreach (Match m in (allowSpace ? TokSpRe : TokRe).Matches(text)) l.Add(m);
            return l;
        }

        /// <summary>Các số trong chữ: số thường và số kiểu kỹ thuật có sẵn (a × 10ⁿ coi là một mục, a là phần được định dạng, phần đuôi giữ nguyên).
        /// focusS ≥ 0: chỉ lấy mục chạm vào đoạn [focusS, focusE] (con trỏ nằm trong hoặc sát số).</summary>
        public static List<NumItem> Items(string text, bool allowSpace, int focusS, int focusE)
        {
            var items = new List<NumItem>();
            List<Match> sci = FindSci(text);
            foreach (Match sm in sci)
            {
                Group gm = sm.Groups["m"];
                items.Add(new NumItem { S = sm.Index, E = sm.Index + sm.Length, Raw = gm.Value, Tail = text.Substring(gm.Index + gm.Length - 0, sm.Index + sm.Length - (gm.Index + gm.Length)), Exp = ParseExp(sm.Groups["e"].Value), IsSci = true, MS = gm.Index });
            }
            foreach (Match m in Find(text, allowSpace))
            {
                bool inSci = false;
                foreach (NumItem it in items) if (it.IsSci && m.Index >= it.S && m.Index < it.E) { inSci = true; break; }
                if (!inSci) items.Add(new NumItem { S = m.Index, E = m.Index + m.Length, Raw = m.Value, Tail = "", MS = m.Index });
            }
            items.Sort((x, y) => x.S.CompareTo(y.S));
            if (focusS >= 0) items = items.Where(m => focusS <= m.E && focusE >= m.S).ToList();
            return items;
        }

        /// <summary>Tìm các thay đổi cần làm. pure = true thì mọi số trong chữ đều được đổi; false thì bỏ qua số mang tính văn bản (năm, số trang...).</summary>
        public static NumEditResult Plan(string text, NumEditOpts o, bool pure) { return Plan(text, o, pure, -1, -1); }

        /// <summary>focusS, focusE: chỉ xét số chứa vị trí này (con trỏ nằm trong số), nếu focusS ≥ 0.</summary>
        public static NumEditResult Plan(string text, NumEditOpts o, bool pure, int focusS, int focusE)
        {
            var res = new NumEditResult { Text = text };
            if (string.IsNullOrEmpty(text)) return res;
            List<NumItem> items = Items(text, pure, focusS, focusE);
            // dấu thập phân của vùng: lấy từ các số rõ ràng
            var raws = items.Select(m => m.Raw).ToList();
            char hint = NumberUtil.DetectHint(raws);
            if (hint == '\0') hint = NumberUtil.HintForStyle(o.FallbackStyle);
            res.Hint = hint;
            var sb = new StringBuilder(); int last = 0;
            foreach (NumItem m in items)
            {
                string raw = m.Raw;
                int s = m.S, e = m.E;
                ParsedNumber p0 = NumberUtil.Parse(raw, '\0');
                if (p0 == null) continue;
                bool hasSep = raw.IndexOf('.') >= 0 || raw.IndexOf(',') >= 0 || raw.IndexOf(' ') >= 0 || raw.IndexOf('\u00A0') >= 0 || raw.IndexOf('\u202F') >= 0;
                if (!pure && !m.IsSci)
                {
                    string k = Detector.ClassifyNumber(text, s, e, new NumCtx());
                    if (k == "text") { res.Skipped++; continue; }
                    if (k == "skip" && !hasSep) { res.Skipped++; continue; }
                }
                ParsedNumber p = NumberUtil.Parse(raw, hint);
                if (p == null) continue;
                res.Total++;
                if (p0.Amb) res.Ambiguous++;
                bool spaceGroup = false;
                foreach (char ch in raw) if (ch == ' ' || ch == '\u00A0' || ch == '\u202F') spaceGroup = true;
                bool hadGroup = p.Thou != '\0' || spaceGroup;
                string src = spaceGroup ? "si" : (hasSep ? (p.Dec == ',' || p.Thou == '.' ? "vn" : "us") : null);
                string style = o.Style ?? src ?? o.FallbackStyle;
                string nu; string whole = text.Substring(s, e - s);
                if (o.Sci)
                {
                    nu = FormatSci(p.D, m.Exp, o.SciDec, style);
                    if (nu == null) { res.Same++; continue; }
                    if (raw.Length > 0 && (raw[0] == '+') && !nu.StartsWith("-")) nu = "+" + nu;
                    if (nu == whole) { res.Same++; continue; }
                    res.Changes.Add(new NumChange { S = s, E = e, Old = whole, New = nu });
                    continue;
                }
                int dec = p.Decimals;
                if (o.SetDec >= 0) dec = o.SetDec;
                else if (o.DecDelta > 0) dec = p.Decimals >= o.DecMax ? p.Decimals : p.Decimals + 1;          // quá 5 chữ số thì giữ nguyên, không làm tròn xuống
                else if (o.DecDelta < 0) dec = Math.Max(0, p.Decimals - 1);
                bool thou = o.Style != null ? o.Thousands : hadGroup;
                nu = NumberUtil.Format(p.D, dec, style, thou);
                char sg = raw[0];
                if (sg == '−') { if (nu.StartsWith("-")) nu = "−" + nu.Substring(1); }
                else if (sg == '+' && !nu.StartsWith("-")) nu = "+" + nu;
                nu += m.Tail;
                if (nu == whole) { res.Same++; continue; }
                res.Changes.Add(new NumChange { S = s, E = e, Old = whole, New = nu });
            }
            foreach (NumChange c in res.Changes) { sb.Append(text, last, c.S - last); sb.Append(c.New); last = c.E; }
            sb.Append(text, last, text.Length - last);
            res.Text = sb.ToString();
            return res;
        }
    }
}
