// Trợ lý Kỹ thuật - Rev0.4: từ điển đơn vị kỹ thuật (rpm, kPa, MPa, m³/h...) và bộ đọc mã LaTeX bị hiển thị nguyên văn (\Delta, \tau_w, P_{CO2}).
// Chỉ làm việc trên chuỗi, không gọi Word, nên kiểm thử được riêng.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    /// <summary>Kết quả đọc một đơn vị đứng ngay sau số.</summary>
    public sealed class UnitHit
    {
        public int Len;                  // độ dài đoạn chữ đơn vị tính cả khoảng trắng đầu (nếu có)
        public int Skip;                 // số ký tự khoảng trắng đứng trước đơn vị
        public string Raw;               // đơn vị như người dùng gõ (không có khoảng trắng đầu)
        public string Canon;             // dạng chuẩn dạng chữ (số mũ bằng ký tự chỉ số trên)
        public List<Run> CanonRuns;      // dạng chuẩn có định dạng (số mũ là chỉ số trên)
        public bool Fixable;             // gõ chưa đúng chuẩn
        public bool Ambiguous;           // có hơn một cách hiểu (MPa hay mPa)
        public bool SiCase;              // sai chữ hoa chữ thường của ký hiệu SI (kpa, KPA)
        public List<string> Alts = new List<string>();   // các cách hiểu khác (chuỗi chuẩn)
        public string Why = "";
    }

    public static class Units
    {
        sealed class U { public string Canon; public bool Exact; public bool Si; public string[] Alt; }
        static readonly Dictionary<string, U> Dict = new Dictionary<string, U>();

        static void Add(bool si, string canon, params string[] alt)
        {
            Dict[canon.ToLowerInvariant()] = new U { Canon = canon, Si = si, Alt = alt };
        }
        static void Exact(string canon) { Dict[canon] = new U { Canon = canon, Exact = true }; }

        static Units()
        {
            // áp suất
            Add(true, "Pa"); Add(true, "kPa"); Add(true, "MPa", "mPa"); Add(true, "GPa"); Add(false, "bar"); Add(false, "mbar");
            Add(false, "barg"); Add(false, "bara"); Add(false, "psi"); Add(false, "psig"); Add(false, "psia"); Add(false, "mmHg"); Add(false, "atm");
            // chiều dài, thể tích, khối lượng
            Add(true, "mm"); Add(true, "cm"); Add(true, "km"); Add(true, "nm"); Add(true, "μm");
            Add(false, "ft"); Add(false, "inch"); Add(false, "mil"); Add(true, "mL"); Add(true, "kg"); Add(true, "mg"); Add(false, "lb");
            // thời gian, tốc độ quay, tần số
            Add(false, "rpm"); Add(false, "min"); Add(false, "hr"); Add(false, "yr"); Add(true, "ms"); Add(true, "Hz"); Add(true, "kHz"); Add(true, "MHz"); Add(true, "GHz");
            // điện, công suất, năng lượng, lực
            Add(true, "kW"); Add(true, "MW", "mW"); Add(true, "GW"); Add(true, "kWh"); Add(true, "MWh", "mWh"); Add(true, "kV"); Add(true, "kN"); Add(true, "MN", "mN");
            Add(true, "kJ"); Add(true, "MJ", "mJ"); Add(true, "GJ"); Add(true, "kA"); Add(false, "kcal");
            // nồng độ, ăn mòn, độ nhớt
            Add(false, "ppm"); Add(false, "ppb"); Add(false, "ppmv"); Add(true, "mol"); Add(true, "mmol"); Add(true, "kmol"); Add(false, "mpy");
            Add(false, "cP"); Add(false, "cSt"); Add(false, "dB");
            // viết tắt có phân biệt hoa thường chặt (một chữ cái hoặc dễ nhầm với từ thường)
            Exact("m"); Exact("s"); Exact("h"); Exact("g"); Exact("L"); Exact("V"); Exact("A"); Exact("W"); Exact("N"); Exact("J"); Exact("K"); Exact("Ω");
            Exact("mA"); Exact("mV");
            // từ chỉ thời gian: chỉ nhận đúng chữ thường, không bắt lỗi hoa thường vì là từ thông thường
            foreach (string w in new string[] { "year", "years", "day", "days", "hour", "hours", "week", "weeks", "month", "months", "sec" }) Exact(w);
        }

        const string Comp = @"(?:°[CF]|℃|%|°|[A-Za-zμµΩ]+)(?:\^\{?[-−]?\d\}?|[²³]|(?<=[A-Za-z])[23](?![0-9]))?";
        static readonly Regex UnitRe = new Regex(@"^([ \u00A0]?)(" + Comp + @"(?:[\/·](?:" + Comp + @"))*)(?![A-Za-z0-9_²³%°])", RegexOptions.Compiled);
        static readonly Regex CompRe = new Regex(@"^(°[CF]|℃|%|°|[A-Za-zμµΩ]+?)(?:\^\{?([-−]?\d)\}?|([²³])|([23]))?$", RegexOptions.Compiled);
        const string SupDigits = "⁰¹²³⁴⁵⁶⁷⁸⁹";

        /// <summary>Đọc đơn vị ngay sau một số. aft bắt đầu đúng chỗ ngay sau số (có thể có một khoảng trắng). Trả về null nếu không phải đơn vị đã biết.</summary>
        public static UnitHit Scan(string aft)
        {
            if (string.IsNullOrEmpty(aft)) return null;
            Match m = UnitRe.Match(aft);
            if (!m.Success) return null;
            string raw = m.Groups[2].Value;
            var hit = new UnitHit { Len = m.Length, Skip = m.Groups[1].Length, Raw = raw };
            string[] parts = Regex.Split(raw, @"(?<sep>[\/·])");
            var runs = new List<Run>(); var canon = new StringBuilder();
            var altSets = new List<List<string>>();            // mỗi phần tử có thể có cách hiểu khác
            bool fix = false, si = false, amb = false;
            string why = "";
            foreach (string part in parts)
            {
                if (part == "/" || part == "·") { Append(runs, canon, part, 0); continue; }
                Match cm = CompRe.Match(part);
                if (!cm.Success) return null;
                string u = cm.Groups[1].Value;
                string exp = cm.Groups[2].Success ? cm.Groups[2].Value.Replace('-', '−') : (cm.Groups[3].Success ? "" + (cm.Groups[3].Value == "²" ? "2" : "3") : (cm.Groups[4].Success ? cm.Groups[4].Value : ""));
                string uCanon = u; U e;
                if (u == "°C" || u == "°F" || u == "%" || u == "°") { }
                else if (u == "um" || u == "µm") { uCanon = "μm"; fix = true; si = true; why += u + " nên viết μm. "; }
                else if (u == "℃") { uCanon = "°C"; fix = true; why += "℃ nên viết °C. "; }
                else if (u == "Ω") { }
                else if (Dict.TryGetValue(u, out e) && e.Exact) { }                           // đúng ký hiệu chặt
                else if (Dict.TryGetValue(u.ToLowerInvariant(), out e) && !e.Exact)
                {
                    if (e.Alt != null && e.Alt.Length > 0)
                    {
                        // MPa hay mPa, MW hay mW: đúng nếu trùng một trong các cách viết chuẩn
                        var all = new List<string>(); all.Add(e.Canon); all.AddRange(e.Alt);
                        if (all.Contains(u)) uCanon = u;
                        else if (u[0] == 'M' || u[0] == 'G' || u[0] == 'K')
                        {
                            uCanon = e.Canon; fix = true; si = true;      // Mpa, MPA: chắc chắn là MPa
                            why += u + " nên viết " + e.Canon + ". ";
                        }
                        else
                        {
                            uCanon = e.Canon; fix = true; amb = true;
                            altSets.Add(all);
                            why += u + " chưa đúng chữ hoa, chữ thường: " + string.Join(" hoặc ", all.ToArray()) + ". ";
                        }
                    }
                    else if (u != e.Canon)
                    {
                        uCanon = e.Canon; fix = true;
                        if (e.Si) si = true;
                        why += u + " nên viết " + e.Canon + ". ";
                    }
                }
                else return null;                                                              // có thành phần lạ: không phải đơn vị
                if (exp.Length > 0)
                {
                    if (!(cm.Groups[3].Success)) { fix = true; why += "Số mũ của " + uCanon + " nên là chỉ số trên. "; }
                    Append(runs, canon, uCanon, 0);
                    string d = exp.Replace('−', '-');
                    bool neg = d.StartsWith("-");
                    string digit = neg ? d.Substring(1) : d;
                    runs.Add(new Run((neg ? "−" : "") + digit, 2));
                    canon.Append(neg ? "⁻" : "").Append(SupDigits[digit[0] - '0']);
                }
                else Append(runs, canon, uCanon, 0);
            }
            hit.Canon = canon.ToString();
            hit.CanonRuns = Merge(runs);
            hit.Fixable = fix; hit.SiCase = si; hit.Ambiguous = amb; hit.Why = why.Trim();
            if (amb)
            {
                // các cách hiểu khác của toàn bộ đơn vị: thay lần lượt phần mơ hồ đầu tiên
                string first = altSets[0][0];
                for (int i = 1; i < altSets[0].Count; i++) hit.Alts.Add(hit.Canon.Replace(first, altSets[0][i]));
            }
            return hit;
        }

        static void Append(List<Run> runs, StringBuilder sb, string t, int kind) { runs.Add(new Run(t, kind)); sb.Append(t); }
        static List<Run> Merge(List<Run> runs)
        {
            var o = new List<Run>();
            foreach (Run r in runs)
            {
                if (o.Count > 0 && o[o.Count - 1].Kind == r.Kind) o[o.Count - 1] = new Run(o[o.Count - 1].T + r.T, r.Kind);
                else o.Add(new Run(r.T, r.Kind));
            }
            return o;
        }

        public static bool IsUnitAfter(string aft) { return Scan(aft) != null; }
    }

    /// <summary>Kết quả đọc một ký hiệu viết kiểu LaTeX hoặc gạch dưới.</summary>
    public sealed class SymHit
    {
        public int Start, End;
        public List<Run> Runs;
        public string Show, Linear, Plain;
        public bool Leak;                // có mã dạng \Delta lộ ra
    }

    /// <summary>Đọc ký hiệu kiểu LaTeX lộ ra trong văn bản: \Delta G_{ads}°, \tau_w, P_{CO_2}, x_i, \pm.</summary>
    public static class Symbols
    {
        static readonly Dictionary<string, string> Map = BuildMap();
        static Dictionary<string, string> BuildMap()
        {
            var d = new Dictionary<string, string>();
            string[] lower = { "alpha", "α", "beta", "β", "gamma", "γ", "delta", "δ", "epsilon", "ε", "varepsilon", "ε", "zeta", "ζ", "eta", "η", "theta", "θ", "vartheta", "ϑ",
                "iota", "ι", "kappa", "κ", "lambda", "λ", "mu", "μ", "nu", "ν", "xi", "ξ", "pi", "π", "rho", "ρ", "sigma", "σ", "tau", "τ", "upsilon", "υ",
                "phi", "φ", "varphi", "φ", "chi", "χ", "psi", "ψ", "omega", "ω", "ell", "ℓ", "hbar", "ħ" };
            string[] upper = { "Gamma", "Γ", "Delta", "Δ", "Theta", "Θ", "Lambda", "Λ", "Xi", "Ξ", "Pi", "Π", "Sigma", "Σ", "Phi", "Φ", "Psi", "Ψ", "Omega", "Ω", "nabla", "∇", "partial", "∂" };
            string[] sym = { "pm", "±", "mp", "∓", "times", "×", "cdot", "·", "div", "÷", "le", "≤", "leq", "≤", "ge", "≥", "geq", "≥", "ne", "≠", "neq", "≠",
                "approx", "≈", "sim", "∼", "propto", "∝", "infty", "∞", "degree", "°", "circ", "°", "rightarrow", "→", "to", "→", "leftarrow", "←",
                "leftrightarrow", "↔", "Rightarrow", "⇒", "Leftrightarrow", "⇔", "rightleftharpoons", "⇌", "sum", "∑", "prod", "∏", "int", "∫", "sqrt", "√",
                "cdots", "⋯", "ldots", "…", "dots", "…", "angstrom", "Å", "AA", "Å", "%", "%", "ll", "≪", "gg", "≫", "equiv", "≡", "in", "∈", "forall", "∀", "ln", "ln", "log", "log", "exp", "exp", "sin", "sin", "cos", "cos", "tan", "tan", "lim", "lim", "max", "max", "min", "min" };
            for (int i = 0; i < lower.Length; i += 2) d[lower[i]] = lower[i + 1];
            for (int i = 0; i < upper.Length; i += 2) d[upper[i]] = upper[i + 1];
            for (int i = 0; i < sym.Length; i += 2) d[sym[i]] = sym[i + 1];
            return d;
        }
        static readonly HashSet<string> LowerGreek = new HashSet<string>(new[] { "alpha", "beta", "gamma", "delta", "epsilon", "varepsilon", "zeta", "eta", "theta", "vartheta", "iota", "kappa", "lambda", "mu", "nu", "xi", "pi", "rho", "sigma", "tau", "upsilon", "phi", "varphi", "chi", "psi", "omega", "ell", "hbar" });
        static readonly HashSet<string> UpperGreek = new HashSet<string>(new[] { "Gamma", "Delta", "Theta", "Lambda", "Xi", "Pi", "Sigma", "Phi", "Psi", "Omega", "nabla", "partial" });
        static readonly HashSet<string> Operators = new HashSet<string>(new[] { "Delta", "nabla", "partial", "delta" });   // đứng trước biến: ΔG, ∂P, δt

        /// <summary>Thay mọi mã LaTeX đã biết trong chuỗi bằng ký tự thật (dùng cho Soạn phương trình). \frac và các mã chưa biết giữ nguyên.</summary>
        public static string ReplaceCommands(string s)
        {
            s = Regex.Replace(s, @"\\frac\s*\{([^{}]*)\}\s*\{([^{}]*)\}", "($1)/($2)");
            s = Regex.Replace(s, @"\\([A-Za-z]+|%)", delegate(Match m)
            {
                string v; return Map.TryGetValue(m.Groups[1].Value, out v) ? v : m.Value;
            });
            return Regex.Replace(s, @"√\s*\{([^{}]*)\}", "√($1)");
        }
        public static bool Known(string cmd) { return Map.ContainsKey(cmd); }
        /// <summary>Tra một mã LaTeX: ký tự thật, và có phải chữ Hy Lạp thường (in nghiêng) không.</summary>
        public static bool TryCmd(string cmd, out string val, out bool italic)
        {
            italic = LowerGreek.Contains(cmd);
            return Map.TryGetValue(cmd, out val);
        }

        /// <summary>Có mã LaTeX đã biết trong chuỗi không.</summary>
        public static bool HasCommand(string s)
        {
            foreach (Match m in Regex.Matches(s, @"\\([A-Za-z]+)")) if (Map.ContainsKey(m.Groups[1].Value)) return true;
            return false;
        }

        // điểm bắt đầu của một ký hiệu: mã lệnh LaTeX, hoặc "chữ_chữ" (chỉ số bắt đầu bằng chữ cái; chỉ số bắt đầu bằng số để dành cho hóa học H_2O, biến A_1)
        public static readonly Regex StartRe = new Regex(@"\\(?:[A-Za-z]+|%)|(?<![\p{L}\p{N}_\\])(?:[A-Z][a-z]|[A-Za-z]|[\u0391-\u03A9\u03B1-\u03C9])_(?=\{?\\?\p{L})", RegexOptions.Compiled);

        sealed class Atom { public string Sym; public bool Greek, Upright, Leak, Operator; }

        static Atom ReadAtom(string t, ref int p)
        {
            if (p >= t.Length) return null;
            if (t[p] == '\\')
            {
                int j = p + 1;
                if (j < t.Length && t[j] == '%') { p = j + 1; return new Atom { Sym = "%", Upright = true, Leak = true }; }
                while (j < t.Length && ((t[j] >= 'a' && t[j] <= 'z') || (t[j] >= 'A' && t[j] <= 'Z'))) j++;
                string cmd = t.Substring(p + 1, j - p - 1), v;
                if (cmd.Length > 0 && !Map.ContainsKey(cmd))
                {
                    // "\cdotcm": mã dính liền đơn vị. Lấy mã đã biết dài nhất mà phần còn lại là một đơn vị trọn vẹn (cm, mm, kg...).
                    for (int L = cmd.Length - 1; L >= 2; L--)
                    {
                        string head = cmd.Substring(0, L), rest = cmd.Substring(L);
                        if (!Map.ContainsKey(head)) continue;
                        UnitHit uh = Units.Scan(" " + rest);
                        if (uh != null && uh.Len == rest.Length + 1) { cmd = head; j = p + 1 + L; break; }
                    }
                }
                if (cmd.Length == 0 || !Map.TryGetValue(cmd, out v)) return null;
                p = j;
                var a = new Atom { Sym = v, Leak = true, Operator = Operators.Contains(cmd) };
                a.Greek = LowerGreek.Contains(cmd) || UpperGreek.Contains(cmd);
                a.Upright = !LowerGreek.Contains(cmd);                 // chữ Hy Lạp hoa, toán tử, ký hiệu: đứng thẳng; chữ Hy Lạp thường: nghiêng
                return a;
            }
            char c = t[p];
            if (c >= 'A' && c <= 'Z' && p + 2 < t.Length && t[p + 1] >= 'a' && t[p + 1] <= 'z' && (t[p + 2] == '_' || t[p + 2] == '^'))
            {
                string w = t.Substring(p, 2); p += 2;                  // Re, Nu, Pr: số đặc trưng không thứ nguyên, đứng thẳng
                return new Atom { Sym = w, Upright = true };
            }
            if (MathParser.IsGreek(c))
            {
                if (p + 1 < t.Length && char.IsLetter(t[p + 1])) return null;
                p++; return new Atom { Sym = c.ToString(), Upright = c <= '\u03A9' };
            }
            if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
            {
                int nx = p + 1;
                bool wordNext = nx < t.Length && char.IsLetter(t[nx]);
                if (wordNext) return null;                              // chữ đứng trong một từ, không phải biến
                p = nx; return new Atom { Sym = c.ToString() };
            }
            return null;
        }

        /// <summary>Đọc một nhóm sau _ hoặc ^: {...} (có thể lồng) hoặc chuỗi chữ số, chữ cái.</summary>
        static string ReadGroup(string t, ref int p, bool isSup)
        {
            if (p >= t.Length) return null;
            if (t[p] == '{')
            {
                int depth = 0, j = p;
                for (; j < t.Length; j++)
                {
                    if (t[j] == '{') depth++;
                    else if (t[j] == '}') { depth--; if (depth == 0) break; }
                }
                if (j >= t.Length) return null;
                string inner = t.Substring(p + 1, j - p - 1); p = j + 1;
                return inner;
            }
            int k = p;
            if (isSup)
            {
                // không có ngoặc: số (có thể có dấu), một chữ cái, hoặc một dấu (*, +, -)
                if (char.IsDigit(t[k])) { while (k < t.Length && char.IsDigit(t[k])) k++; }
                else if ((t[k] == '-' || t[k] == '−' || t[k] == '+') && k + 1 < t.Length && char.IsDigit(t[k + 1])) { k++; while (k < t.Length && char.IsDigit(t[k])) k++; }
                else if (char.IsLetter(t[k]) || t[k] == '*' || t[k] == '+' || t[k] == '-' || t[k] == '−') k++;
            }
            else
            {
                while (k < t.Length && (char.IsLetterOrDigit(t[k]) || t[k] == ',')) k++;
                while (k > p && t[k - 1] == ',') k--;
            }
            if (k == p) return null;
            string g = t.Substring(p, k - p); p = k; return g;
        }

        /// <summary>Làm sạch nội dung chỉ số: bỏ mã \mathrm, \text, thay mã LaTeX, gộp _2 thành 2 (CO_2 thành CO2).</summary>
        static string CleanScript(string c)
        {
            string s = Regex.Replace(c, @"\\(?:mathrm|text|textrm|rm|mathit|textit)\s*\{([^{}]*)\}", "$1");
            s = ReplaceCommands(s);
            s = Regex.Replace(s, @"_\{?([A-Za-z0-9]+)\}?", "$1");
            s = s.Replace("{", "").Replace("}", "").Replace("\\", "").Trim();
            return s;
        }

        static bool IsChemSub(string c)
        {
            return c.Length >= 2 && Regex.IsMatch(c, @"^[A-Z]") && Regex.IsMatch(c, @"^[A-Za-z0-9()]+$") && Chem.ParseF(c) != null;
        }

        static string SubLinear(string c)
        {
            if (IsChemSub(c)) return "_(" + Chem.RunsToLinear(Chem.RenderRuns(c)) + ")";
            return "_(\"" + c + "\")";
        }

        /// <summary>Đọc ký hiệu bắt đầu tại i. Thành công: trả về SymHit với các đoạn chữ có định dạng.</summary>
        public static SymHit TryParse(string t, int i)
        {
            int p = i; var runs = new List<Run>(); var lin = new StringBuilder(); var show = new StringBuilder(); bool leak = false;
            Atom a = ReadAtom(t, ref p);
            if (a == null) return null;
            leak |= a.Leak;
            Action<Atom> put = delegate(Atom x)
            {
                runs.Add(new Run(x.Sym, x.Upright ? 0 : 3)); lin.Append(x.Sym); show.Append(x.Sym);
            };
            // chữ cái Latin đơn: nghiêng; lệnh Hy Lạp thường: nghiêng; Hy Lạp hoa, toán tử, ký hiệu: thẳng
            if (!a.Leak && !a.Upright) { /* biến Latin: nghiêng */ }
            put(a);
            bool hasBase = false;
            if (a.Operator)
            {
                // ΔG, ∂P: toán tử rồi biến. Cho phép một khoảng trắng.
                int q = p; if (q < t.Length && t[q] == ' ') q++;
                Atom b = ReadAtom(t, ref q);
                if (b != null && (!b.Leak || b.Greek)) { p = q; leak |= b.Leak; put(b); hasBase = true; }
            }
            else hasBase = true;
            string sub = null, sup = null;
            for (int guard = 0; guard < 2 && p < t.Length; guard++)
            {
                if (t[p] == '_' && sub == null && hasBase)
                {
                    int q = p + 1; string g = ReadGroup(t, ref q, false);
                    if (g == null) break;
                    string cs = CleanScript(g);
                    if (cs.Length == 0) break;
                    sub = cs; p = q;
                }
                else if (t[p] == '^' && sup == null && hasBase)
                {
                    int q = p + 1; string g = ReadGroup(t, ref q, true);
                    if (g == null) break;
                    string cs = CleanScript(g);
                    if (cs.Length == 0) break;
                    sup = cs; p = q;
                }
                else break;
            }
            bool ring = false;
            if (p < t.Length && t[p] == '°' && hasBase) { ring = true; p++; }
            if (sub != null) { runs.Add(new Run(sub, 1)); lin.Append(SubLinear(sub)); show.Append('_').Append(Detector.ShowText(new List<Run> { new Run(sub, 1) })); }
            if (sup != null)
            {
                string sp = sup.Replace('-', '−');
                runs.Add(new Run(sp, 2)); lin.Append("^(").Append(sp).Append(')'); show.Append('^').Append(sp);
            }
            if (ring) { runs.Add(new Run("°", 0)); lin.Append('°'); show.Append('°'); }
            int end = p;
            // mã $...$ bao quanh
            int st = i;
            if (st > 0 && t[st - 1] == '$' && end < t.Length && t[end] == '$') { st--; end++; }
            string plain = Chem.RunsPlain(runs);
            if (!leak && sub == null && sup == null) return null;          // một chữ cái trần: không phải ký hiệu
            if (!leak && !(sub != null || sup != null)) return null;
            var hit = new SymHit { Start = st, End = end, Runs = Merge(runs), Show = show.ToString(), Linear = lin.ToString(), Plain = plain, Leak = leak };
            return hit;
        }

        static List<Run> Merge(List<Run> runs)
        {
            var o = new List<Run>();
            foreach (Run r in runs)
            {
                if (o.Count > 0 && o[o.Count - 1].Kind == r.Kind && r.Kind == 0) o[o.Count - 1] = new Run(o[o.Count - 1].T + r.T, 0);
                else o.Add(new Run(r.T, r.Kind));
            }
            return o;
        }
    }
}
