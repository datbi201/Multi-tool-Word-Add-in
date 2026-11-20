// Trợ lý Kỹ thuật - Rev0.2: bộ nhận diện của mục Chuẩn hóa (hóa học, toán, ký hiệu theo ngữ cảnh, số kỹ thuật).
// Chuyển từ bản demo JavaScript. Chỉ làm việc trên chuỗi, không gọi Word, nên kiểm thử được riêng.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    public sealed class Fix
    {
        public string Label, Show, Plain, Linear;
        public List<Run> Runs;      // chữ có định dạng (null nếu chỉ thay chữ thuần)
        public bool Sb;             // hiện kèm phần đề xuất trên nút
        public bool IsChem;         // công thức hóa học: mọi chữ cái phải đứng thẳng
        public MNode Node;          // cây công thức (có phân số, căn...) để dựng Equation bằng OMML; null thì dựng từ Runs

        public static Fix FromRuns(string label, List<Run> runs, string linear)
        {
            var f = new Fix();
            f.Label = label; f.Runs = runs; f.Plain = Chem.RunsPlain(runs); f.Linear = linear; f.Show = Detector.ShowText(runs);
            return f;
        }
        public static Fix FromChem(string label, string std)
        {
            List<Run> runs = Chem.RenderRuns(std);
            Fix f = FromRuns(label, runs, Chem.RunsToLinear(runs)); f.IsChem = true; return f;
        }
        public static Fix FromNode(string label, MNode node)
        {
            var f = new Fix(); f.Label = label; f.Node = node; f.Runs = Omml.ToRuns(node); f.Plain = Chem.RunsPlain(f.Runs); f.Linear = f.Plain; f.Show = ShowNode(f.Runs);
            return f;
        }
        public static string ShowRuns(List<Run> runs) { return ShowNode(runs); }
        static string ShowNode(List<Run> runs)
        {
            var sb = new StringBuilder();
            foreach (Run r in runs)
            {
                string one = Detector.ShowText(new List<Run> { r });
                if (r.Kind == 1 && !Regex.IsMatch(r.T, @"^\d+$")) sb.Append('_').Append(one);
                else if (r.Kind == 2 && !Regex.IsMatch(r.T, @"^[-−+]?\d+$")) sb.Append('^').Append(one);
                else sb.Append(one);
            }
            return sb.ToString();
        }
        /// <summary>Cây công thức để dựng Equation: có sẵn, hoặc dựng từ các đoạn chữ.</summary>
        public MNode GetNode() { return Node ?? (Runs != null ? Omml.FromRuns(Runs, IsChem) : null); }
        public static Fix Text(string label, string text, bool sb)
        {
            var f = new Fix(); f.Label = label; f.Plain = text; f.Show = text; f.Sb = sb; return f;
        }
        /// <summary>Có đoạn nào cần định dạng chỉ số trên, dưới hay nghiêng không.</summary>
        public bool NeedsFormat { get { return Node != null || (Runs != null && Runs.Any(r => r.Kind != 0)); } }
    }

    public sealed class Issue
    {
        public string Cat, Sev, Key, Before, Reason, Disp, NumStyle;     // Disp: chữ hiển thị trong danh sách nếu khác Before (ví dụ mô tả dấu cách thừa)
        public List<Fix> Fixes = new List<Fix>();
    }

    public sealed class Hit { public int S, E; public Issue Issue; }

    public sealed class NumCtx { public bool Heading, InTable, NumericCell; public Func<int, int, string> Language; }

    public sealed class ReviewStats { public int AutoText, OkNum; }

    public sealed class ReviewOpts
    {
        public bool Chem = true, Math = true, Sym = true, Num = true, NumTable = true;
        public string NumStyle = "us", NumDec = "keep";
        public bool NumThou = true, NumAddDec = false;
        public ReviewStats Stats = new ReviewStats();
    }

    public static class Detector
    {
        public static string CatLabel(string c)
        {
            switch (c)
            {
                case "chem": return "Hóa học"; case "math": return "Toán"; case "sym": return "Ký hiệu biến"; case "num": return "Số kỹ thuật";
                case "unit": return "Đơn vị"; case "term": return "Thuật ngữ"; case "orth": return "Chính tả"; case "abbr": return "Viết tắt";
                case "long": return "Câu dài"; case "inf": return "Khẩu ngữ"; case "dup": return "Lặp từ"; case "punc": return "Dấu câu";
                case "ai": return "Văn mẫu"; case "conn": return "Từ nối"; case "pron": return "Xưng hô";
                case "gr-spell": return "Chính tả"; case "gr-agree": return "Ngữ pháp"; case "gr-punc": return "Dấu câu";
                case "ai-long": return "Câu dài AI"; case "ai-style": return "Văn phong AI"; case "ai-gram": return "Ngữ pháp AI"; case "ai-term": return "Thuật ngữ AI"; case "ed-style": return "Văn phong"; case "gr-style": return "Văn phong (Anh)"; case "gr-term": return "Thuật ngữ (Anh)";
                default: return c;
            }
        }
        public static string SevLabel(string s)
        {
            switch (s) { case "err": return "Sai"; case "sus": return "Nghi ngờ"; default: return "Chưa chuẩn"; }
        }

        const string SUPS = "⁰¹²³⁴⁵⁶⁷⁸⁹";
        const string SUBS = "₀₁₂₃₄₅₆₇₈₉";

        /// <summary>Chữ hiển thị trong danh sách: dùng ký tự chỉ số Unicode cho dễ đọc.</summary>
        public static string ShowText(List<Run> runs)
        {
            var sb = new StringBuilder();
            foreach (Run r in runs)
                foreach (char c in r.T)
                {
                    if (r.Kind == 1 && c >= '0' && c <= '9') sb.Append(SUBS[c - '0']);
                    else if (r.Kind == 2 && c >= '0' && c <= '9') sb.Append(SUPS[c - '0']);
                    else if (r.Kind == 2 && c == '+') sb.Append('⁺');
                    else if (r.Kind == 2 && c == '−') sb.Append('⁻');
                    else sb.Append(c);
                }
            return sb.ToString();
        }

        static string CtxKey(string txt, int s) { int a = Math.Max(0, s - 14); return txt.Substring(a, s - a); }

        // ------------------------------------------------------------------ toán: chuẩn hóa chuỗi gõ phẳng
        static readonly Regex GreekRe = new Regex(@"(?<![A-Za-z\\])(alpha|beta|gamma|delta|Delta|sigma|mu|pi|rho|tau|theta|lambda)(?![A-Za-z])");
        static readonly Dictionary<string, string> Greek = new Dictionary<string, string> {
            { "alpha", "α" }, { "beta", "β" }, { "gamma", "γ" }, { "delta", "δ" }, { "Delta", "Δ" }, { "sigma", "σ" }, { "mu", "μ" },
            { "pi", "π" }, { "rho", "ρ" }, { "tau", "τ" }, { "theta", "θ" }, { "lambda", "λ" } };

        /// <summary>Chuẩn hóa biểu thức gõ phẳng: chỉ số Unicode, chữ Hy Lạp, toán tử, căn.</summary>
        public static string NormMath(string src)
        {
            string s = src;
            s = Regex.Replace(s, "[⁰¹²³⁴⁵⁶⁷⁸⁹]+", m => "^" + new string(m.Value.Select(c => (char)('0' + SUPS.IndexOf(c))).ToArray()));
            s = Regex.Replace(s, "[₀-₉]+", m => "_" + new string(m.Value.Select(c => (char)('0' + (c - '₀'))).ToArray()));
            s = GreekRe.Replace(s, m => Greek[m.Value]);
            s = s.Replace("*", " · ").Replace("<=", "≤").Replace(">=", "≥").Replace("+-", "±");
            s = Regex.Replace(s, @"(?<!\\)sqrt\(", "√(");
            s = Regex.Replace(s, @"\s{2,}", " ");
            return s.Trim();
        }

        /// <summary>Biểu thức có ^ và _ sang các đoạn chữ: số mũ thành chỉ số trên, chỉ số thành chỉ số dưới.</summary>
        public static List<Run> MathRuns(string src)
        {
            var runs = new List<Run>(); var cur = new StringBuilder();
            Action flush = delegate { if (cur.Length > 0) { AddItalicVars(runs, cur.ToString()); cur.Length = 0; } };
            int i = 0;
            while (i < src.Length)
            {
                char c = src[i];
                if (c == '^' || c == '_')
                {
                    int j = i + 1; string g;
                    if (j < src.Length && src[j] == '{')
                    {
                        int k = src.IndexOf('}', j);
                        if (k < 0) { cur.Append(c); i++; continue; }
                        g = src.Substring(j + 1, k - j - 1); j = k + 1;
                        if (g.IndexOf('_') >= 0) g = Regex.Replace(g, @"_\{?([A-Za-z0-9]+)\}?", "$1");        // P_{CO_2}: cả cụm CO2 là chỉ số dưới
                    }
                    else
                    {
                        int k = j;
                        while (k < src.Length && (char.IsLetterOrDigit(src[k]) || src[k] == '+' || (c == '^' && src[k] == '-') || (c == '_' && src[k] == ','))) k++;
                        g = src.Substring(j, k - j); j = k;
                    }
                    if (g.Length == 0) { cur.Append(c); i++; continue; }
                    flush(); runs.Add(new Run(g.Replace('-', '−'), c == '^' ? 2 : 1)); i = j; continue;
                }
                cur.Append(c); i++;
            }
            flush();
            return runs;
        }

        static readonly Regex VarRe = new Regex(@"(?<![A-Za-zÀ-ỹα-ωΑ-Ω])([A-Za-zα-ωΑ-Ω])(?![A-Za-zÀ-ỹα-ωΑ-Ω])");

        /// <summary>Chữ thường trong biểu thức: biến một chữ cái (x, y, r, α) in nghiêng, phần còn lại đứng thẳng.</summary>
        static void AddItalicVars(List<Run> runs, string text)
        {
            int last = 0;
            foreach (Match m in VarRe.Matches(text))
            {
                if (m.Index > last) runs.Add(new Run(text.Substring(last, m.Index - last), 0));
                bool upperGreek = m.Value[0] >= 'Α' && m.Value[0] <= 'Ω';      // chữ Hy Lạp hoa (Δ, Σ, Ω) đứng thẳng
                runs.Add(new Run(m.Value, upperGreek ? 0 : 3)); last = m.Index + m.Length;
            }
            if (last < text.Length) runs.Add(new Run(text.Substring(last), 0));
        }

        // ------------------------------------------------------------------ ký hiệu theo ngữ cảnh
        static readonly string[][] QUANT = {
            new[] { "thickness", "t", "độ dày" }, new[] { "pressure", "P", "áp suất" }, new[] { "temperature", "T", "nhiệt độ" },
            new[] { "stress", "S", "ứng suất" }, new[] { "diameter", "D", "đường kính" }, new[] { "length", "L", "chiều dài" },
            new[] { "velocity", "v", "vận tốc" }, new[] { "flow rate", "Q", "lưu lượng" }, new[] { "area", "A", "diện tích" } };
        static readonly Regex StopW = new Regex(@"^(Pa|Is|It|In|On|Of|To|The|This|Thus|Then|Tab|As|At|An|And|Are|Be|By|If|Or)$");

        static Fix SymFix(string sym, string sub)
        {
            var runs = new List<Run>();
            runs.Add(new Run(sym, 3));
            if (sub.Length > 0) runs.Add(new Run(sub, 1));
            string linear = sym + (sub.Length > 0 ? "_(\"" + sub + "\")" : "");
            var f = Fix.FromRuns("Sửa", runs, linear);
            f.Show = sym + (sub.Length > 0 ? "_" + sub : "");
            return f;
        }

        // ------------------------------------------------------------------ nhận diện
        static readonly Regex IonRe = new Regex(@"(?<![A-Za-z0-9^])([A-Z][A-Za-z0-9()]*?)(\^?\{?\d*[+-]\}?)(?![A-Za-z0-9])");
        static readonly Regex FormulaRe = new Regex(@"(?<![A-Za-z0-9^_₀-₉])(?<![A-Za-z0-9)]\()(?<!_\{)([A-Z][A-Za-z0-9()]*)(?![A-Za-z0-9])");
        static readonly Regex MathRe = new Regex(
            @"[A-Za-z0-9.()]+(?:\^\{?[\w+\-]+\}?|_\{?[\w,]+\}?)?(?:\s*[=+*\/\-<>]\s*[A-Za-z0-9.()]+(?:\^\{?[\w+\-]+\}?|_\{?[\w,]+\}?)?)+",
            RegexOptions.None, TimeSpan.FromSeconds(2));
        // công thức hóa học gõ kiểu gạch dưới: H_2O, O_2, SO_4^{2-}, O^{2-}, Fe^{3+}
        static readonly Regex ChemUnderRe = new Regex(
            @"(?<![A-Za-z0-9_^])((?:(?:[A-Z][a-z]?(?:_\{?\d+\}?|\d+)?)|(?:\((?:[A-Z][a-z]?(?:_\{?\d+\}?|\d+)?)+\)(?:_\{?\d+\}?|\d+)?))+)(\^\{?\d*[+\-−]\}?)?(?![A-Za-z0-9_])");
        static readonly HashSet<string> Diatomic = new HashSet<string>(new[] { "H2", "N2", "O2", "F2", "Cl2", "Br2", "I2", "O3", "P4", "S8" });
        static readonly Regex GreekCommaRe = new Regex(@"(?<![\p{L}\p{N}])([\u03B1-\u03C9\u0391-\u03A9])([a-z]{1,4}),\s?([A-Za-z]{2,4})(?=\s*[=≈])");
        static readonly Regex ParenPowRe = new Regex(@"\(([^()]{1,40})\)\^(?:\{[^{}]{1,12}\}|[-−+]?\d+(?:[.,]\d+)?|[A-Za-z])");
        static readonly Regex UnderRe = new Regex(@"(?<![A-Za-z0-9_\\])([A-Za-z])_(?:\{([A-Za-z0-9,]+)\}|([A-Za-z0-9,]+))(?![A-Za-z0-9_])");

        public static List<Hit> Detect(string txt, ReviewOpts o, Func<NumCtx> ctx)
        {
            var ms = new List<Hit>(); var cl = new List<int[]>();
            Func<int, int, bool> free = (s, e) => !cl.Any(c => s < c[1] && e > c[0]);
            Action<int, int, Issue> add = delegate(int s, int e, Issue iss)
            {
                if (free(s, e)) { cl.Add(new[] { s, e }); ms.Add(new Hit { S = s, E = e, Issue = iss }); }
            };
            Match m;
            if (o.Sym || o.Math)
            {
                // 0a. cả đoạn công thức gõ kiểu LaTeX: Re_{cyl} = \frac{a}{b} = ..., f = 0,0791 · (7.725)^{-0,3}. Chạy đầu tiên để cả công thức là một mục, không bị tách vụn.
                foreach (MathSpan sp in MathSpans.Find(txt))
                {
                    string bf = txt.Substring(sp.S, sp.E - sp.S);
                    MNode node = MathParser.Parse(bf);
                    Fix nf = Fix.FromNode("Sửa", node);
                    var iss = new Issue { Cat = "math", Sev = sp.Leak ? "err" : "fmt", Key = "m|" + bf + "|" + CtxKey(txt, sp.S), Before = bf };
                    iss.Reason = (sp.Leak ? "Công thức đang hiện nguyên mã LaTeX (\\frac, { }, _, ^). " : "Công thức gõ phẳng. ")
                        + "Dựng lại thành Equation của Word: " + (sp.HasFrac ? "phân số xếp trên dưới, " : "") + "biến in nghiêng, chỉ số dưới và số mũ nằm đúng vị trí, chữ trong chỉ số (cyl, req...) đứng thẳng.";
                    iss.Fixes.Add(nf);
                    add(sp.S, sp.E, iss);
                }
                // 0b. chữ Hy Lạp kèm chỉ số gõ phẳng, phân cách bằng dấu phẩy: τw, RCE = ...
                foreach (Match gm in GreekCommaRe.Matches(txt))
                {
                    string sym = gm.Groups[1].Value, sub = gm.Groups[2].Value + "," + gm.Groups[3].Value;
                    var iss = new Issue { Cat = "sym", Sev = "sus", Key = "g|" + gm.Value + "|" + CtxKey(txt, gm.Index), Before = gm.Value,
                        Reason = "Ký hiệu có chỉ số gõ phẳng (" + gm.Value + "). Có thể là " + sym + " với chỉ số dưới \"" + sub + "\" đứng thẳng. Cần bạn xác nhận." };
                    iss.Fixes.Add(SymFix(sym, sub));
                    add(gm.Index, gm.Index + gm.Length, iss);
                }
                // 0. mã LaTeX lộ ra và ký hiệu gạch dưới: \Delta G_{ads}°, \tau_w, P_{CO_2}, t_{req}. Chạy trước hóa học để P_{CO_2} không bị tách làm đôi.
                int pos = 0;
                for (Match sm = Symbols.StartRe.Match(txt); sm.Success; sm = sm.NextMatch())
                {
                    if (sm.Index < pos) continue;
                    SymHit sh = Symbols.TryParse(txt, sm.Index);
                    if (sh == null) continue;
                    string bf = txt.Substring(sh.Start, sh.End - sh.Start);
                    var iss = new Issue { Cat = "sym", Sev = sh.Leak ? "err" : "fmt", Key = "s|" + bf + "|" + CtxKey(txt, sh.Start), Before = bf };
                    iss.Reason = sh.Leak
                        ? "Mã LaTeX đang hiển thị nguyên văn trong văn bản. Đổi thành ký hiệu thật " + sh.Show + " (chữ Hy Lạp hoa và toán tử đứng thẳng, biến in nghiêng, chỉ số là chỉ số dưới)."
                        : "Ký hiệu gõ kiểu gạch dưới. Biến in nghiêng, phần sau gạch dưới là chỉ số dưới (cả cụm, ví dụ CO2 cùng nằm ở chỉ số dưới).";
                    Fix sf = Fix.FromRuns("Sửa", sh.Runs, sh.Linear); sf.Show = sh.Show;
                    iss.Fixes.Add(sf);
                    add(sh.Start, sh.End, iss);
                    pos = sh.End;
                }
            }
            if (o.Chem)
            {
                // 1. phương trình phản ứng
                foreach (Match rm in Chem.RxnRe.Matches(txt))
                {
                    Rxn a = Chem.AnalyzeRxn(rm.Groups[1].Value, rm.Groups[2].Value, rm.Groups[3].Value);
                    if (a.Unknown) continue;
                    var kinds = a.Flags.Select(f => f.Type).ToList();
                    var reasons = a.Flags.Select(f => f.Msg).ToList();
                    string sev = "fmt";
                    if (kinds.Contains("case")) sev = "err";
                    if (kinds.Contains("val") || kinds.Contains("order") || a.Balanced == false) sev = kinds.Contains("case") ? "err" : "sus";
                    if (a.Balanced == false)
                        reasons.Add(a.BalStr != null ? "Phương trình chưa cân bằng. Gợi ý cân bằng: " + a.BalStr.Replace("->", "→")
                                                      : "Phương trình chưa cân bằng và chưa tự cân bằng được, hãy kiểm tra lại các chất");
                    var iss = new Issue { Cat = "chem", Sev = sev, Key = "r|" + rm.Value + "|" + CtxKey(txt, rm.Index), Before = rm.Value };
                    var f0 = Fix.FromChem("Sửa", a.Std);
                    if (sev == "fmt" && f0.Plain == rm.Value && !f0.NeedsFormat) continue;
                    iss.Fixes.Add(f0);
                    if (a.BalStr != null) iss.Fixes.Add(Fix.FromChem("Sửa và cân bằng", a.BalStr));
                    foreach (RxnAlt ra in a.Alts)
                    {
                        iss.Fixes.Add(Fix.FromChem(ra.Label, ra.Std));
                        if (ra.BalStr != null) iss.Fixes.Add(Fix.FromChem(ra.Label + ", cân bằng", ra.BalStr));
                    }
                    iss.Reason = reasons.Count > 0 ? string.Join(". ", reasons.ToArray()) + "."
                        : "Phương trình đang gõ phẳng. Chuyển chỉ số dưới và chỉ số trên đúng chuẩn.";
                    add(rm.Index, rm.Index + rm.Length, iss);
                }
                // 1b. công thức gõ kiểu gạch dưới: H_2O (chỉ số 2 là chỉ số dưới, chữ cái đứng thẳng)
                foreach (Match cu in ChemUnderRe.Matches(txt))
                {
                    string core = cu.Groups[1].Value, chg = cu.Groups[2].Value;
                    if (core.IndexOf('_') < 0 && chg.Length == 0) continue;
                    string plain = Regex.Replace(core, @"_\{?(\d+)\}?", "$1");
                    if (Regex.IsMatch(core, @"_\{?1\}?(?!\d)") && chg.Length == 0) continue;     // A_1, P_1V_1: chỉ số thứ tự của biến, không ai viết H_1
                    int groups = Regex.Matches(plain, "[A-Z][a-z]?").Count;
                    bool diat = Diatomic.Contains(plain);
                    if (groups < 2 && !(chg.Length > 0 || diat)) continue;      // T_2, P_1, D_2: biến chứ không phải hóa học
                    string charge = chg.Length > 0 ? "^" + Regex.Replace(chg.Substring(1), @"[{}]", "").Replace('−', '-') : "";
                    Species sp = Chem.NormSpecies(plain + charge);
                    if (sp.Unknown || sp.Comp == null) continue;
                    ChemFlag val = sp.Flags.FirstOrDefault(f => f.Type == "val");
                    if (val != null && val.Sug == null) continue;
                    var iss = new Issue { Cat = "chem", Sev = val != null ? "sus" : "fmt", Key = "r|" + cu.Value + "|" + CtxKey(txt, cu.Index), Before = cu.Value };
                    iss.Reason = val != null ? val.Msg + ". Kiểm tra lại số nguyên tử."
                        : "Công thức hóa học đang gõ kiểu gạch dưới. Chỉ số nguyên tử là chỉ số dưới" + (charge.Length > 0 ? ", điện tích là chỉ số trên" : "") + "; các chữ cái ký hiệu nguyên tố đứng thẳng, không nghiêng.";
                    iss.Fixes.Add(Fix.FromChem("Sửa", sp.Out));
                    add(cu.Index, cu.Index + cu.Length, iss);
                }
                // 2. ion đứng riêng: Fe2+, SO4^2-
                foreach (Match im in IonRe.Matches(txt))
                {
                    string b = im.Groups[1].Value;
                    if (!Regex.IsMatch(b, @"^[A-Z][a-z]?\d*$|^\(.*\)|[a-z)]")) continue;
                    Species sp = Chem.NormSpecies(im.Value);
                    if (sp.Unknown || sp.Comp == null || sp.Charge == 0) continue;
                    if (!Regex.IsMatch(im.Value, @"\d[+-]$|\^|[+-]$")) continue;
                    var iss = new Issue { Cat = "chem", Sev = "fmt", Key = "r|" + im.Value + "|" + CtxKey(txt, im.Index), Before = im.Value,
                        Reason = "Ion đang gõ phẳng. Điện tích " + sp.Out.Substring(sp.Out.LastIndexOf('^') + 1) + " phải là chỉ số trên." };
                    var f0 = Fix.FromChem("Sửa", sp.Out);
                    if (f0.Plain == im.Value && !f0.NeedsFormat) continue;
                    iss.Fixes.Add(f0);
                    add(im.Index, im.Index + im.Length, iss);
                }
                // 3. công thức đứng riêng: H2S, CO2, NaCl2, OH2
                foreach (Match fm in FormulaRe.Matches(txt))
                {
                    string tok = fm.Groups[1].Value;
                    while (tok.EndsWith(")") && tok.Count(ch => ch == ')') > tok.Count(ch => ch == '(')) tok = tok.Substring(0, tok.Length - 1);   // "(H2O)" -> H2O
                    if (tok.Length == 0 || Chem.ParseF(tok) == null) continue;
                    int groups = Regex.Matches(tok, "[A-Z][a-z]?").Count;
                    if (groups < 2 || !(Regex.IsMatch(tok, @"\d") || Chem.COMMON.Contains(tok))) continue;
                    Species sp = Chem.NormSpecies(tok);
                    if (sp.Unknown) continue;
                    ChemFlag val = sp.Flags.FirstOrDefault(f => f.Type == "val"), ord = sp.Flags.FirstOrDefault(f => f.Type == "order");
                    if (val != null && val.Sug == null) continue;   // không chắc là hóa học (ví dụ SS316)
                    string sev = "fmt", reason = "Công thức đang gõ phẳng. Chuyển số nguyên tử thành chỉ số dưới.";
                    if (val != null) { sev = "sus"; reason = val.Msg + ". Kiểm tra lại số nguyên tử."; }
                    else if (ord != null) { sev = "sus"; reason = ord.Msg + "."; }
                    var f0 = Fix.FromChem("Sửa", sp.Out);
                    if (sev == "fmt" && f0.Plain == tok && !f0.NeedsFormat) continue;
                    var iss = new Issue { Cat = "chem", Sev = sev, Key = "r|" + tok + "|" + CtxKey(txt, fm.Index), Before = tok, Reason = reason };
                    iss.Fixes.Add(f0);
                    add(fm.Index, fm.Index + tok.Length, iss);
                }
            }
            if (o.Math)
            {
                // 3b. lũy thừa của một nhóm trong ngoặc, gõ kiểu ^: (7.725)^{-0,3}, (a+b)^2
                foreach (Match pm in ParenPowRe.Matches(txt))
                {
                    if (!MathSpans.Mathy(pm.Groups[1].Value)) continue;
                    var iss = new Issue { Cat = "math", Sev = "fmt", Key = "p|" + pm.Value + "|" + CtxKey(txt, pm.Index), Before = pm.Value,
                        Reason = "Số mũ gõ kiểu ^ sau dấu ngoặc. Chuyển thành số mũ (chỉ số trên) thật." };
                    iss.Fixes.Add(Fix.FromNode("Sửa", MathParser.Parse(pm.Value)));
                    add(pm.Index, pm.Index + pm.Length, iss);
                }
                // 4. biểu thức toán gõ phẳng: x^2 + y^2 = r^2
                foreach (Match mm in MathRe.Matches(txt))
                {
                    if (mm.Value.IndexOf('^') < 0 && mm.Value.IndexOf('_') < 0) continue;
                    string norm = NormMath(mm.Value);
                    List<Run> runs = MathRuns(norm);
                    var iss = new Issue { Cat = "math", Sev = "fmt", Key = "r|" + mm.Value + "|" + CtxKey(txt, mm.Index), Before = mm.Value,
                        Reason = "Biểu thức toán đang gõ phẳng. Chuyển thành công thức Word hoặc chỉ số trên, dưới để số mũ và chỉ số nằm đúng vị trí." };
                    var f0 = Fix.FromRuns("Sửa", runs, norm);
                    iss.Fixes.Add(f0);
                    add(mm.Index, mm.Index + mm.Length, iss);
                }
            }
            if (o.Sym)
            {
                // 5. ký hiệu biến theo ngữ cảnh: "thickness, treq,min"
                foreach (string[] q in QUANT)
                {
                    string w = q[0], sym = q[1], vn = q[2];
                    var r2 = new Regex(@"\b" + w.Replace(" ", @"\s+") + @"\s*,?\s+(?:be\s+)?([A-Za-z])([A-Za-z]*(?:,[A-Za-z]+)*)(?![A-Za-z0-9])", RegexOptions.IgnoreCase);
                    foreach (Match qm in r2.Matches(txt))
                    {
                        if (qm.Groups[1].Value != sym) continue;
                        string sub = qm.Groups[2].Value, tok = qm.Groups[1].Value + sub;
                        if (StopW.IsMatch(tok)) continue;
                        bool comma = sub.IndexOf(',') >= 0;
                        if (sub.Length > 0 && !comma && !(sym == sym.ToUpperInvariant() && Regex.IsMatch(sub, "^[a-z]{1,4}$"))) continue;
                        int s = qm.Index + qm.Length - tok.Length;
                        var iss = new Issue { Cat = "sym", Sev = comma ? "err" : "sus", Key = "r|" + tok + "|" + CtxKey(txt, s), Before = tok };
                        iss.Reason = "Ngữ cảnh: \"" + w + "\" là " + vn + ", ký hiệu thường là " + sym + ". "
                            + (sub.Length > 0 ? "Phần \"" + sub + "\" là chỉ số dưới đứng thẳng, chữ " + sym + " in nghiêng." : "Biến số in nghiêng.")
                            + (comma ? "" : " Chưa chắc chắn nên cần bạn xác nhận.");
                        iss.Fixes.Add(SymFix(sym, sub));
                        add(s, s + tok.Length, iss);
                    }
                }
                // 6. dạng gạch dưới rõ ràng: t_req,min hoặc P_{H2S}
                foreach (Match um in UnderRe.Matches(txt))
                {
                    string sym = um.Groups[1].Value, sub = um.Groups[2].Success ? um.Groups[2].Value : um.Groups[3].Value;
                    var iss = new Issue { Cat = "sym", Sev = "fmt", Key = "r|" + um.Value + "|" + CtxKey(txt, um.Index), Before = um.Value,
                        Reason = "Ký hiệu gõ kiểu gạch dưới. Chữ " + sym + " in nghiêng, phần \"" + sub + "\" là chỉ số dưới." };
                    iss.Fixes.Add(SymFix(sym, sub));
                    add(um.Index, um.Index + um.Length, iss);
                }
            }
            if (o.Num) DetectNumbers(txt, o, ctx, add);

            // sắp xếp theo vị trí, bỏ mục chồng lấn
            ms.Sort((a, b) => a.S.CompareTo(b.S));
            var keep = new List<Hit>(); int end = -1;
            foreach (Hit h in ms) if (h.S >= end) { keep.Add(h); end = h.E; }
            return keep;
        }

        // ------------------------------------------------------------------ số kỹ thuật trong văn bản
        const string MON = @"(?:January|February|March|April|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sept|Sep|Oct|Nov|Dec)";
        static readonly Regex PFX = new Regex(@"(?<![\p{L}\p{N}])(?:page|pages|pp?\.?|table|tbl\.?|fig(?:ure)?s?\.?|rev(?:ision)?\.?|section|sec\.?|clause|item|no\.?|eq(?:uation)?\.?|chapter|appendix|ref\.?|step|line|row|column|part|phase|tag|dwg|drawing|doc(?:ument)?\.?|API|ASME|ISO|ASTM|EN|DIN|BS|NACE|ANSI|AWS|DNV|SP|RP|trang|bảng|hình|mục|điều|chương|khoản|điểm|bước|loại|cấp|mã|stt|id|nr|example|case|class|grade|type|series|model|level|stage|lần|đợt|số hiệu|số|ví dụ|biểu đồ|sơ đồ|phụ lục|phụ biểu|công thức|phương trình|hạng mục|giai đoạn|quy trình|tài liệu|thông tư|quyết định|nghị định|qcvn|tcvn|tcxd|gost|iec|ieee|nfpa|ansi/api|annex|tab\.?|ex\.?|version|ver\.?|v|edition|ed\.?|tiểu mục|phần|dòng|cột|hàng|lô|block|giếng|well|tank|bồn|bình|bơm|pump|unit|vessel|tuyến|trạm|kịch bản|scenario|test|thử nghiệm|mẫu|sample|specimen|coupon|run|sheet|tờ|bản vẽ|hình vẽ|ảnh|photo|plate|slide|scheme|equation no\.?|formula|eqn\.?|eq\.?|đề mục|danh mục|stt)\s*$", RegexOptions.IgnoreCase);
        static readonly Regex NUMRE = new Regex(@"(?<![A-Za-z0-9_.,\/])(?<!\d-)(\d+(?:[.,]\d+)*)(?![0-9_])");
        static readonly Regex MonAfter = new Regex(@"^\s+" + MON + @"\b", RegexOptions.IgnoreCase);
        static readonly Regex MonBefore = new Regex(@"\b" + MON + @"\.?\s+$", RegexOptions.IgnoreCase);

        static string Sub(string t, int a, int len) { if (a >= t.Length) return ""; return t.Substring(a, Math.Min(len, t.Length - a)); }

        static readonly Regex YearCue = new Regex(@"(?:\b(?:in|since|from|until|till|by|during|before|after|circa|ca|year|years|of|năm|từ|đến|vào|trong|sau|trước|khoảng|hồi)\s*|\(|c\.\s*)$", RegexOptions.IgnoreCase);
        static readonly Regex VaguePre = new Regex(@"(?:=|≈|~|≥|≤|>|<|\b(?:approximately|about|around|roughly|equals?|is|was|của|là|bằng|khoảng|xấp xỉ|tới|đạt)\s*)\s*$", RegexOptions.IgnoreCase);
        static readonly Regex DateWord = new Regex(@"(?:day|ngày|tháng|quý|quarter|q|week|tuần|năm|year|fy)\s*$", RegexOptions.IgnoreCase);

        static readonly Regex GroupUs = new Regex(@"^\d{1,3}(?:,\d{3})+(?:\.\d+)?$"), GroupVn = new Regex(@"^\d{1,3}(?:\.\d{3})+(?:,\d+)?$");
        /// <summary>Số có đúng dạng nhóm hàng nghìn không: 1,234,567.89 hoặc 1.234.567,89 (các nhóm sau đủ 3 chữ số).</summary>
        static bool GroupedThousands(string tok) { return GroupUs.IsMatch(tok) || GroupVn.IsMatch(tok); }

        /// <summary>
        /// Phân loại một số trong văn bản bằng ngữ cảnh. "text" = số mang tính văn bản (năm, ngày, số mục, số trang, trích dẫn, số hiệu): không đụng tới.
        /// "tech" = số kỹ thuật chắc chắn. "doubt" = chưa rõ. "skip" = bình thường, không báo.
        /// </summary>
        public static string ClassifyNumber(string txt, int s, int e, NumCtx c)
        {
            string tok = txt.Substring(s, e - s);
            string bef = txt.Substring(Math.Max(0, s - 24), s - Math.Max(0, s - 24)), aft = Sub(txt, e, 24);
            bool unit = Units.Scan(aft) != null;
            bool hasSep = tok.IndexOf('.') >= 0 || tok.IndexOf(',') >= 0;
            int nsep = tok.Count(ch => ch == '.' || ch == ',');
            if (c.Heading) return "text";
            // số thứ tự ở đầu đoạn: "1.1. Bối cảnh", "2,2. Quan hệ", "• 3.4) Nội dung", "1.1:"
            if (!unit && Regex.IsMatch(bef, @"^[\s•·\-–—*▪●○■□►]*$") && Regex.IsMatch(aft, @"^[.):](\s|$)")) return "text";
            if (MonAfter.IsMatch(aft) || MonBefore.IsMatch(bef)) return "text";                       // 15 March, March 2024
            if (Regex.IsMatch(aft, @"^[\/\-–—:]\d") || Regex.IsMatch(bef, @"\d[\/\-–—:]$")) return "text";    // 12/05/2024, 10-20
            if (PFX.IsMatch(bef)) return "text";                                                       // page 12, Table 3, ISO 9001, bảng 2
            if (bef.EndsWith("(") && aft.StartsWith(")")) return "text";                              // (1), (2024)
            if (bef.EndsWith("[") && (aft.StartsWith("]") || aft.StartsWith(",") || aft.StartsWith("-") || aft.StartsWith("–"))) return "text";   // trích dẫn [12], [1,2]
            if (Regex.IsMatch(aft, @"^s\b") && tok.Length == 4) return "text";                       // 1990s
            if (DateWord.IsMatch(bef) && !unit) return "text";                                         // ngày 12, tháng 5, năm 1990
            if (tok.Length > 1 && tok[0] == '0' && char.IsDigit(tok[1])) return "text";                // 007, 0123: mã, số điện thoại
            if (nsep >= 2 && !GroupedThousands(tok)) return "text";   // 3.2.1, 12.05.2020, 1.2.3.4: số mục, ngày, phiên bản
            if (!hasSep && tok.Length == 4 && !unit)
            {
                int yr = int.Parse(tok, CultureInfo.InvariantCulture);
                if (yr >= 1000 && yr <= 2100) return "text";                                          // năm
            }
            // số mục ở đầu dòng: "3.2 Corrosion rate", "1.1 Giới thiệu"
            if (!unit && hasSep && Regex.IsMatch(bef, @"^\s*$") && Regex.IsMatch(aft, @"^\s+[A-ZÀ-Ỵ]")) return "text";
            if (unit) return "tech";
            if (c.InTable && c.NumericCell) return "tech";
            if (hasSep)
            {
                ParsedNumber p = NumberUtil.Parse(tok);
                if (p == null) return "text";
                if (!p.Amb) return "tech";
                return VaguePre.IsMatch(bef) ? "tech" : "doubt";       // 1.234 có thể là một nghìn hai trăm... hay 1,234: chỉ nghi ngờ khi không có manh mối
            }
            if (tok.Length >= 5) return "doubt";                        // 12500 không đơn vị: có thể là 12,500
            if (tok.Length == 4 && VaguePre.IsMatch(bef)) return "tech";    // "= 2500", "approximately 3000"
            return "skip";                                              // số nguyên thường ngày, kể cả 4 chữ số không có manh mối
        }

        /// <summary>Đơn vị viết chưa chuẩn: rpm/RPM, kpa/kPa, m3/h thành m³/h, ℃ thành °C.</summary>
        static void AddUnitIssue(string txt, int e, UnitHit uh, Action<int, int, Issue> add)
        {
            int s = e + uh.Skip, en = e + uh.Len;
            string raw = txt.Substring(s, en - s);
            var iss = new Issue { Cat = "unit", Sev = uh.Ambiguous ? "sus" : (uh.SiCase ? "err" : "fmt"), Key = "u|" + raw + "|" + CtxKey(txt, s), Before = raw };
            iss.Reason = "Đơn vị \"" + raw + "\" viết chưa đúng chuẩn. " + uh.Why + (uh.Ambiguous ? " Chưa chắc chắn nên cần bạn xác nhận." : "");
            Fix f = Fix.FromRuns("Sửa", uh.CanonRuns, Chem.RunsToLinear(uh.CanonRuns)); f.Show = uh.Canon;
            iss.Fixes.Add(f);
            foreach (string alt in uh.Alts) iss.Fixes.Add(Fix.Text("Hiểu là " + alt, alt, true));
            add(s, en, iss);
        }

        /// <summary>Thiếu dấu cách giữa giá trị và đơn vị: 25ppm, 94,86%, 80°C. Chèn dấu cách không ngắt, đồng thời chuẩn lại đơn vị nếu cũng sai.</summary>
        static void AddSpaceIssue(string txt, int e, UnitHit uh, Action<int, int, Issue> add)
        {
            int en = e + uh.Len;
            string raw = txt.Substring(e, en - e);
            var iss = new Issue { Cat = "unit", Sev = "fmt", Key = "sp|" + raw + "|" + CtxKey(txt, e), Before = raw };
            iss.Reason = "Thiếu dấu cách giữa giá trị và đơn vị \"" + raw + "\". Quy ước SI cần một dấu cách không ngắt (để số và đơn vị không bị tách xuống dòng khác)." + (uh.Fixable ? " " + uh.Why : "");
            var runs = new List<Run>(); runs.Add(new Run("\u00A0", 0)); runs.AddRange(uh.CanonRuns);
            Fix f = Fix.FromRuns("Thêm dấu cách", runs, "\u00A0" + Chem.RunsToLinear(uh.CanonRuns));
            f.Show = "(cách) " + uh.Canon;
            iss.Fixes.Add(f);
            foreach (string alt in uh.Alts) iss.Fixes.Add(Fix.Text("Hiểu là " + alt, "\u00A0" + alt, true));
            add(e, en, iss);
        }

        static readonly Regex SentenceBreak = new Regex(@"[;!?\r\n\a]|[.]\s|\s[:—]\s");
        static readonly Regex LocalNumRe = new Regex(
            @"(?<![A-Za-z0-9_.,/])(?<!\d-)(\d{1,3}(?:[ \u00A0\u202F]\d{3})+(?:[.,]\d+)?(?![0-9.,])|\d+(?:[.,]\d+)*)(?![0-9_])");
        static readonly Regex ContextWord = new Regex(@"\p{L}[\p{L}\p{M}]*");
        static readonly Regex ViMarks = new Regex("[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]", RegexOptions.IgnoreCase);
        static readonly HashSet<string> EnglishWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "the", "is", "was", "are", "of", "at", "to", "and", "with", "from", "measured",
            "pressure", "temperature", "thickness", "rate", "flow", "value", "test", "limit",
            "speed", "load", "design", "diameter", "depth", "density", "weight", "mass", "length", "width",
            "height", "minimum", "maximum", "shall", "must", "reading", "coefficient", "result", "corrosion",
            "equals", "about", "approximately", "between", "above", "below"
        };

        // Score words close to this number inside the same clause. Unit symbols and other numbers do not vote.
        static string LocalNumberLanguage(string text, int start, int end, NumCtx context)
        {
            int left = Math.Max(0, start - 100), right = Math.Min(text.Length, end + 100);
            // A split found in the text before the number delimits the left context.
            var before = SentenceBreak.Matches(text.Substring(left, start - left));
            if (before.Count > 0)
            {
                Match last = before[before.Count - 1];
                left += last.Index + last.Length;
            }
            Match after = SentenceBreak.Match(text.Substring(end, right - end));
            if (after.Success) right = end + after.Index;

            int vi = 0, en = 0;
            foreach (Match word in ContextWord.Matches(text.Substring(left, right - left)))
            {
                int wordStart = left + word.Index, wordEnd = wordStart + word.Length;
                int distance = wordEnd <= start ? start - wordEnd : wordStart >= end ? wordStart - end : 100;
                if (distance >= 75) continue;
                int weight = distance < 18 ? 4 : distance < 38 ? 2 : 1;
                if (ViMarks.IsMatch(word.Value.Normalize(NormalizationForm.FormC))) vi += weight;
                else if (EnglishWords.Contains(word.Value)) en += weight;
            }
            if (vi >= en + 2) return "vi";
            if (en >= vi + 2) return "en";
            return context.Language == null ? "" : context.Language(left, right);
        }

        static string FmtVal(ParsedNumber p, ReviewOpts o, string style)
        {
            int dec = o.NumDec == "keep" ? p.Decimals : int.Parse(o.NumDec, CultureInfo.InvariantCulture);
            if (o.NumDec != "keep" && p.Decimals == 0 && !o.NumAddDec) dec = 0;
            return NumberUtil.Format(p.D, dec, style, o.NumThou);
        }

        static void DetectNumbers(string txt, ReviewOpts o, Func<NumCtx> ctxf, Action<int, int, Issue> add)
        {
            if (!Regex.IsMatch(txt, @"\d")) return;
            NumCtx c = ctxf != null ? ctxf() : new NumCtx();
            if (!o.NumTable && c.InTable) return;
            bool automatic = o.NumStyle == "auto";
            MatchCollection all = (automatic ? LocalNumRe : NUMRE).Matches(txt);
            char hint = '\0';
            if (!automatic)
            {
                hint = NumberUtil.DetectHint(all.Cast<Match>().Select(x => x.Groups[1].Value));
                if (hint == '\0') hint = NumberUtil.HintForStyle(o.NumStyle);
            }
            foreach (Match nm in all)
            {
                string tok = nm.Groups[1].Value; int s = nm.Index, e = s + tok.Length;
                if (e < txt.Length && (char.IsLetter(txt[e]) || txt[e] == '°') && Units.Scan(Sub(txt, e, 24)) == null) continue;   // 5G, 2nd, 10x: không phải số kèm đơn vị
                UnitHit uh = Units.Scan(Sub(txt, e, 24));
                string k = ClassifyNumber(txt, s, e, c);
                if (uh != null && !c.Heading)
                {
                    if (uh.Skip == 0 && uh.Canon != "°" && k != "text") AddSpaceIssue(txt, e, uh, add);
                    else if (uh.Fixable) AddUnitIssue(txt, e, uh, add);
                }
                if (k == "text") { o.Stats.AutoText++; continue; }
                if (k == "skip") continue;
                ParsedNumber p0 = NumberUtil.Parse(tok);          // để biết số có mơ hồ không
                if (p0 == null) continue;
                string language = automatic ? LocalNumberLanguage(txt, s, e, c) : "";
                string style = automatic ? (language == "en" ? "us" : language == "vi" ? "vn" : "") : o.NumStyle;
                if (automatic && style.Length == 0) continue;
                // A comma or point followed by exactly three digits can change the value by 1000.
                // In automatic mode respect the local convention, without guessing a different value.
                char localMark = style == "us" ? '.' : ',';
                ParsedNumber p = NumberUtil.Parse(tok, automatic ? localMark : hint);
                if (p == null) continue;
                bool decimalMismatch = automatic && p.Decimals > 0 && p.Dec != localMark;
                string convention = automatic ? (language == "en" ? "tiếng Anh quanh số" : "tiếng Việt quanh số") : "quy ước đang chọn";
                string unit = uh != null ? uh.Raw : "";
                if (k == "tech")
                {
                    bool spaced = automatic && tok.IndexOfAny(new[] { ' ', '\u00A0', '\u202F' }) >= 0;
                    if (spaced && !decimalMismatch && o.NumDec == "keep") { o.Stats.OkNum++; continue; }
                    string output = spaced && o.NumDec == "keep" ? tok.Replace(p.Dec, localMark) : FmtVal(p, o, style);
                    if (output == tok) { o.Stats.OkNum++; continue; }
                    string why = unit.Length > 0 ? "Số kỹ thuật, đơn vị \"" + unit + "\"." : (c.InTable ? "Số trong bảng." : "Số kỹ thuật.");
                    var iss = new Issue { Cat = "num", Sev = decimalMismatch ? "err" : "fmt", NumStyle = style, Key = "n|" + tok + "|" + CtxKey(txt, s) + "|" + style, Before = tok,
                        Reason = why + " Đọc là " + p.D.ToString(CultureInfo.InvariantCulture) + (p.Decimals > 0 ? " (" + p.Decimals + " số thập phân)" : "") + ", đổi sang " + output + " theo " + convention + "." };
                    iss.Fixes.Add(Fix.Text("Sửa", output, false));
                    add(s, e, iss);
                }
                else
                {
                    string digits = Regex.Replace(tok, @"[.,\s]", "");
                    decimal integerValue;
                    if (!decimal.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out integerValue)) continue;
                    string th = NumberUtil.Format(integerValue, 0, style, o.NumThou);
                    var iss = new Issue { Cat = "num", Sev = "sus", NumStyle = style, Key = "n|" + tok + "|" + CtxKey(txt, s) + "|" + style, Before = tok };
                    if (th != tok) iss.Fixes.Add(Fix.Text("Hiểu là hàng nghìn", th, true));
                    if (p0.Amb)
                    {
                        char sep = tok.IndexOf(',') >= 0 ? ',' : '.';
                        ParsedNumber pd = NumberUtil.Parse(tok, sep);       // dấu này là thập phân
                        string dc = FmtVal(pd, o, style);
                        if (dc != tok || iss.Fixes.Count == 0) iss.Fixes.Add(Fix.Text("Hiểu là số thập phân", dc, true));
                        iss.Reason = "Chưa rõ \"" + tok + "\" là " + digits + " (dấu là phân cách hàng nghìn) hay " + pd.D.ToString(CultureInfo.InvariantCulture) + " (dấu là thập phân). Không có đơn vị đi kèm để suy ra.";
                    }
                    else iss.Reason = "Số có từ 4 chữ số và không có đơn vị đi kèm, chưa chắc là số kỹ thuật (có thể là mã số hoặc năm).";
                    if (iss.Fixes.Count == 0) continue;     // đã đúng quy ước theo mọi cách hiểu
                    add(s, e, iss);
                }
            }
        }
    }
}
