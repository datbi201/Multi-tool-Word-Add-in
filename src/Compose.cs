// Trợ lý Kỹ thuật - Rev0.4: Soạn phương trình. Tự nhận hóa học hay toán, lý từ chuỗi gõ; phát hiện, cảnh báo và đề xuất sửa công thức, phương trình viết sai.
// Chỉ làm việc trên chuỗi, không gọi Word.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    public sealed class ComposeOption
    {
        public string Label;
        public Fix Fix;
        public string Show { get { return Fix == null ? "" : Fix.Show; } }
    }

    public sealed class ComposeResult
    {
        public bool Ok;
        public string Kind = "";                       // "chem" hoặc "math" (tự nhận)
        public string Message = "";                    // giải thích, cảnh báo hoặc lỗi (tiếng Việt)
        public Fix Fix;                                // phương án đang chọn: các đoạn chữ có định dạng và chuỗi UnicodeMath
        public bool Complex;                           // có phân số, căn, tổng, tích phân...: bắt buộc dùng Equation của Word
        public string Show = "";                       // bản xem trước dạng chữ
        public List<string> Warnings = new List<string>();
        public List<ComposeOption> Options = new List<ComposeOption>();
        public int Selected;                           // phương án mặc định

        public void Select(int i)
        {
            if (Options.Count == 0) return;
            if (i < 0 || i >= Options.Count) i = 0;
            Selected = i; Fix = Options[i].Fix; Show = Fix.Show;
        }
    }

    public static class Compose
    {
        static readonly Regex ComplexRe = new Regex(@"[/√∑∫∏]|\\[A-Za-z]|_\(|\^\(");
        static readonly Regex ArrowRe = new Regex(@"<=>|<->|=>|->|-->|→|⇌");

        /// <summary>Soạn phương trình, tự nhận loại. Chuỗi gõ đơn giản: 2H2 + O2 -> 2H2O, SO4^2-, H_2O, x^2 + y^2 = r^2, t_{req} = (P D)/(2 S E), \Delta G, sqrt(x), sum_(i=1)^n.</summary>
        public static ComposeResult Build(string src)
        {
            string t = (src ?? "").Trim();
            var res = new ComposeResult();
            if (t.Length == 0) { res.Message = "Chưa nhập phương trình."; return res; }
            return BuildAs(t, DetectKind(t), 2);
        }

        /// <summary>Nhận ra chuỗi gõ là hóa học (phương trình phản ứng, ion, chất) hay toán, lý.</summary>
        public static string DetectKind(string t)
        {
            string s = Regex.Replace(t, @"_\{?(\d+)\}?", "$1").Replace('−', '-').Trim();
            if (Regex.IsMatch(s, @"\\[A-Za-z]") || Regex.IsMatch(s, @"[/√∑∫∏*]")) return "math";
            Match m = Chem.RxnRe.Match(s);
            if (m.Success && m.Index == 0 && m.Length == s.Length)
            {
                Rxn a = Chem.AnalyzeRxn(m.Groups[1].Value, m.Groups[2].Value, m.Groups[3].Value);
                if (!a.Unknown && (ArrowRe.IsMatch(m.Groups[2].Value) || a.Species.Any(ChemLike))) return "chem";
                return "math";
            }
            var parts = Regex.Split(s, @"\s+\+\s+(?=\d*[A-Za-z(])").Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
            if (parts.Count == 0) return "math";
            if (s.IndexOf('^') >= 0 && !Regex.IsMatch(s, @"\^\{?\d*[+\-]\}?(?![\w])")) return "math";     // x^2: số mũ, không phải điện tích
            bool any = false;
            foreach (string p in parts)
            {
                if (Regex.IsMatch(p, @"\s")) return "math";                        // "n + 1", "H2 O": có khoảng trắng giữa các ký tự thì không phải một chất
                Species sp = Chem.NormSpecies(p);
                if (sp.Unknown) return "math";
                if (ChemLike(sp)) any = true;
            }
            return any ? "chem" : "math";
        }

        static bool ChemLike(Species sp)
        {
            if (sp.Unknown || sp.Comp == null) return false;
            if (sp.Comp.Keys.Count >= 2 || sp.Charge != 0) return true;
            return Regex.IsMatch(sp.Out, @"\d") || (sp.Out.Length >= 2 && char.IsLower(sp.Out[1]));
        }

        /// <summary>Dựng theo loại đã biết (kind = "chem" hoặc "math"). balance = chọn sẵn bản cân bằng nếu có.</summary>
        public static ComposeResult Build(string src, string kind, bool balance) { return BuildAs(src, kind, balance ? 1 : 0); }

        /// <summary>balMode: 0 = không chọn sẵn bản cân bằng, 1 = chọn sẵn, 2 = tự chọn khi người dùng chưa gõ hệ số nào.</summary>
        static ComposeResult BuildAs(string src, string kind, int balMode)
        {
            var res = new ComposeResult();
            string t = (src ?? "").Trim();
            if (t.Length == 0) { res.Message = "Chưa nhập phương trình."; return res; }
            res.Kind = kind;
            if (kind == "chem") BuildChem(t, balMode, res); else BuildMath(t, res);
            if (res.Options.Count > 0) { res.Select(res.Selected); res.Ok = true; }
            return res;
        }

        static void AddOpt(ComposeResult res, string label, string std)
        {
            foreach (ComposeOption o in res.Options) if (o.Fix.Plain == Chem.RunsPlain(Chem.RenderRuns(std)) && o.Show == Detector.ShowText(Chem.RenderRuns(std))) return;
            res.Options.Add(new ComposeOption { Label = label, Fix = Fix.FromChem(label, std) });
        }

        static void BuildChem(string t, int balMode, ComposeResult res)
        {
            string s = Regex.Replace(t, @"_\{?(\d+)\}?", "$1");                 // H_2O -> H2O
            s = s.Replace('−', '-');
            var notes = res.Warnings;
            Match m = Chem.RxnRe.Match(s);
            if (m.Success && m.Index == 0 && m.Length == s.Length)
            {
                Rxn a = Chem.AnalyzeRxn(m.Groups[1].Value, m.Groups[2].Value, m.Groups[3].Value);
                if (!a.Unknown)
                {
                    foreach (ChemFlag f in a.Flags) notes.Add(f.Msg);
                    bool valFix = a.Flags.Any(f => f.Type == "val" && f.Sug != null);
                    bool unbalanced = a.Balanced == false;
                    string first = valFix || a.Flags.Count > 0 ? "Đã sửa lỗi công thức" : "Công thức chuẩn";
                    AddOpt(res, first, a.Std);
                    if (unbalanced && a.BalStr != null) { AddOpt(res, "Sửa và cân bằng", a.BalStr); }
                    if (valFix) AddOpt(res, "Giữ công thức như đã gõ", Chem.BuildRxn(a.Species, a.NL, a.Arrow, null, true));
                    foreach (RxnAlt ra in a.Alts)
                    {
                        AddOpt(res, ra.Label, ra.Std);
                        if (ra.BalStr != null) AddOpt(res, ra.Label + ", cân bằng", ra.BalStr);
                    }
                    if (unbalanced)
                    {
                        if (a.BalStr != null)
                        {
                            notes.Add("Phương trình chưa cân bằng. Bản cân bằng: " + a.BalStr.Replace("->", "→").Replace("<=>", "⇌"));
                            if (balMode == 1 || (balMode == 2 && !a.HasCoef)) res.Selected = res.Options.FindIndex(o => o.Label == "Sửa và cân bằng");
                        }
                        else notes.Add("Phương trình chưa cân bằng và chưa tự cân bằng được, hãy kiểm tra lại các chất");
                    }
                    if (res.Selected < 0) res.Selected = 0;
                    res.Message = notes.Count > 0 ? string.Join(". ", notes.ToArray()) + "." : "Phương trình đúng chuẩn: số nguyên tử là chỉ số dưới, điện tích là chỉ số trên, chữ cái đứng thẳng.";
                    return;
                }
            }
            var parts = Regex.Split(s, @"\s+\+\s+(?=\d*[A-Za-z(])").Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
            var specs = parts.Select(Chem.NormSpecies).ToList();
            var outs = new List<string>(); var typed = new List<string>(); bool unknown = false, valChange = false;
            foreach (Species sp in specs)
            {
                if (sp.Unknown) { unknown = true; outs.Add(sp.Raw); typed.Add(sp.Raw); continue; }
                foreach (ChemFlag f in sp.Flags) notes.Add(f.Msg);
                if (sp.Flags.Any(f => f.Type == "val" && f.Sug != null)) valChange = true;
                outs.Add(sp.Coef + sp.Out);
                typed.Add(sp.Coef + (sp.TypedOut.Length > 0 ? sp.TypedOut : sp.Out) + (sp.Charge != 0 ? (sp.Out.Contains("^") ? sp.Out.Substring(sp.Out.IndexOf('^')) : "") : ""));
            }
            AddOpt(res, notes.Count > 0 ? "Đã sửa lỗi công thức" : "Công thức chuẩn", string.Join(" + ", outs.ToArray()));
            if (valChange) AddOpt(res, "Giữ công thức như đã gõ", string.Join(" + ", typed.ToArray()));
            // chất có nhiều gợi ý: thêm cách đọc khác cho chất đầu tiên như vậy
            foreach (Species sp in specs)
            {
                if (sp.Alts.Count < 2) continue;
                for (int k = 1; k < sp.Alts.Count; k++)
                {
                    var alt = new List<string>();
                    foreach (Species q in specs) alt.Add(q == sp ? q.Coef + sp.Alts[k] : (q.Unknown ? q.Raw : q.Coef + q.Out));
                    AddOpt(res, "Đọc " + sp.TypedOut + " là " + sp.Alts[k], string.Join(" + ", alt.ToArray()));
                }
                break;
            }
            if (unknown) notes.Add("Không nhận ra hết các chất, chỉ chuyển chỉ số và điện tích theo cách gõ");
            res.Message = notes.Count > 0 ? string.Join(". ", notes.ToArray()) + "." : "Công thức hóa học: số nguyên tử là chỉ số dưới, điện tích là chỉ số trên, chữ cái đứng thẳng.";
        }

        /// <summary>Chuỗi toán sang UnicodeMath để Word dựng thành Equation: ^{ab} thành ^(ab), _{ab} thành _(ab).</summary>
        public static string ToUnicodeMath(string norm)
        {
            string s = Regex.Replace(norm, @"_\{?([A-Za-z0-9]+)_\{?([A-Za-z0-9]+)\}?\}", "_($1$2)");      // _{CO_2}
            s = Regex.Replace(s, @"\^\{([^}]*)\}", "^($1)");
            s = Regex.Replace(s, @"_\{([^}]*)\}", "_($1)");
            return s;
        }

        static void BuildMath(string t, ComposeResult res)
        {
            string src = t;
            string rep = Symbols.ReplaceCommands(src);
            if (rep != src) res.Warnings.Add("Đã đổi mã LaTeX thành ký hiệu thật");
            src = rep;
            src = Regex.Replace(src, @"\bsum(?=\s*[_^(])", "∑"); src = Regex.Replace(src, @"\bprod(?=\s*[_^(])", "∏"); src = Regex.Replace(src, @"\bint(?=\s*[_^(])", "∫");
            string norm = Detector.NormMath(src);
            if (norm.Count(c => c == '(') != norm.Count(c => c == ')')) res.Warnings.Add("Số dấu ngoặc tròn mở và đóng chưa khớp");
            if (norm.Count(c => c == '{') != norm.Count(c => c == '}')) res.Warnings.Add("Số dấu ngoặc nhọn mở và đóng chưa khớp");
            foreach (Match um in Regex.Matches(norm, @"\\[A-Za-z]+")) res.Warnings.Add("Mã " + um.Value + " chưa hỗ trợ, giữ nguyên");
            res.Complex = ComplexRe.IsMatch(norm) || Regex.IsMatch(norm, @"\b(?:sqrt|sum|int|lim)\b");
            List<Run> runs = Detector.MathRuns(norm);
            Fix mf = Fix.FromRuns("Chèn", runs, ToUnicodeMath(norm));
            mf.Show = Fix.ShowRuns(runs);                                              // chỉ số chữ hiện dạng _req để phân biệt với chữ thường
            try { mf.Node = MathParser.Parse(norm); } catch { mf.Node = null; }       // dựng Equation bằng OMML: phân số xếp trên dưới, căn, chỉ số đúng chỗ
            res.Options.Add(new ComposeOption { Label = "Công thức đã chuẩn hóa", Fix = mf });
            res.Message = (res.Warnings.Count > 0 ? string.Join(". ", res.Warnings.ToArray()) + ". " : "")
                + (res.Complex
                    ? "Có phân số, căn hoặc ký hiệu toán học nên sẽ chèn bằng Equation của Word."
                    : "Biến một chữ cái in nghiêng, số mũ và chỉ số được đặt đúng vị trí.");
        }
    }
}
