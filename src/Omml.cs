// Trợ lý Kỹ thuật - Rev0.5: cây công thức (MNode), bộ đọc LaTeX rút gọn, và bộ dựng OMML (Equation chuẩn của Word).
// Chỉ làm việc trên chuỗi, không gọi Word. Công thức dựng thẳng bằng OMML nên không còn dấu nháy, ngoặc thừa như khi nhờ Word BuildUp.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace TroLyKyThuat
{
    public enum MK { Text, Seq, Paren, Sub, Sup, SubSup, Frac, Rad, Acc }

    /// <summary>Một nút công thức. Text: chữ (Up = đứng thẳng, ngược lại in nghiêng). Frac: A = tử, B = mẫu. Sub, Sup, SubSup: A = cơ sở, B = chỉ số dưới, C = chỉ số trên.</summary>
    public sealed class MNode
    {
        public MK K; public string T = ""; public bool Up, Op, Sp, Slash;
        public MNode A, B, C;
        public List<MNode> Kids = new List<MNode>();

        public static MNode Txt(string t, bool up) { return new MNode { K = MK.Text, T = t, Up = up }; }
        public static MNode Space() { return new MNode { K = MK.Text, T = " ", Up = true, Sp = true }; }
        public static MNode Seq(List<MNode> kids) { var n = new MNode { K = MK.Seq }; n.Kids = kids; return n; }
        public static MNode Of(MK k, MNode a, MNode b, MNode c) { return new MNode { K = k, A = a, B = b, C = c }; }
    }

    public static class MathParser
    {
        static readonly HashSet<string> Words = new HashSet<string>(new[] { "ln", "log", "exp", "sin", "cos", "tan", "cot", "max", "min", "lim", "Re", "Nu", "Pr", "Sc", "Sh", "Ra", "Gr", "det", "arg" });
        const string SupU = "⁰¹²³⁴⁵⁶⁷⁸⁹⁻⁺ⁿⁱ", SupA = "0123456789-+ni", SubU = "₀₁₂₃₄₅₆₇₈₉₊₋", SubA = "0123456789+-";
        const string OpChars = "=≈≠+-−–·×*/<>≤≥±∓→⇌←↔⇒⇔∝∼∈≪≫≡,;:|[]";

        public static bool IsGreek(char c) { return (c >= '\u0391' && c <= '\u03A9') || (c >= '\u03B1' && c <= '\u03C9'); }
        public static bool IsUniSup(char c) { return SupU.IndexOf(c) >= 0; }

        static string Normalize(string s)
        {
            s = s.Replace("$", "");
            s = Regex.Replace(s, @"(?<![\p{L}\p{N}])([\u03B1-\u03C9\u0391-\u03A9])([a-z]{1,4}),\s?([A-Za-z]{2,4})(?=\s*[=≈])", "$1_{$2,$3}");   // τw, RCE = ... thành τ_{w,RCE}
            s = Regex.Replace(s, @"\\(?:left|right|big|Big|bigg|Bigg)(?![A-Za-z])\s*", "");
            s = Regex.Replace(s, @"\\[,;:! ]|\\quad|\\qquad|\\\\", " ");
            s = Regex.Replace(s, @"\\[()\[\]]", "");
            return s;
        }

        public static MNode Parse(string src)
        {
            var p = new Parser(Normalize(src ?? ""), false);
            return p.Seq(false);
        }

        sealed class Parser
        {
            readonly string s; int i; readonly bool up;
            public Parser(string text, bool upright) { s = text; up = upright; }

            public MNode Seq(bool inGroup)
            {
                var kids = new List<MNode>();
                while (i < s.Length)
                {
                    char c = s[i];
                    if (c == '}') { i++; continue; }
                    MNode a = Atom();
                    if (a == null) continue;
                    kids.Add(Scripts(a));
                }
                return MNode.Seq(ApplySlash(kids));
            }

            static bool IsOp(MNode n) { return n.K == MK.Text && n.Op; }

            /// <summary>a/b thành phân số: tử là dãy ký hiệu viết liền ngay trước dấu /, mẫu là nhóm hoặc ký hiệu ngay sau.</summary>
            static List<MNode> ApplySlash(List<MNode> kids)
            {
                if (!kids.Any(k => k.K == MK.Text && k.T == "/" && k.Op)) return kids;
                var o = new List<MNode>();
                for (int j = 0; j < kids.Count; j++)
                {
                    MNode k = kids[j];
                    if (k.K == MK.Text && k.T == "/" && k.Op)
                    {
                        int r = j + 1; while (r < kids.Count && kids[r].Sp) r++;
                        int lastReal = o.Count - 1; while (lastReal >= 0 && o[lastReal].Sp) lastReal--;
                        if (r < kids.Count && !IsOp(kids[r]) && lastReal >= 0 && !IsOp(o[lastReal]))
                        {
                            int st = lastReal;
                            if (o[lastReal].K != MK.Paren)
                                while (st - 1 >= 0 && (o[st - 1].Sp || !IsOp(o[st - 1])) && o[st - 1].K != MK.Paren) st--;
                            while (st < o.Count && o[st].Sp) st++;
                            var numKids = o.GetRange(st, lastReal - st + 1);
                            o.RemoveRange(st, o.Count - st);
                            MNode num = numKids.Count == 1 ? numKids[0] : MNode.Seq(numKids);
                            var f = MNode.Of(MK.Frac, num, kids[r], null); f.Slash = true;
                            o.Add(f); j = r; continue;
                        }
                    }
                    o.Add(k);
                }
                return o;
            }

            string RawGroup()
            {
                // đọc {...} cân ngoặc, trả về nội dung bên trong (không gồm dấu ngoặc), không tiến i nếu không phải {
                if (i >= s.Length || s[i] != '{') return null;
                int d = 0, j = i;
                for (; j < s.Length; j++) { if (s[j] == '{') d++; else if (s[j] == '}') { d--; if (d == 0) break; } }
                if (j >= s.Length) { string rest = s.Substring(i + 1); i = s.Length; return rest; }
                string inner = s.Substring(i + 1, j - i - 1); i = j + 1; return inner;
            }

            void SkipWs() { while (i < s.Length && char.IsWhiteSpace(s[i])) i++; }

            MNode GroupNode(bool upright)
            {
                SkipWs();
                if (i >= s.Length) return MNode.Seq(new List<MNode>());
                if (s[i] == '{') { string g = RawGroup(); return new Parser(g, upright).Seq(false); }
                MNode a = Atom();
                return a ?? MNode.Seq(new List<MNode>());
            }

            static string StripCmds(string g)
            {
                g = Regex.Replace(g, @"\\(?:mathrm|text|textrm|rm|mathit|textit|mathbf|operatorname)\s*\{([^{}]*)\}", "$1");
                g = Symbols.ReplaceCommands(g);
                return g.Replace("{", "").Replace("}", "").Replace("\\", "");
            }

            MNode Atom()
            {
                char c = s[i];
                if (char.IsWhiteSpace(c)) { while (i < s.Length && char.IsWhiteSpace(s[i])) i++; return MNode.Space(); }
                if (c == '\\') return Command();
                if (c == '{') { string g = RawGroup(); return new Parser(g, up).Seq(false); }
                if (c == '√') { i++; return MNode.Of(MK.Rad, GroupNode(up), null, null); }
                if (c == '(')
                {
                    int d = 0, j = i;
                    for (; j < s.Length; j++) { if (s[j] == '(') d++; else if (s[j] == ')') { d--; if (d == 0) break; } }
                    if (j < s.Length)
                    {
                        string inner = s.Substring(i + 1, j - i - 1); i = j + 1;
                        return MNode.Of(MK.Paren, new Parser(inner, up).Seq(false), null, null);
                    }
                    i++; return OpNode("(");
                }
                if (char.IsDigit(c))
                {
                    int j = i; while (j < s.Length && char.IsDigit(s[j])) j++;
                    while (j + 1 < s.Length && (s[j] == '.' || s[j] == ',') && char.IsDigit(s[j + 1])) { j++; while (j < s.Length && char.IsDigit(s[j])) j++; }
                    string num = s.Substring(i, j - i); i = j;
                    return MNode.Txt(num, true);
                }
                if (char.IsLetter(c))
                {
                    int j = i; while (j < s.Length && char.IsLetter(s[j])) j++;
                    string run = s.Substring(i, j - i);
                    bool nextScript = j < s.Length && (s[j] == '_' || s[j] == '^' || IsUniSup(s[j]));
                    if (up) { i = j; return MNode.Txt(run, true); }
                    if (Words.Contains(run) || (run.Length >= 3 && !nextScript) || (run.Length == 2 && char.IsUpper(run[0]) && char.IsLower(run[1]) && nextScript))
                    { i = j; return MNode.Txt(run, true); }
                    // một chữ cái là biến (in nghiêng); dãy chữ ngắn viết liền là tích các biến, chỉ chữ cuối nhận chỉ số
                    var seq = new List<MNode>();
                    for (int k = i; k < j; k++)
                    {
                        char ch = s[k];
                        bool upperGreek = ch >= '\u0391' && ch <= '\u03A9';
                        seq.Add(MNode.Txt(ch.ToString(), upperGreek));
                    }
                    i = j;
                    if (seq.Count == 1) return seq[0];
                    MNode last = seq[seq.Count - 1]; seq.RemoveAt(seq.Count - 1);
                    var wrap = MNode.Seq(seq);
                    wrap.Kids.Add(Scripts(last));
                    return wrap;
                }
                // ký hiệu, toán tử
                i++;
                if (c == '-') c = '−';
                if (c == '*') c = '·';
                return OpNode(c.ToString());
            }

            static MNode OpNode(string t) { var n = MNode.Txt(t, true); n.Op = OpChars.IndexOf(t[0]) >= 0; return n; }

            MNode Command()
            {
                int j = i + 1;
                while (j < s.Length && ((s[j] >= 'a' && s[j] <= 'z') || (s[j] >= 'A' && s[j] <= 'Z'))) j++;
                string cmd = s.Substring(i + 1, j - i - 1);
                if (cmd.Length == 0)
                {
                    char nx = j < s.Length ? s[j] : ' ';
                    i = Math.Min(s.Length, j + 1);
                    if (nx == '%' || nx == '{' || nx == '}' || nx == '_' || nx == '&' || nx == '#') return MNode.Txt(nx.ToString(), true);
                    return null;
                }
                i = j;
                if (cmd == "frac" || cmd == "dfrac" || cmd == "tfrac")
                {
                    MNode a = GroupNode(up), b = GroupNode(up);
                    return MNode.Of(MK.Frac, a, b, null);
                }
                if (cmd == "sqrt")
                {
                    SkipWs();
                    if (i < s.Length && s[i] == '[') { int e = s.IndexOf(']', i); i = e < 0 ? s.Length : e + 1; }
                    return MNode.Of(MK.Rad, GroupNode(up), null, null);
                }
                string accChr = cmd == "hat" ? "\u0302" : cmd == "bar" || cmd == "overline" ? "\u0305" : cmd == "vec" ? "\u20D7" : cmd == "dot" ? "\u0307" : cmd == "tilde" ? "\u0303" : null;
                if (accChr != null) { var a = MNode.Of(MK.Acc, GroupNode(up), null, null); a.T = accChr; return a; }
                if (cmd == "mathrm" || cmd == "text" || cmd == "textrm" || cmd == "rm" || cmd == "operatorname" || cmd == "mathbf")
                {
                    SkipWs(); string g = RawGroup(); if (g == null) return null;
                    return MNode.Txt(StripCmds(g), true);
                }
                if (cmd == "mathit" || cmd == "textit")
                {
                    SkipWs(); string g = RawGroup(); if (g == null) return null;
                    return MNode.Txt(StripCmds(g), false);
                }
                string v; bool italic;
                if (Symbols.TryCmd(cmd, out v, out italic))
                {
                    var n = MNode.Txt(v, up || !italic);
                    n.Op = !italic && !(char.IsLetter(v[0]) && v.Length > 0 && MathParser.IsGreek(v[0])) && OpChars.IndexOf(v[0]) >= 0;
                    return n;
                }
                return MNode.Txt(cmd, true);          // mã chưa biết: giữ chữ, bỏ dấu gạch chéo ngược
            }

            MNode Scripts(MNode atom)
            {
                MNode sub = null, sup = null;
                for (int guard = 0; guard < 4 && i < s.Length; guard++)
                {
                    char c = s[i];
                    if (c == '_' && sub == null) { i++; sub = ReadScript(false); if (sub == null) break; }
                    else if (c == '^' && sup == null) { i++; sup = ReadScript(true); if (sup == null) break; }
                    else if (IsUniSup(c) && sup == null)
                    {
                        var sb = new StringBuilder();
                        while (i < s.Length && IsUniSup(s[i])) { sb.Append(SupA[SupU.IndexOf(s[i])]); i++; }
                        sup = MNode.Txt(sb.ToString().Replace('-', '−'), true);
                    }
                    else if (SubU.IndexOf(c) >= 0 && sub == null)
                    {
                        var sb = new StringBuilder();
                        while (i < s.Length && SubU.IndexOf(s[i]) >= 0) { sb.Append(SubA[SubU.IndexOf(s[i])]); i++; }
                        sub = MNode.Txt(sb.ToString().Replace('-', '−'), true);
                    }
                    else break;
                }
                if (sub != null && sup != null) return MNode.Of(MK.SubSup, atom, sub, sup);
                if (sub != null) return MNode.Of(MK.Sub, atom, sub, null);
                if (sup != null) return MNode.Of(MK.Sup, atom, null, sup);
                return atom;
            }

            MNode ReadScript(bool isSup)
            {
                if (i >= s.Length) return null;
                char c = s[i];
                if (c == '{') { string g = RawGroup(); return new Parser(g, true).Seq(false); }
                if (c == '(')
                {
                    int d = 0, j = i;
                    for (; j < s.Length; j++) { if (s[j] == '(') d++; else if (s[j] == ')') { d--; if (d == 0) break; } }
                    if (j >= s.Length) return null;
                    string inner = s.Substring(i + 1, j - i - 1); i = j + 1;
                    return new Parser(inner, true).Seq(false);
                }
                if (c == '\\')
                {
                    int save = i; MNode a = Command();
                    if (a == null) { i = save; return null; }
                    return a;
                }
                int k = i;
                if (isSup)
                {
                    if (c == '-' || c == '−' || c == '+') k++;
                    if (k < s.Length && char.IsDigit(s[k]))
                    {
                        while (k < s.Length && char.IsDigit(s[k])) k++;
                        if (k + 1 < s.Length && (s[k] == '.' || s[k] == ',') && char.IsDigit(s[k + 1])) { k++; while (k < s.Length && char.IsDigit(s[k])) k++; }
                    }
                    else if (k == i && (char.IsLetter(c) || c == '*' || c == '\'')) k++;
                    else if (k > i && k < s.Length && char.IsLetter(s[k])) k++;
                    else return null;
                }
                else
                {
                    while (k < s.Length && (char.IsLetterOrDigit(s[k]) || s[k] == ',')) k++;
                    while (k > i && s[k - 1] == ',') k--;
                }
                if (k == i) return null;
                string t = s.Substring(i, k - i).Replace('-', '−'); i = k;
                return MNode.Txt(t, true);
            }
        }
    }

    /// <summary>Bộ dựng OMML (Equation của Word) và dạng chữ có định dạng (các đoạn Run) từ cây công thức.</summary>
    public static class Omml
    {
        const string NsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string NsM = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        const string MathFont = "<w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr>";
        const string Ctrl = "<m:ctrlPr><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/><w:i/></w:rPr></m:ctrlPr>";

        // ------------------------------------------------------------------ từ các đoạn Run
        /// <summary>Đoạn Run (chữ thường, chỉ số dưới, chỉ số trên, nghiêng) sang cây công thức. Chỉ số bám vào ký tự đứng ngay trước.</summary>
        public static MNode FromRuns(List<Run> runs, bool chem)
        {
            var kids = new List<MNode>();
            Action<string, bool> addText = delegate(string t, bool up)
            {
                t = t.Replace(" ", "");
                if (t.Length == 0) return;
                if (kids.Count > 0 && kids[kids.Count - 1].K == MK.Text && kids[kids.Count - 1].Up == up && !kids[kids.Count - 1].Sp) kids[kids.Count - 1].T += t;
                else kids.Add(MNode.Txt(t, up));
            };
            foreach (Run r in runs)
            {
                if (r.Kind == 0 || r.Kind == 3) { addText(r.T, r.Kind == 0 || chem); continue; }
                string txt = r.T.Replace(" ", "");
                if (txt.Length == 0) continue;
                MNode body = MNode.Txt(txt, true);
                MNode last = kids.Count > 0 ? kids[kids.Count - 1] : null;
                if (last != null && last.K == MK.Text && last.T.Length > 1)
                {
                    string head = last.T.Substring(0, last.T.Length - 1), ch = last.T.Substring(last.T.Length - 1);
                    last.T = head; last = MNode.Txt(ch, last.Up); kids.Add(last);
                }
                else if (last != null && last.K == MK.Text) { /* một ký tự: dùng luôn làm cơ sở */ }
                if (last != null && last.K == MK.Sub && r.Kind == 2) { kids[kids.Count - 1] = MNode.Of(MK.SubSup, last.A, last.B, body); continue; }
                if (last != null && last.K == MK.Sup && r.Kind == 1) { kids[kids.Count - 1] = MNode.Of(MK.SubSup, last.A, body, last.C); continue; }
                MNode bas = last ?? MNode.Txt("", true);
                if (last != null) kids.RemoveAt(kids.Count - 1);
                kids.Add(r.Kind == 1 ? MNode.Of(MK.Sub, bas, body, null) : MNode.Of(MK.Sup, bas, null, body));
            }
            return MNode.Seq(kids);
        }

        // ------------------------------------------------------------------ cây thành Run (xem trước, phương án dự phòng)
        public static List<Run> ToRuns(MNode n)
        {
            var o = new List<Run>(); Flat(n, o);
            var m = new List<Run>();
            foreach (Run r in o)
            {
                if (r.T.Length == 0) continue;
                if (m.Count > 0 && m[m.Count - 1].Kind == r.Kind) m[m.Count - 1] = new Run(m[m.Count - 1].T + r.T, r.Kind);
                else m.Add(r);
            }
            return m;
        }

        public static string PlainOf(MNode n)
        {
            var sb = new StringBuilder();
            foreach (Run r in ToRuns(n)) sb.Append(r.T);
            return sb.ToString();
        }

        static bool Simple(MNode n)
        {
            if (n == null) return true;
            if (n.K == MK.Text) return true;
            if (n.K == MK.Paren) return true;
            if (n.K == MK.Seq) { var real = n.Kids.Where(k => !k.Sp).ToList(); return real.Count == 1 && Simple(real[0]); }
            if (n.K == MK.Sub || n.K == MK.Sup || n.K == MK.SubSup) return Simple(n.A);
            return false;
        }

        static void Flat(MNode n, List<Run> o)
        {
            if (n == null) return;
            switch (n.K)
            {
                case MK.Text: o.Add(new Run(n.T, n.Up ? 0 : 3)); break;
                case MK.Seq: foreach (MNode k in n.Kids) Flat(k, o); break;
                case MK.Paren: o.Add(new Run("(", 0)); Flat(n.A, o); o.Add(new Run(")", 0)); break;
                case MK.Sub: Flat(n.A, o); o.Add(new Run(PlainOf(n.B), 1)); break;
                case MK.Sup: Flat(n.A, o); o.Add(new Run(PlainOf(n.C), 2)); break;
                case MK.SubSup: Flat(n.A, o); o.Add(new Run(PlainOf(n.B), 1)); o.Add(new Run(PlainOf(n.C), 2)); break;
                case MK.Frac:
                    {
                        MNode a = n.Slash ? n.A : Strip(n.A), b = n.Slash ? n.B : Strip(n.B);
                        bool wa = !n.Slash && !Simple(a), wb = !n.Slash && !Simple(b);
                        if (wa) o.Add(new Run("(", 0)); Flat(a, o); if (wa) o.Add(new Run(")", 0));
                        o.Add(new Run("/", 0));
                        if (wb) o.Add(new Run("(", 0)); Flat(b, o); if (wb) o.Add(new Run(")", 0));
                        break;
                    }
                case MK.Rad: o.Add(new Run("√(", 0)); Flat(n.A, o); o.Add(new Run(")", 0)); break;
                case MK.Acc: Flat(n.A, o); o.Add(new Run(n.T, 0)); break;
            }
        }

        internal static MNode Strip(MNode n)
        {
            if (n == null) return n;
            if (n.K == MK.Paren) return n.A;
            if (n.K == MK.Seq)
            {
                var real = n.Kids.Where(k => !k.Sp).ToList();
                if (real.Count == 1 && real[0].K == MK.Paren) return real[0].A;
            }
            return n;
        }

        public static bool IsComplex(MNode n, bool stacked)
        {
            if (n == null) return false;
            if (n.K == MK.Frac) { if (!n.Slash || stacked) return true; }
            if (n.K == MK.Rad) return true;
            if (IsComplex(n.A, stacked) || IsComplex(n.B, stacked) || IsComplex(n.C, stacked)) return true;
            foreach (MNode k in n.Kids) if (IsComplex(k, stacked)) return true;
            return false;
        }

        // ------------------------------------------------------------------ OMML
        static string Esc(string t) { return t.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;"); }

        public static string Build(MNode n, bool stacked)
        {
            var sb = new StringBuilder(); Emit(sb, n, stacked); return sb.ToString();
        }

        static void Group(StringBuilder sb, string tag, MNode n, bool stacked)
        {
            sb.Append("<m:").Append(tag).Append('>'); Emit(sb, n, stacked); sb.Append("</m:").Append(tag).Append('>');
        }

        static void Emit(StringBuilder sb, MNode n, bool stacked)
        {
            if (n == null) return;
            switch (n.K)
            {
                case MK.Text:
                    if (n.Sp || n.T.Length == 0) return;
                    sb.Append("<m:r>");
                    if (n.Up) sb.Append("<m:rPr><m:sty m:val=\"p\"/></m:rPr>");
                    sb.Append(MathFont).Append("<m:t xml:space=\"preserve\">").Append(Esc(n.T)).Append("</m:t></m:r>");
                    break;
                case MK.Seq: foreach (MNode k in n.Kids) Emit(sb, k, stacked); break;
                case MK.Paren:
                    sb.Append("<m:d><m:dPr><m:begChr m:val=\"(\"/><m:endChr m:val=\")\"/>").Append(Ctrl).Append("</m:dPr>");
                    Group(sb, "e", n.A, stacked); sb.Append("</m:d>");
                    break;
                case MK.Sub:
                    sb.Append("<m:sSub><m:sSubPr>").Append(Ctrl).Append("</m:sSubPr>");
                    Group(sb, "e", n.A, stacked); Group(sb, "sub", n.B, stacked); sb.Append("</m:sSub>");
                    break;
                case MK.Sup:
                    sb.Append("<m:sSup><m:sSupPr>").Append(Ctrl).Append("</m:sSupPr>");
                    Group(sb, "e", n.A, stacked); Group(sb, "sup", n.C, stacked); sb.Append("</m:sSup>");
                    break;
                case MK.SubSup:
                    sb.Append("<m:sSubSup><m:sSubSupPr>").Append(Ctrl).Append("</m:sSubSupPr>");
                    Group(sb, "e", n.A, stacked); Group(sb, "sub", n.B, stacked); Group(sb, "sup", n.C, stacked); sb.Append("</m:sSubSup>");
                    break;
                case MK.Frac:
                    if (n.Slash && !stacked)
                    {
                        Emit(sb, n.A, stacked);
                        sb.Append("<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr>").Append(MathFont).Append("<m:t>/</m:t></m:r>");
                        Emit(sb, n.B, stacked);
                    }
                    else
                    {
                        sb.Append("<m:f><m:fPr><m:type m:val=\"").Append(stacked ? "bar" : "lin").Append("\"/>").Append(Ctrl).Append("</m:fPr>");
                        Group(sb, "num", Strip(n.A), stacked); Group(sb, "den", Strip(n.B), stacked); sb.Append("</m:f>");
                    }
                    break;
                case MK.Rad:
                    sb.Append("<m:rad><m:radPr><m:degHide m:val=\"1\"/>").Append(Ctrl).Append("</m:radPr><m:deg/>");
                    Group(sb, "e", Strip(n.A), stacked); sb.Append("</m:rad>");
                    break;
                case MK.Acc:
                    sb.Append("<m:acc><m:accPr><m:chr m:val=\"").Append(n.T).Append("\"/>").Append(Ctrl).Append("</m:accPr>");
                    Group(sb, "e", n.A, stacked); sb.Append("</m:acc>");
                    break;
            }
        }

        // ------------------------------------------------------------------ đọc OMML có sẵn trong Word (Equation → cây công thức)
        /// <summary>Đọc công thức đầu tiên (m:oMath) trong một chuỗi XML bất kỳ: gói Flat OPC của Range.WordOpenXML, hoặc chỉ đoạn OMML. Không đọc được thì trả về null.</summary>
        public static MNode FromOmml(string xml)
        {
            if (string.IsNullOrEmpty(xml)) return null;
            try
            {
                var d = new XmlDocument(); d.XmlResolver = null; d.LoadXml(xml);
                XmlNode root = null;
                foreach (XmlNode n in d.GetElementsByTagName("*"))
                    if (n.LocalName == "oMath" && n.NamespaceURI == NsM) { root = n; break; }
                if (root == null) return null;
                MNode res = ReadKids(root);
                return res;
            }
            catch (Exception ex) { Log.W("FromOmml loi: " + ex.Message); return null; }
        }

        static bool IsM(XmlNode n, string name) { return n.NodeType == XmlNodeType.Element && n.NamespaceURI == NsM && n.LocalName == name; }
        static XmlNode Child(XmlNode n, string name) { foreach (XmlNode c in n.ChildNodes) if (IsM(c, name)) return c; return null; }
        static string Attr(XmlNode n, string name)
        {
            if (n == null) return null;
            foreach (XmlAttribute a in n.Attributes) if (a.LocalName == name) return a.Value;
            return null;
        }
        static string ChrOf(XmlNode pr, string name, string dflt)
        {
            XmlNode c = pr == null ? null : Child(pr, name);
            if (c == null) return dflt;
            string v = Attr(c, "val");
            return v == null ? dflt : v;
        }

        static MNode ReadKids(XmlNode parent)
        {
            var kids = new List<MNode>();
            foreach (XmlNode c in parent.ChildNodes)
            {
                if (c.NodeType != XmlNodeType.Element) continue;
                if (c.NamespaceURI != NsM) continue;
                string ln = c.LocalName;
                if (ln.EndsWith("Pr")) continue;                     // thuộc tính: rPr, fPr, dPr, ctrlPr...
                ReadOne(c, kids);
            }
            return kids.Count == 1 && kids[0].K == MK.Seq ? kids[0] : MNode.Seq(kids);
        }

        static MNode Part(XmlNode n, string name) { XmlNode c = n == null ? null : Child(n, name); return c == null ? MNode.Seq(new List<MNode>()) : ReadKids(c); }

        static void ReadOne(XmlNode c, List<MNode> o)
        {
            switch (c.LocalName)
            {
                case "r":
                    {
                        XmlNode pr = Child(c, "rPr");
                        string sty = pr == null ? null : Attr(Child(pr, "sty"), "val");
                        bool nor = pr != null && Child(pr, "nor") != null;
                        bool up = nor || sty == "p" || sty == "b";
                        var sb = new StringBuilder();
                        foreach (XmlNode t in c.ChildNodes) if (IsM(t, "t")) sb.Append(t.InnerText);
                        string txt = sb.ToString();
                        if (txt.Length == 0) return;
                        // chữ cái Latin và Hy Lạp thường trong Equation mặc định in nghiêng; số, dấu, chữ Hy Lạp hoa thì đứng thẳng
                        int i = 0;
                        while (i < txt.Length)
                        {
                            bool it = !up && IsMathItalic(txt[i]);
                            int j = i + 1;
                            while (j < txt.Length && (!up && IsMathItalic(txt[j])) == it) j++;
                            o.Add(MNode.Txt(txt.Substring(i, j - i), !it));
                            i = j;
                        }
                        return;
                    }
                case "sSub": o.Add(MNode.Of(MK.Sub, Part(c, "e"), Part(c, "sub"), null)); return;
                case "sSup": o.Add(MNode.Of(MK.Sup, Part(c, "e"), null, Part(c, "sup"))); return;
                case "sSubSup": o.Add(MNode.Of(MK.SubSup, Part(c, "e"), Part(c, "sub"), Part(c, "sup"))); return;
                case "f":
                    {
                        string type = ChrOf(Child(c, "fPr"), "type", "bar");
                        var f = MNode.Of(MK.Frac, Part(c, "num"), Part(c, "den"), null); f.Slash = type == "lin";
                        o.Add(f); return;
                    }
                case "d":
                    {
                        XmlNode pr = Child(c, "dPr");
                        string beg = ChrOf(pr, "begChr", "("), end = ChrOf(pr, "endChr", ")"), sep = ChrOf(pr, "sepChr", "|");
                        var inner = new List<MNode>(); bool first = true;
                        foreach (XmlNode e in c.ChildNodes)
                        {
                            if (!IsM(e, "e")) continue;
                            if (!first) inner.Add(MNode.Txt(sep, true));
                            inner.Add(ReadKids(e)); first = false;
                        }
                        MNode body = MNode.Seq(inner);
                        if (beg == "(" && end == ")") o.Add(MNode.Of(MK.Paren, body, null, null));
                        else { o.Add(MNode.Txt(beg, true)); o.Add(body); o.Add(MNode.Txt(end, true)); }
                        return;
                    }
                case "rad": o.Add(MNode.Of(MK.Rad, Part(c, "e"), null, null)); return;
                case "acc": { var a = MNode.Of(MK.Acc, Part(c, "e"), null, null); a.T = ChrOf(Child(c, "accPr"), "chr", "\u0302"); o.Add(a); return; }
                case "nary":
                    {
                        string chr = ChrOf(Child(c, "naryPr"), "chr", "∫");
                        o.Add(MNode.Txt(chr, true));
                        MNode lo = Part(c, "sub"), hi = Part(c, "sup");
                        if (lo.Kids.Count > 0) o.Add(MNode.Of(MK.Sub, MNode.Seq(new List<MNode>()), lo, null));
                        if (hi.Kids.Count > 0) o.Add(MNode.Of(MK.Sup, MNode.Seq(new List<MNode>()), null, hi));
                        o.Add(Part(c, "e")); return;
                    }
                default:
                    // thành phần chưa hỗ trợ (hàm, thanh ngang, ma trận...): lấy nội dung chữ của các phần con theo thứ tự
                    o.Add(ReadKids(c)); return;
            }
        }

        static bool IsMathItalic(char ch)
        {
            return (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '\u03B1' && ch <= '\u03C9');
        }

        /// <summary>Gói OMML thành tài liệu Flat OPC để Range.InsertXML chèn vào Word. display = công thức riêng một dòng.</summary>
        public static string Package(string omml, bool display)
        {
            string body = display ? "<m:oMathPara><m:oMath>" + omml + "</m:oMath></m:oMathPara>" : "<m:oMath>" + omml + "</m:oMath>";
            return "<?xml version=\"1.0\" standalone=\"yes\"?>"
                + "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">"
                + "<pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"
                + "</pkg:xmlData></pkg:part>"
                + "<pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData>"
                + "<w:document xmlns:w=\"" + NsW + "\" xmlns:m=\"" + NsM + "\"><w:body><w:p>" + body + "</w:p></w:body></w:document>"
                + "</pkg:xmlData></pkg:part></pkg:package>";
        }
    }
}
