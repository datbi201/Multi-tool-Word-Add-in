// Trợ lý Kỹ thuật - Rev0.5: tìm cả đoạn công thức gõ kiểu LaTeX hoặc chữ phẳng trong một đoạn văn: Re_{cyl} = \frac{a}{b} = ..., f = 0,0791 · (7.725)^{-0,3} ≈ ...
// Chỉ làm việc trên chuỗi, không gọi Word.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    public sealed class MathSpan { public int S, E; public bool Leak, HasFrac; }

    public static class MathSpans
    {
        sealed class Tk { public int S, E; public char K; public bool Script, Frac, Eq; }

        static readonly HashSet<string> KnownWords = new HashSet<string>(new[] { "ln", "log", "exp", "sin", "cos", "tan", "cot", "max", "min", "lim", "Re", "Nu", "Pr", "Sc", "Sh", "Ra", "Gr", "det", "arg" });
        const string OpChars = "=≈≠+-−–·×*/<>≤≥±∓→⇌←↔⇒⇔∝∼";
        const string EqChars = "=≈≠<>≤≥∝";

        static readonly Regex GreekComma = new Regex(@"\G[\u03B1-\u03C9\u0391-\u03A9][a-z]{1,4},\s?[A-Za-z]{2,4}(?=\s*[=≈])");
        static bool IsAscii(char c) { return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'); }

        static int GroupEnd(string t, int k)
        {
            if (k >= t.Length || t[k] != '{') return -1;
            int d = 0;
            for (int j = k; j < t.Length; j++) { if (t[j] == '{') d++; else if (t[j] == '}') { d--; if (d == 0) return j + 1; } }
            return -1;
        }

        static int ParenEnd(string t, int k)
        {
            int d = 0;
            for (int j = k; j < t.Length && j < k + 300; j++) { if (t[j] == '(') d++; else if (t[j] == ')') { d--; if (d == 0) return j + 1; } }
            return -1;
        }

        static int ScriptEnd(string t, int p, bool sup)
        {
            int n = t.Length;
            if (p >= n) return -1;
            char c = t[p];
            if (c == '{') return GroupEnd(t, p);
            if (c == '(') return ParenEnd(t, p);
            if (c == '\\') { int j = p + 1; while (j < n && IsAscii(t[j])) j++; return j > p + 1 ? j : -1; }
            int k = p;
            if (sup)
            {
                if (c == '-' || c == '−' || c == '+') k++;
                if (k < n && char.IsDigit(t[k]))
                {
                    while (k < n && char.IsDigit(t[k])) k++;
                    if (k + 1 < n && (t[k] == '.' || t[k] == ',') && char.IsDigit(t[k + 1])) { k++; while (k < n && char.IsDigit(t[k])) k++; }
                    return k;
                }
                if (k == p && (char.IsLetter(c) || c == '*')) return p + 1;
                if (k > p && k < n && char.IsLetter(t[k])) return k + 1;
                return -1;
            }
            while (k < n && (char.IsLetterOrDigit(t[k]) || t[k] == ',')) k++;
            while (k > p && t[k - 1] == ',') k--;
            return k > p ? k : -1;
        }

        static void Scripts(string t, ref int i, Tk tk)
        {
            while (i < t.Length)
            {
                char c = t[i];
                if (c == '_') { int e = ScriptEnd(t, i + 1, false); if (e < 0) break; i = e; tk.Script = true; }
                else if (c == '^') { int e = ScriptEnd(t, i + 1, true); if (e < 0) break; i = e; tk.Script = true; }
                else if (MathParser.IsUniSup(c)) { while (i < t.Length && MathParser.IsUniSup(t[i])) i++; tk.Script = true; }
                else break;
            }
            tk.E = i;
        }

        static List<Tk> Lex(string t)
        {
            var r = new List<Tk>(); int i = 0, n = t.Length;
            while (i < n)
            {
                char c = t[i]; int s = i;
                if (char.IsWhiteSpace(c)) { while (i < n && char.IsWhiteSpace(t[i])) i++; r.Add(new Tk { S = s, E = i, K = 'S' }); continue; }
                if (c == '\\')
                {
                    int j = i + 1; while (j < n && IsAscii(t[j])) j++;
                    string cmd = t.Substring(i + 1, j - i - 1);
                    if (cmd.Length == 0)
                    {
                        if (j < n && ",;:! ".IndexOf(t[j]) >= 0) { i = j + 1; r.Add(new Tk { S = s, E = i, K = 'S' }); continue; }
                        i = Math.Min(n, j + 1); r.Add(new Tk { S = s, E = i, K = 'X' }); continue;
                    }
                    if (cmd == "frac" || cmd == "dfrac" || cmd == "tfrac")
                    {
                        int k = j; while (k < n && char.IsWhiteSpace(t[k])) k++;
                        int e1 = GroupEnd(t, k);
                        int e2 = -1;
                        if (e1 > 0) { int k2 = e1; while (k2 < n && char.IsWhiteSpace(t[k2])) k2++; e2 = GroupEnd(t, k2); }
                        if (e1 > 0 && e2 > 0) { var tk = new Tk { S = s, K = 'F', Frac = true }; i = e2; Scripts(t, ref i, tk); r.Add(tk); continue; }
                        i = j; r.Add(new Tk { S = s, E = j, K = 'W' }); continue;
                    }
                    if (cmd == "sqrt")
                    {
                        int k = j; if (k < n && t[k] == '[') { int e = t.IndexOf(']', k); if (e > 0) k = e + 1; }
                        int e1 = GroupEnd(t, k);
                        if (e1 > 0) { var tk = new Tk { S = s, K = 'F', Frac = true }; i = e1; Scripts(t, ref i, tk); r.Add(tk); continue; }
                        i = j; r.Add(new Tk { S = s, E = j, K = 'W' }); continue;
                    }
                    if (cmd == "left" || cmd == "right" || cmd == "big" || cmd == "Big" || cmd == "bigg" || cmd == "Bigg") { i = j; r.Add(new Tk { S = s, E = j, K = 'S' }); continue; }
                    if (cmd == "mathrm" || cmd == "text" || cmd == "textrm" || cmd == "rm" || cmd == "operatorname" || cmd == "mathbf" || cmd == "mathit" || cmd == "hat" || cmd == "bar" || cmd == "vec" || cmd == "overline" || cmd == "dot" || cmd == "tilde")
                    {
                        int k = j; while (k < n && char.IsWhiteSpace(t[k])) k++;
                        int e = GroupEnd(t, k);
                        if (e > 0) { var tk = new Tk { S = s, K = 'C' }; i = e; Scripts(t, ref i, tk); r.Add(tk); continue; }
                        i = j; r.Add(new Tk { S = s, E = j, K = 'W' }); continue;
                    }
                    if (Symbols.Known(cmd))
                    {
                        string v; bool it; Symbols.TryCmd(cmd, out v, out it);
                        if (v.Length > 0 && OpChars.IndexOf(v[0]) >= 0 && !char.IsLetter(v[0])) { i = j; r.Add(new Tk { S = s, E = j, K = 'O', Eq = EqChars.IndexOf(v[0]) >= 0 }); continue; }
                        var tk = new Tk { S = s, K = 'C' }; i = j; Scripts(t, ref i, tk); r.Add(tk); continue;
                    }
                    i = j; r.Add(new Tk { S = s, E = j, K = 'W' }); continue;
                }
                if (char.IsDigit(c))
                {
                    int j = i; while (j < n && char.IsDigit(t[j])) j++;
                    while (j + 1 < n && (t[j] == '.' || t[j] == ',') && char.IsDigit(t[j + 1])) { j++; while (j < n && char.IsDigit(t[j])) j++; }
                    var tk = new Tk { S = s, K = 'N' }; i = j; Scripts(t, ref i, tk); r.Add(tk); continue;
                }
                if (MathParser.IsGreek(c))
                {
                    Match gm = GreekComma.Match(t, i);
                    if (gm.Success) { i += gm.Length; r.Add(new Tk { S = s, E = i, K = 'I', Script = true }); continue; }
                }
                if (char.IsLetter(c))
                {
                    int j = i; while (j < n && char.IsLetter(t[j])) j++;
                    string run = t.Substring(i, j - i);
                    bool next = j < n && (t[j] == '_' || t[j] == '^' || MathParser.IsUniSup(t[j]));
                    bool allUpper = run.All(ch => ch >= 'A' && ch <= 'Z');
                    bool ok;
                    if (next) ok = run.Length <= 2 || (allUpper && run.Length <= 4);
                    else ok = (run.Length == 1 && (IsAscii(c) || MathParser.IsGreek(c))) || KnownWords.Contains(run);
                    if (!ok) { i = j; r.Add(new Tk { S = s, E = j, K = 'W' }); continue; }
                    var tk = new Tk { S = s, K = 'I' }; i = j; Scripts(t, ref i, tk); r.Add(tk); continue;
                }
                if (c == '(' || c == '{')
                {
                    int e = c == '(' ? ParenEnd(t, i) : GroupEnd(t, i);
                    if (e > 0 && Mathy(t.Substring(i + 1, e - i - 2)))
                    {
                        var tk = new Tk { S = s, K = 'P' }; i = e; Scripts(t, ref i, tk); r.Add(tk); continue;
                    }
                    i++; r.Add(new Tk { S = s, E = i, K = 'X' }); continue;
                }
                if (OpChars.IndexOf(c) >= 0) { i++; r.Add(new Tk { S = s, E = i, K = 'O', Eq = EqChars.IndexOf(c) >= 0 }); continue; }
                i++; r.Add(new Tk { S = s, E = i, K = 'X' });
            }
            return r;
        }

        /// <summary>Nội dung trong ngoặc chỉ gồm ký hiệu toán (biến, số, toán tử, mã LaTeX), không có chữ thường là từ.</summary>
        public static bool Mathy(string inner)
        {
            if (inner.Trim().Length == 0) return false;
            foreach (Tk k in Lex(inner))
            {
                if (k.K == 'W') return false;
                if (k.K == 'X' && ",.;|".IndexOf(inner[k.S]) < 0) return false;
            }
            return true;
        }

        /// <summary>Tìm các đoạn công thức: chứa phân số hoặc căn LaTeX, hoặc có chỉ số/số mũ cùng dấu so sánh (=, ≈...) mà các ký hiệu viết liền một mạch.</summary>
        public static List<MathSpan> Find(string txt)
        {
            var res = new List<MathSpan>();
            if (txt.IndexOf('\\') < 0 && txt.IndexOf('_') < 0 && txt.IndexOf('^') < 0 && !txt.Any(MathParser.IsUniSup)) return res;
            List<Tk> tks = Lex(txt);
            int a = 0;
            while (a < tks.Count)
            {
                if ("SOIPNCF".IndexOf(tks[a].K) < 0) { a++; continue; }
                int b = a; while (b < tks.Count && "SOIPNCF".IndexOf(tks[b].K) >= 0) b++;
                int lo = a, hi = b - 1;
                while (lo <= hi && (tks[lo].K == 'S' || tks[lo].K == 'O')) lo++;
                while (hi >= lo && (tks[hi].K == 'S' || tks[hi].K == 'O')) hi--;
                if (lo <= hi)
                {
                    bool frac = false, script = false, eq = false; int operands = 0;
                    for (int k = lo; k <= hi; k++)
                    {
                        Tk x = tks[k];
                        if (x.K == 'F') frac = true;
                        if (x.Script) script = true;
                        if (x.K == 'O' && x.Eq) eq = true;
                        if (x.K != 'S' && x.K != 'O') operands++;
                    }
                    int s = tks[lo].S, e = tks[hi].E;
                    string seg = txt.Substring(s, e - s);
                    bool leak = seg.IndexOf('\\') >= 0 || seg.IndexOf('{') >= 0;
                    if (frac || (script && eq && operands >= 2) || (leak && operands >= 2 && eq))
                        res.Add(new MathSpan { S = s, E = e, Leak = leak, HasFrac = frac });
                }
                a = b;
            }
            return res;
        }
    }
}
