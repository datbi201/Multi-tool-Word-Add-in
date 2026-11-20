// Trợ lý Kỹ thuật - Rev0.2: engine hóa học (đọc công thức, hóa trị, điện tích, cân bằng phương trình).
// Chuyển từ bản demo JavaScript, chỉ dùng C# 5. Không phụ thuộc Word nên kiểm thử được riêng.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    /// <summary>Thành phần nguyên tố của một chất, giữ thứ tự xuất hiện.</summary>
    public sealed class Comp
    {
        public List<string> Keys = new List<string>();
        public Dictionary<string, int> N = new Dictionary<string, int>();
        public void Add(string sym, int n)
        {
            if (!N.ContainsKey(sym)) { N[sym] = 0; Keys.Add(sym); }
            N[sym] += n;
        }
        public int Get(string sym) { int v; return N.TryGetValue(sym, out v) ? v : 0; }
    }

    public sealed class ChemFlag { public string Type; public string Msg; public string Sug; public List<string> Alts = new List<string>(); }

    public sealed class Species
    {
        public string Raw, Out, Coef = "";
        public string TypedOut = "";                 // công thức đã chuẩn hóa chữ hoa nhưng chưa đổi theo gợi ý
        public List<string> Alts = new List<string>();   // các công thức hợp hóa trị gần với cách gõ (tốt nhất đứng đầu)
        public Comp Comp;
        public int Charge;
        public List<ChemFlag> Flags = new List<ChemFlag>();
        public bool Unknown;
    }

    public sealed class Rxn
    {
        public string Std, BalStr;
        public bool? Balanced;
        public int NL;                                   // số chất ở vế trái
        public string Arrow = "->";
        public bool HasCoef;                             // người dùng đã gõ hệ số ở ít nhất một chất
        public List<ChemFlag> Flags = new List<ChemFlag>();
        public bool Unknown;
        public List<Species> Species = new List<Species>();
        public List<RxnAlt> Alts = new List<RxnAlt>();   // các cách đọc khác khi một chất có nhiều gợi ý (FeCl2 hay FeCl3)
    }

    public sealed class RxnAlt { public string Label, Std, BalStr; }

    public static class Chem
    {
        static readonly HashSet<string> ELS = new HashSet<string>(
            ("H He Li Be B C N O F Ne Na Mg Al Si P S Cl Ar K Ca Sc Ti V Cr Mn Fe Co Ni Cu Zn Ga Ge As Se Br Kr Rb Sr Y Zr Nb Mo Tc Ru Rh Pd Ag Cd In Sn Sb Te I Xe Cs Ba La Ce Hf Ta W Re Os Ir Pt Au Hg Tl Pb Bi Po At Rn Fr Ra Ac Th Pa U").Split(' '));

        static readonly Dictionary<string, int[]> OX = BuildOx();
        static Dictionary<string, int[]> BuildOx()
        {
            var d = new Dictionary<string, int[]>();
            d["H"] = new[] { 1, -1 }; d["Li"] = new[] { 1 }; d["Na"] = new[] { 1 }; d["K"] = new[] { 1 }; d["Rb"] = new[] { 1 }; d["Cs"] = new[] { 1 };
            d["Mg"] = new[] { 2 }; d["Ca"] = new[] { 2 }; d["Sr"] = new[] { 2 }; d["Ba"] = new[] { 2 }; d["Zn"] = new[] { 2 }; d["Cd"] = new[] { 2 };
            d["Al"] = new[] { 3 }; d["B"] = new[] { 3 }; d["Fe"] = new[] { 2, 3 }; d["Cu"] = new[] { 1, 2 }; d["Ni"] = new[] { 2, 3 }; d["Co"] = new[] { 2, 3 };
            d["Mn"] = new[] { 2, 3, 4, 6, 7 }; d["Cr"] = new[] { 2, 3, 6 }; d["Ti"] = new[] { 2, 3, 4 }; d["Pb"] = new[] { 2, 4 }; d["Sn"] = new[] { 2, 4 };
            d["Ag"] = new[] { 1 }; d["Hg"] = new[] { 1, 2 }; d["Au"] = new[] { 1, 3 }; d["Pt"] = new[] { 2, 4 };
            d["O"] = new[] { -2, -1 }; d["F"] = new[] { -1 }; d["Cl"] = new[] { -1, 1, 3, 5, 7 }; d["Br"] = new[] { -1, 1, 5 }; d["I"] = new[] { -1, 1, 5, 7 };
            d["S"] = new[] { -2, 2, 4, 6 }; d["N"] = new[] { -3, 1, 2, 3, 4, 5 }; d["P"] = new[] { -3, 3, 5 };
            d["C"] = new[] { -4, -3, -2, -1, 0, 1, 2, 3, 4 }; d["Si"] = new[] { 4, -4 }; d["Se"] = new[] { -2, 4, 6 }; d["As"] = new[] { -3, 3, 5 };
            d["Mo"] = new[] { 2, 3, 4, 5, 6 }; d["V"] = new[] { 2, 3, 4, 5 }; d["W"] = new[] { 4, 6 }; d["Zr"] = new[] { 4 };
            return d;
        }
        static readonly Dictionary<string, int> ELEMENTAL = new Dictionary<string, int> { { "H", 2 }, { "N", 2 }, { "O", 2 }, { "F", 2 }, { "Cl", 2 }, { "Br", 2 }, { "I", 2 } };
        public static readonly HashSet<string> COMMON = new HashSet<string>(new[] { "NaCl", "HCl", "NaOH", "KOH", "CaO", "CO", "FeS", "HF", "KCl", "CaCl2", "MgO", "NaF", "NaHCO3", "HBr", "HI" });
        static readonly Dictionary<string, string> ALIAS = new Dictionary<string, string> {
            { "OH2", "H2O" }, { "SH2", "H2S" }, { "ClH", "HCl" }, { "FH", "HF" }, { "BrH", "HBr" }, { "ClNa", "NaCl" }, { "HOH", "H2O" },
            { "O2C", "CO2" }, { "HNaO", "NaOH" }, { "O4SH2", "H2SO4" }, { "SO4H2", "H2SO4" } };

        static readonly HashSet<string> MixedOk = new HashSet<string>(new[] { "Fe3O4", "Pb3O4", "Mn3O4", "Co3O4", "FeS2", "Fe3S4", "Fe7S8", "Fe9S10", "CuFeS2", "Fe2O3", "Cu2S", "MoS2", "N2H4", "N2O", "H2O2", "Na2O2", "CaO2", "BaO2", "K2O2", "KO2" });
        static readonly HashSet<string> OrganicEls = new HashSet<string>(new[] { "C", "H", "O", "N", "S", "P", "F", "Cl", "Br", "I" });

        public static long Gcd(long a, long b) { a = Math.Abs(a); b = Math.Abs(b); while (b != 0) { long t = a % b; a = b; b = t; } return a; }

        // ---------------------------------------------------------------- đọc công thức
        sealed class FParser
        {
            readonly string s; int i;
            public FParser(string str) { s = str; }
            int Num()
            {
                int st = i;
                while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
                return i > st ? int.Parse(s.Substring(st, i - st)) : 1;
            }
            Comp Group()
            {
                var c = new Comp();
                while (i < s.Length && s[i] != ')')
                {
                    if (s[i] == '(')
                    {
                        i++;
                        Comp g = Group();
                        if (i >= s.Length || s[i] != ')') throw new FormatException();
                        i++;
                        int m = Num();
                        foreach (string k in g.Keys) c.Add(k, g.N[k] * m);
                    }
                    else
                    {
                        if (!(s[i] >= 'A' && s[i] <= 'Z')) throw new FormatException();
                        string sym = s[i].ToString();
                        if (i + 1 < s.Length && s[i + 1] >= 'a' && s[i + 1] <= 'z') sym += s[i + 1];
                        if (!ELS.Contains(sym))
                        {
                            if (sym.Length == 2 && ELS.Contains(sym.Substring(0, 1))) sym = sym.Substring(0, 1);
                            else throw new FormatException();
                        }
                        i += sym.Length;
                        c.Add(sym, Num());
                    }
                }
                return c;
            }
            public Comp Parse()
            {
                try { Comp c = Group(); return (i == s.Length && c.Keys.Count > 0) ? c : null; }
                catch (FormatException) { return null; }
            }
        }
        public static Comp ParseF(string s) { return new FParser(s).Parse(); }

        /// <summary>true = hợp hóa trị, false = không, null = có nguyên tố chưa có dữ liệu.</summary>
        public static bool? Valid(Comp comp, int charge)
        {
            var els = comp.Keys;
            foreach (string e in els) if (!OX.ContainsKey(e)) return null;
            if (charge == 0 && MixedOk.Contains(FStr(comp))) return true;                       // Fe3O4, FeS2...: hóa trị hỗn hợp, hợp chất có thật
            if (charge == 0 && els.Contains("C") && els.Contains("H") && els.All(x => OrganicEls.Contains(x))) return true;   // hợp chất hữu cơ (C3H8, C2H5OH): không kiểm hóa trị kiểu vô cơ
            if (charge == 0 && els.Count == 1)
            {
                string e = els[0]; int n = comp.N[e], v;
                if (n == 1 || (ELEMENTAL.TryGetValue(e, out v) && v == n) || (e == "O" && n == 3) || (e == "P" && n == 4) || (e == "S" && n == 8)) return true;
            }
            bool hPlus = els.Contains("H") && els.Any(x => HPlusWith.Contains(x));      // H đi với phi kim: chỉ +1 (CaOH không được coi là hợp lệ vì H -1)
            return Rec(comp, els, 0, 0, charge, hPlus);
        }
        static readonly HashSet<string> HPlusWith = new HashSet<string>(new[] { "O", "F", "Cl", "Br", "I", "S", "N", "Se", "P", "As" });
        static bool Rec(Comp comp, List<string> els, int k, int sum, int charge, bool hPlus)
        {
            if (k == els.Count) return sum == charge;
            string e = els[k];
            foreach (int o in OX[e])
            {
                if (hPlus && e == "H" && o < 0) continue;
                if (Rec(comp, els, k + 1, sum + o * comp.N[e], charge, hPlus)) return true;
            }
            return false;
        }

        /// <summary>Công thức dạng cation + anion hay gặp mà số nguyên tử không cân bằng điện tích (CaOH, NH4SO4): sai, dù tổ hợp số oxi hóa lạ nào đó khớp.</summary>
        static bool IonicBad(string s)
        {
            if (s.IndexOf('^') >= 0) return false;
            Match cm = CatPre.Match(s);
            if (!cm.Success) return false;
            string cat = cm.Groups["cat"].Value, cn = cm.Groups["cn"].Value, rest = s.Substring(cm.Length);
            int[] qs = null;
            if (cat == "NH4") qs = new[] { 1 };
            else if (ELS.Contains(cat) && !NonMetal.Contains(cat) && OX.ContainsKey(cat)) qs = OX[cat].Where(v => v > 0).ToArray();
            if (qs == null || qs.Length == 0 || rest.Length == 0) return false;
            foreach (Ion a in Anions)
            {
                Match am = Regex.Match(rest, @"^(?:\(" + a.F + @"\)(?<n>\d*)|" + a.F + @"(?<n2>\d*))$");
                if (!am.Success) continue;
                string ns = am.Groups["n"].Success && am.Groups["n"].Value.Length > 0 ? am.Groups["n"].Value : am.Groups["n2"].Value;
                if (!rest.StartsWith("(") && ns.Length > 0 && char.IsDigit(a.F[a.F.Length - 1])) return false;
                int nc = cn.Length > 0 ? int.Parse(cn) : 1, na = ns.Length > 0 ? int.Parse(ns) : 1;
                foreach (int q in qs) if ((long)q * nc == (long)a.Q * na) return false;
                return true;
            }
            return false;
        }

        public static string FStr(Comp c)
        {
            var sb = new StringBuilder();
            foreach (string k in c.Keys) { sb.Append(k); if (c.N[k] > 1) sb.Append(c.N[k]); }
            return sb.ToString();
        }

        // ---------------------------------------------------------------- gợi ý sửa công thức sai hóa trị
        sealed class Ion { public string F; public int Q; public Ion(string f, int q) { F = f; Q = q; } }
        // anion hay gặp (điện tích là độ lớn). Sắp xếp theo độ dài giảm dần để khớp phần dài nhất trước.
        static readonly Ion[] Anions = {
            new Ion("H2PO4", 1), new Ion("HPO4", 2), new Ion("Cr2O7", 2), new Ion("S2O3", 2), new Ion("HCO3", 1), new Ion("HSO4", 1), new Ion("HSO3", 1),
            new Ion("ClO4", 1), new Ion("ClO3", 1), new Ion("MnO4", 1), new Ion("CrO4", 2), new Ion("SiO3", 2), new Ion("C2O4", 2),
            new Ion("PO4", 3), new Ion("SO4", 2), new Ion("SO3", 2), new Ion("NO3", 1), new Ion("NO2", 1), new Ion("CO3", 2),
            new Ion("O2", 2), new Ion("S2", 2), new Ion("ClO", 1), new Ion("OH", 1), new Ion("CN", 1), new Ion("HS", 1), new Ion("Cl", 1), new Ion("Br", 1), new Ion("F", 1), new Ion("I", 1), new Ion("O", 2), new Ion("S", 2) };
        static readonly HashSet<string> NonMetal = new HashSet<string>(new[] { "C", "N", "O", "F", "Cl", "Br", "I", "S", "Se", "P", "As", "Si", "B" });
        static readonly Regex CatPre = new Regex(@"^(?<cat>NH4|[A-Z][a-z]?)(?<cn>\d*)", RegexOptions.Compiled);

        static string AnionStr(Ion a, int n)
        {
            bool poly = a.F.Length > 2 || Regex.IsMatch(a.F, @"\d") || a.F == "OH" || a.F == "HS" || a.F == "CN";
            if (n == 1) return a.F;
            return poly ? "(" + a.F + ")" + n : a.F + n;
        }

        /// <summary>Gợi ý công thức hợp hóa trị gần nhất với cách gõ (tốt nhất đứng đầu, tối đa 3). Trả về danh sách rỗng nếu không có gợi ý chắc chắn.</summary>
        public static List<string> Suggest(string s, Comp comp)
        {
            var found = new List<KeyValuePair<string, int>>();
            if (s.IndexOf('^') >= 0) return new List<string>();
            // 1. tách thành cation + anion hay gặp: AlSO4 -> Al2(SO4)3, H3SO4 -> H2SO4, FeCl -> FeCl2 hoặc FeCl3
            Match cm = CatPre.Match(s);
            if (cm.Success)
            {
                string cat = cm.Groups["cat"].Value; string cn = cm.Groups["cn"].Value;
                string rest = s.Substring(cm.Length);
                int[] qs = null;
                if (cat == "NH4") qs = new[] { 1 };
                else if (ELS.Contains(cat) && !NonMetal.Contains(cat) && OX.ContainsKey(cat)) qs = OX[cat].Where(v => v > 0).ToArray();
                if (qs != null && qs.Length > 0 && rest.Length > 0)
                {
                    foreach (Ion a in Anions)
                    {
                        Match am = Regex.Match(rest, @"^(?:\(" + a.F + @"\)(?<n>\d*)|" + a.F + @"(?<n2>\d*))$");
                        if (!am.Success) continue;
                        string ns = am.Groups["n"].Success && am.Groups["n"].Value.Length > 0 ? am.Groups["n"].Value : am.Groups["n2"].Value;
                        bool paren = rest.StartsWith("(");
                        if (!paren && ns.Length > 0 && char.IsDigit(a.F[a.F.Length - 1])) continue;   // SO42: không rõ, bỏ
                        int typedCat = cn.Length > 0 ? int.Parse(cn) : 1, typedAn = ns.Length > 0 ? int.Parse(ns) : 1;
                        foreach (int q in qs)
                        {
                            long g = Gcd(q, a.Q);
                            int nc = (int)(a.Q / g), na = (int)(q / g);
                            string catStr = cat == "NH4" ? (nc > 1 ? "(NH4)" + nc : "NH4") : cat + (nc > 1 ? nc.ToString() : "");
                            string f = catStr + AnionStr(a, na);
                            if (f == s) continue;
                            Comp c = ParseF(f);
                            if (c == null || Valid(c, 0) != true) continue;
                            found.Add(new KeyValuePair<string, int>(f, Math.Abs(nc - typedCat) + Math.Abs(na - typedAn)));
                        }
                        break;      // chỉ lấy anion khớp dài nhất
                    }
                }
            }
            // 2. hai nguyên tố bất kỳ (Mg3N2, B2O3...): mọi tổ hợp hợp hóa trị, tối giản
            var els = comp.Keys;
            if (found.Count == 0 && els.Count == 2 && s.IndexOf('(') < 0)
            {
                for (int a = 1; a <= 6; a++)
                    for (int b = 1; b <= 6; b++)
                    {
                        if (Gcd(a, b) != 1) continue;
                        var c = new Comp(); c.Add(els[0], a); c.Add(els[1], b);
                        if (Valid(c, 0) != true) continue;
                        string f = FStr(c);
                        if (f == s) continue;
                        found.Add(new KeyValuePair<string, int>(f, Math.Abs(a - comp.N[els[0]]) + Math.Abs(b - comp.N[els[1]])));
                    }
            }
            var res = new List<string>();
            foreach (var kv in found.OrderBy(x => x.Value).ThenBy(x => x.Key.Length)) { if (!res.Contains(kv.Key)) res.Add(kv.Key); if (res.Count >= 3) break; }
            return res;
        }

        static string SuggestCorr(string s, Comp comp) { List<string> l = Suggest(s, comp); return l.Count > 0 ? l[0] : null; }

        // ---------------------------------------------------------------- chữ thường: liệt kê mọi cách đọc
        static List<List<string>> Segs(string r)
        {
            var o = new List<List<string>>();
            if (r.Length == 0) { o.Add(new List<string>()); return o; }
            for (int L = 1; L <= 2; L++)
            {
                if (r.Length < L) continue;
                string t = r.Substring(0, L), sym = char.ToUpperInvariant(t[0]) + t.Substring(1);
                if (ELS.Contains(sym))
                    foreach (List<string> rest in Segs(r.Substring(L)))
                    {
                        var l = new List<string>(); l.Add(sym); l.AddRange(rest); o.Add(l);
                    }
            }
            return o;
        }
        static readonly Regex PartRe = new Regex(@"[a-z]+|\d+|[()]", RegexOptions.Compiled);
        static List<string> LowerParses(string tok)
        {
            var res = new List<string>(); res.Add("");
            foreach (Match pm in PartRe.Matches(tok.ToLowerInvariant()))
            {
                string p = pm.Value;
                if (p[0] >= 'a' && p[0] <= 'z')
                {
                    var ss = Segs(p).Select(a => string.Join("", a.ToArray())).ToList();
                    if (ss.Count == 0) return new List<string>();
                    var nr = new List<string>();
                    foreach (string r in res) foreach (string s in ss) { if (nr.Count < 80) nr.Add(r + s); }
                    res = nr;
                }
                else res = res.Select(r => r + p).ToList();
            }
            return res;
        }

        // ---------------------------------------------------------------- một chất
        sealed class Attempt { public string Core; public Comp Comp; public bool Cased; public string Orig; }
        sealed class Cand { public string Core; public int Charge; }

        static Attempt TryAttempt(string c0, int chg)
        {
            string orig = c0;
            Comp comp = ParseF(c0); bool cased = false;
            bool hasLower = c0.Any(ch => ch >= 'a' && ch <= 'z'), hasUpper = c0.Any(ch => ch >= 'A' && ch <= 'Z');
            if (comp == null || (c0 == c0.ToLowerInvariant() && !COMMON.Contains(c0)) || (hasLower && !hasUpper))
            {
                var ps = LowerParses(c0).Select(s => new { s, comp = ParseF(s) }).Where(x => x.comp != null).ToList();
                if (ps.Count > 0)
                {
                    var pick = ps.Where(x => Valid(x.comp, chg) == true).ToList();
                    if (pick.Count == 0) pick = ps;
                    if (pick.Count > 1)
                    {
                        var d = pick.Where(x => COMMON.Contains(x.s)).ToList();
                        if (d.Count == 1) pick = d;
                    }
                    if (pick.Count == 1 && (comp == null || pick[0].s != c0)) { comp = pick[0].comp; c0 = pick[0].s; cased = true; }
                    else if (pick.Count > 1 && ParseF(c0) == null) return null;
                }
            }
            return comp != null ? new Attempt { Core = c0, Comp = comp, Cased = cased, Orig = orig } : null;
        }

        static readonly Regex CoefRe = new Regex(@"^(\d+)(?=[A-Za-z(])");
        static readonly Regex ChargeRe = new Regex(@"^(.*?)\^?\{?(\d*)([+-])\}?$");

        public static Species NormSpecies(string raw)
        {
            var sp = new Species { Raw = raw };
            string t = raw;
            Match cm = CoefRe.Match(t);
            if (cm.Success) { sp.Coef = cm.Groups[1].Value; t = t.Substring(sp.Coef.Length); }
            Match ch = ChargeRe.Match(t);
            var cand = new List<Cand>();
            if (ch.Success && ch.Groups[1].Value.Length > 0 && Regex.IsMatch(ch.Groups[1].Value, @"[A-Za-z0-9)]$") && !Regex.IsMatch(ch.Groups[1].Value, @"^\d+$"))
            {
                int sg = ch.Groups[3].Value == "+" ? 1 : -1;
                int mag = ch.Groups[2].Value.Length > 0 ? int.Parse(ch.Groups[2].Value) : 1;
                cand.Add(new Cand { Core = ch.Groups[1].Value, Charge = sg * mag });
                if (ch.Groups[2].Value.Length > 0 && t.IndexOf('^') < 0) cand.Add(new Cand { Core = ch.Groups[1].Value + ch.Groups[2].Value, Charge = sg });
            }
            cand.Add(new Cand { Core = t, Charge = 0 });
            Attempt r = null; int usedCharge = 0;
            foreach (Cand c in cand)
            {
                Attempt a = TryAttempt(c.Core, c.Charge);
                if (a != null && Valid(a.Comp, c.Charge) != false) { r = a; usedCharge = c.Charge; break; }
            }
            if (r == null)
            {
                foreach (Cand c in cand)
                {
                    Attempt a = TryAttempt(c.Core, c.Charge);
                    if (a != null) { r = a; usedCharge = c.Charge; break; }
                }
            }
            if (r == null) { sp.Out = raw; sp.Unknown = true; return sp; }
            string output = r.Core;
            Comp comp = r.Comp;
            if (r.Cased) sp.Flags.Add(new ChemFlag { Type = "case", Msg = r.Orig + " đọc thành " + output + " (ký hiệu nguyên tố phải viết hoa chữ đầu)" });
            bool? v = Valid(comp, usedCharge);
            if (v == true && usedCharge == 0 && !MixedOk.Contains(FStr(comp)) && IonicBad(output)) v = false;
            sp.TypedOut = output;
            if (v == false)
            {
                List<string> alts = usedCharge == 0 ? Suggest(output, comp) : new List<string>();
                string sg = alts.Count > 0 ? alts[0] : null;
                string msg = output + " không hợp hóa trị";
                if (alts.Count == 1) msg += ", thường viết là " + sg;
                else if (alts.Count > 1) msg += ", có thể là " + string.Join(" hoặc ", alts.ToArray());
                var fl = new ChemFlag { Type = "val", Msg = msg, Sug = sg }; fl.Alts.AddRange(alts);
                sp.Flags.Add(fl); sp.Alts.AddRange(alts);
                if (sg != null) { Comp c2 = ParseF(sg); if (c2 != null) { output = sg; comp = c2; } }
            }
            else if (v == true && usedCharge == 0)
            {
                string al;
                if (ALIAS.TryGetValue(output, out al))
                {
                    sp.Flags.Add(new ChemFlag { Type = "order", Msg = output + " không sai về hóa học nhưng cách viết chuẩn là " + al, Sug = al });
                    output = al; comp = ParseF(al);
                }
            }
            if (usedCharge != 0) output += "^" + (Math.Abs(usedCharge) > 1 ? Math.Abs(usedCharge).ToString() : "") + (usedCharge > 0 ? "+" : "-");
            sp.Out = output; sp.Comp = comp; sp.Charge = usedCharge; sp.Unknown = false;
            return sp;
        }

        // ---------------------------------------------------------------- cân bằng
        public static int[] Balance(List<Species> sp, int nL)
        {
            if (sp.Any(s => s.Comp == null)) return null;
            var els = new List<string>();
            foreach (Species s in sp) foreach (string k in s.Comp.Keys) if (!els.Contains(k)) els.Add(k);
            var rows = new List<long[]>();
            foreach (string e in els)
            {
                var row = new long[sp.Count];
                for (int j = 0; j < sp.Count; j++) row[j] = sp[j].Comp.Get(e) * (j < nL ? 1 : -1);
                rows.Add(row);
            }
            var cr = new long[sp.Count];
            for (int j = 0; j < sp.Count; j++) cr[j] = sp[j].Charge * (j < nL ? 1 : -1);
            rows.Add(cr);
            int n = sp.Count, r0 = 0;
            var piv = new List<int>();
            for (int c = 0; c < n && r0 < rows.Count; c++)
            {
                int p = r0;
                while (p < rows.Count && rows[p][c] == 0) p++;
                if (p == rows.Count) continue;
                var tmp = rows[r0]; rows[r0] = rows[p]; rows[p] = tmp;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (i != r0 && rows[i][c] != 0)
                    {
                        long a = rows[r0][c], b = rows[i][c];
                        var nr = new long[n];
                        for (int k = 0; k < n; k++) nr[k] = rows[i][k] * a - rows[r0][k] * b;
                        long g = 0; foreach (long v in nr) g = Gcd(g, v);
                        if (g > 1) for (int k = 0; k < n; k++) nr[k] /= g;
                        rows[i] = nr;
                    }
                }
                piv.Add(c); r0++;
            }
            var free = new List<int>();
            for (int c = 0; c < n; c++) if (!piv.Contains(c)) free.Add(c);
            if (free.Count != 1) return null;
            int f = free[0];
            var fn = new long[piv.Count]; var fd = new long[piv.Count];
            for (int i = 0; i < piv.Count; i++) { fn[i] = -rows[i][f]; fd[i] = rows[i][piv[i]]; }
            long L = 1;
            for (int i = 0; i < piv.Count; i++) L = L * Math.Abs(fd[i]) / Gcd(L, fd[i]);
            var x = new long[n]; x[f] = L;
            for (int i = 0; i < piv.Count; i++) x[piv[i]] = fn[i] * L / fd[i];
            foreach (long v in x) if (v == 0 || (v < 0) != (x[f] < 0)) return null;
            long gg = 0; foreach (long v in x) gg = Gcd(gg, v);
            var res = new int[n];
            for (int i = 0; i < n; i++) res[i] = (int)(Math.Abs(x[i]) / gg);
            return res;
        }

        public static bool IsBalanced(List<Species> sp, int nL)
        {
            var tot = new Dictionary<string, long>(); long q = 0;
            for (int j = 0; j < sp.Count; j++)
            {
                long m = (sp[j].Coef.Length > 0 ? long.Parse(sp[j].Coef) : 1) * (j < nL ? 1 : -1);
                foreach (string e in sp[j].Comp.Keys) { long cur; tot.TryGetValue(e, out cur); tot[e] = cur + sp[j].Comp.N[e] * m; }
                q += sp[j].Charge * m;
            }
            return q == 0 && tot.Values.All(v => v == 0);
        }

        // ---------------------------------------------------------------- phương trình
        const string SPEC = @"\d*[A-Za-z(][A-Za-z0-9()]*(?:\^?\{?\d*[+-]\}?(?![A-Za-z0-9]))?";
        public static readonly Regex RxnRe = new Regex(
            @"(?<![A-Za-z0-9^])((?:" + SPEC + @"\s*\+\s*)*" + SPEC + @")\s*(<=>|<->|=>|->|-->|→|⇌|=)\s*((?:" + SPEC + @"\s*\+\s*)*" + SPEC + @")(?![A-Za-z0-9^])",
            RegexOptions.None, TimeSpan.FromSeconds(2));
        static readonly Regex SplitRe = new Regex(@"\s*\+\s*(?=\d*[A-Za-z(])");

        static List<string> SplitSide(string s)
        {
            return SplitRe.Split(s).Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
        }

        static string BuildRxn(List<Species> all, int nL, string ar, int[] co) { return BuildRxn(all, nL, ar, co, false); }
        /// <summary>typed = true: dùng công thức như đã gõ (chỉ sửa chữ hoa), không thay theo gợi ý hóa trị.</summary>
        public static string BuildRxn(List<Species> all, int nL, string ar, int[] co, bool typed)
        {
            var lp = new List<string>(); var rp = new List<string>();
            for (int i = 0; i < all.Count; i++)
            {
                string c = co != null ? (co[i] > 1 ? co[i].ToString() : "") : all[i].Coef;
                string o = typed && all[i].TypedOut.Length > 0 ? all[i].TypedOut + ChargeSuffix(all[i]) : all[i].Out;
                (i < nL ? lp : rp).Add(c + o);
            }
            return string.Join(" + ", lp.ToArray()) + " " + ar + " " + string.Join(" + ", rp.ToArray());
        }

        static string ChargeSuffix(Species s)
        {
            int q = s.Charge; if (q == 0) return "";
            return "^" + (Math.Abs(q) > 1 ? Math.Abs(q).ToString() : "") + (q > 0 ? "+" : "-");
        }

        public static Rxn AnalyzeRxn(string left, string arrow, string right)
        {
            var L = SplitSide(left).Select(NormSpecies).ToList();
            var R = SplitSide(right).Select(NormSpecies).ToList();
            var all = new List<Species>(L); all.AddRange(R);
            string ar = (arrow.IndexOf('<') >= 0 || arrow.IndexOf('⇌') >= 0) ? "<=>" : "->";
            var res = new Rxn { Species = all, NL = L.Count, Arrow = ar };
            res.HasCoef = all.Any(x => x.Coef.Length > 0);
            foreach (Species s in all) res.Flags.AddRange(s.Flags);
            res.Unknown = all.Any(s => s.Unknown);
            res.Std = BuildRxn(all, L.Count, ar, null);
            if (!res.Unknown)
            {
                bool bal = IsBalanced(all, L.Count);
                res.Balanced = bal;
                if (!bal) { int[] co = Balance(all, L.Count); if (co != null) res.BalStr = BuildRxn(all, L.Count, ar, co); }
                // chất có nhiều gợi ý (FeCl2 hoặc FeCl3): thêm các cách đọc khác
                int made = 0;
                for (int i = 0; i < all.Count && made < 3; i++)
                {
                    Species s0 = all[i];
                    for (int k = 1; k < s0.Alts.Count && made < 3; k++)
                    {
                        Comp c2 = ParseF(s0.Alts[k]); if (c2 == null) continue;
                        var v = new Species { Raw = s0.Raw, Out = s0.Alts[k], Coef = s0.Coef, Comp = c2, Charge = 0 };
                        var all2 = new List<Species>(all); all2[i] = v;
                        var alt = new RxnAlt { Label = "Đọc " + s0.TypedOut + " là " + s0.Alts[k], Std = BuildRxn(all2, L.Count, ar, null) };
                        if (!IsBalanced(all2, L.Count)) { int[] co2 = Balance(all2, L.Count); if (co2 != null) alt.BalStr = BuildRxn(all2, L.Count, ar, co2); }
                        res.Alts.Add(alt); made++;
                    }
                }
            }
            return res;
        }

        // ---------------------------------------------------------------- hiển thị
        /// <summary>Chuỗi chuẩn (có ^ và -&gt;) sang các đoạn chữ có định dạng: chỉ số dưới cho số nguyên tử, chỉ số trên cho điện tích.</summary>
        public static List<Run> RenderRuns(string src)
        {
            var runs = new List<Run>(); var cur = new StringBuilder();
            Action flush = delegate { if (cur.Length > 0) { runs.Add(new Run(cur.ToString(), 0)); cur.Length = 0; } };
            int i = 0;
            while (i < src.Length)
            {
                if (string.CompareOrdinal(src, i, "<=>", 0, 3) == 0) { cur.Append('⇌'); i += 3; continue; }
                if (string.CompareOrdinal(src, i, "->", 0, 2) == 0) { cur.Append('→'); i += 2; continue; }
                char c = src[i];
                if (c == '^')
                {
                    int j = i + 1;
                    if (j < src.Length && src[j] == '{') j++;
                    int ds = j; while (j < src.Length && src[j] >= '0' && src[j] <= '9') j++;
                    string digits = src.Substring(ds, j - ds);
                    if (j < src.Length && (src[j] == '+' || src[j] == '-'))
                    {
                        string sign = src[j] == '-' ? "−" : "+"; j++;
                        if (j < src.Length && src[j] == '}') j++;
                        flush(); runs.Add(new Run(digits + sign, 2)); i = j; continue;
                    }
                }
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == ')' || c == ']')
                {
                    cur.Append(c); i++;
                    int j = i; while (j < src.Length && src[j] >= '0' && src[j] <= '9') j++;
                    if (j > i) { flush(); runs.Add(new Run(src.Substring(i, j - i), 1)); i = j; }
                    continue;
                }
                cur.Append(c); i++;
            }
            flush();
            return runs;
        }

        public static string RunsPlain(List<Run> runs)
        {
            var sb = new StringBuilder(); foreach (Run r in runs) sb.Append(r.T); return sb.ToString();
        }

        /// <summary>Các đoạn chữ sang chuỗi UnicodeMath để Word dựng thành equation. Chữ cái đặt trong ngoặc kép để in thẳng (không nghiêng).</summary>
        public static string RunsToLinear(List<Run> runs)
        {
            var sb = new StringBuilder();
            foreach (Run r in runs)
            {
                if (r.Kind == 1) sb.Append('_').Append(r.T);
                else if (r.Kind == 2) sb.Append("^(").Append(r.T).Append(')');
                else
                {
                    int i = 0;
                    while (i < r.T.Length)
                    {
                        char c = r.T[i];
                        if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                        {
                            int j = i; while (j < r.T.Length && ((r.T[j] >= 'A' && r.T[j] <= 'Z') || (r.T[j] >= 'a' && r.T[j] <= 'z'))) j++;
                            sb.Append('"').Append(r.T.Substring(i, j - i)).Append('"'); i = j;
                        }
                        else { sb.Append(c); i++; }
                    }
                }
            }
            return sb.ToString();
        }
    }

    /// <summary>Một đoạn chữ có kiểu: 0 thường, 1 chỉ số dưới, 2 chỉ số trên, 3 nghiêng.</summary>
    public sealed class Run
    {
        public string T; public int Kind;
        public Run(string t, int k) { T = t; Kind = k; }
    }
}
