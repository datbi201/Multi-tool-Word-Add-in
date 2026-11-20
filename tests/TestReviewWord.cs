// Kiểm thử WordReview với tài liệu Word giả có văn bản, Range tự dịch chuyển, chỉ số trên/dưới và công thức.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using TroLyKyThuat;

public class RDoc : Rec
{
    public StringBuilder T = new StringBuilder(); public List<bool> Sub = new List<bool>(), Sup = new List<bool>(), Ita = new List<bool>(), Eq = new List<bool>(); public List<int> Hl = new List<int>(); public List<string> Fnt = new List<string>();
    public readonly HashSet<int> CellEnds = new HashSet<int>();
    public List<RRange> Live = new List<RRange>();
    public class EqRec { public RRange R; public string Xml; }
    public List<EqRec> EqList = new List<EqRec>();
    public RDoc(string text) : base("rdoc") { Insert(0, text); }
    public void Insert(int at, string s) { T.Insert(at, s); for (int i = 0; i < s.Length; i++) { Sub.Insert(at, false); Sup.Insert(at, false); Ita.Insert(at, false); Eq.Insert(at, false); Hl.Insert(at, 0); Fnt.Insert(at, ""); } }
    public RRange Content { get { return Mk(0, T.Length); } }
    public RRange Mk(int s, int e) { var r = new RRange(this, s, e); Live.Add(r); return r; }
    public RRange Range(int s, int e) { return Mk(s, e); }
    public string FullName { get { return "test.docx"; } }
    public int UndoCalls;
    public void Undo(int count) { UndoCalls += count; }
    public void Replace(int s, int e, string n)
    {
        int d = n.Length - (e - s);
        T.Remove(s, e - s); Sub.RemoveRange(s, e - s); Sup.RemoveRange(s, e - s); Ita.RemoveRange(s, e - s); Eq.RemoveRange(s, e - s); Hl.RemoveRange(s, e - s); Fnt.RemoveRange(s, e - s);
        Insert(s, n);
        foreach (var r in Live)
        {
            if (r.S >= e) { r.S += d; r.E += d; }
            else if (r.S >= s && r.E <= e) { /* nằm trong vùng bị thay: xử lý riêng */ }
            else if (r.S <= s && r.E >= e) r.E += d;
            else if (r.E > e) r.E += d;
        }
    }
}
public class RRange : Rec
{
    public RDoc D; public int S, E;
    public RRange(RDoc d, int s, int e) : base("rrange") { D = d; S = s; E = e; }
    public string Text
    {
        get
        {
            string text = D.T.ToString(S, E - S);
            if (D.CellEnds.Count == 0) return text;
            var wordText = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                wordText.Append(text[i]);
                if (D.CellEnds.Contains(S + i)) wordText.Append('\a');
            }
            return wordText.ToString();
        }
        set { int s = S, e = E; D.Replace(s, e, value); S = s; E = s + value.Length; }
    }
    public int Start { get { return S; } }
    public int End { get { return E; } }
    public IEnumerable Paragraphs
    {
        get
        {
            var l = new List<RPara>();
            string all = D.T.ToString();
            if (S == E)
            {
                int ps = S == 0 ? 0 : all.LastIndexOf('\r', Math.Max(0, S - 1)) + 1; if (S > 0 && S <= all.Length && all[S - 1] == '\r') ps = S;
                int pe = all.IndexOf('\r', S); pe = pe < 0 ? all.Length : pe + 1;
                l.Add(new RPara(D, ps, pe)); return l;
            }
            int a = S;
            while (a < E)
            {
                int ps = a == 0 ? 0 : all.LastIndexOf('\r', Math.Max(0, a - 1)) + 1; if (a > 0 && a <= all.Length && all[a - 1] == '\r') ps = a;
                int pe = all.IndexOf('\r', a); pe = pe < 0 ? all.Length : pe + 1;
                l.Add(new RPara(D, ps, pe));
                a = Math.Max(a + 1, pe);
            }
            return l;
        }
    }
    public RRange Range { get { return this; } }
    public static string LastXml = "";
    public void InsertXML(string xml)
    {
        LastXml = xml;
        var sb = new StringBuilder();
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(xml, "<m:t[^>]*>([^<]*)</m:t>")) sb.Append(m.Groups[1].Value);
        string v = sb.ToString(); int s0 = S;
        D.Replace(S, E, v); S = s0; E = s0 + v.Length;
        for (int i = S; i < E; i++) D.Eq[i] = true;
        D.EqList.Add(new RDoc.EqRec { R = D.Mk(S, E), Xml = xml });
    }
    public string WordOpenXML { get { var e = D.EqList.FirstOrDefault(x => x.R.S == S && x.R.E == E); return e != null ? e.Xml : "<empty/>"; } }
    public void Delete()
    {
        int s = S, e = E; D.Replace(s, e, ""); S = E = s;
        D.EqList.RemoveAll(x => x.R.S >= s && x.R.E <= e + 0 && x.R.E > x.R.S && x.R.S == s || x.R.S == s && x.R.E <= e);
    }
    public void InsertAfter(string v) { int e = E; D.Replace(e, e, v); E = e + v.Length; }
    public void Collapse(int dir) { if (dir == 0) S = E; else E = S; }
    public void SetRange(int s, int e) { S = s; E = e; }
    public bool Information(int k) { return false; }
    public object HighlightColorIndex { get { int v = D.Hl[S]; for (int i = S; i < E; i++) if (D.Hl[i] != v) return 9999999; return v; } set { for (int i = S; i < E; i++) D.Hl[i] = Convert.ToInt32(value); } }
    public RFont Font { get { return new RFont(this); } }
    public ROMaths OMaths { get { return new ROMaths(this); } }
    public RFields Fields { get { return new RFields(); } }
    public void Select() { Rec.Log.Add("select " + S + "-" + E); }
}
public class RFields : Rec { public RFields() : base("f") { } public int Count { get { return 0; } } }
public class RPara : Rec
{
    RDoc d; int s, e; public RPara(RDoc d, int s, int e) : base("para") { this.d = d; this.s = s; this.e = e; }
    public RRange Range { get { return d.Mk(s, e); } }
    public int OutlineLevel { get { return 10; } }
}
public class RFont : Rec
{
    RRange r; public RFont(RRange r) : base("font") { this.r = r; }
    int Get(List<bool> l) { bool any = false, all = true; for (int i = r.S; i < r.E; i++) { if (l[i]) any = true; else all = false; } return all && any ? -1 : (any ? 9999999 : 0); }
    void Set(List<bool> l, object v) { bool on = Convert.ToInt32(v) != 0; for (int i = r.S; i < r.E; i++) l[i] = on; }
    public object Subscript { get { return Get(r.D.Sub); } set { Set(r.D.Sub, value); } }
    public object Superscript { get { return Get(r.D.Sup); } set { Set(r.D.Sup, value); } }
    public object Italic { get { return Get(r.D.Ita); } set { Set(r.D.Ita, value); } }
    public string Name { get { return r.D.Fnt[r.S]; } set { for (int i = r.S; i < r.E; i++) r.D.Fnt[i] = value; } }
    public object Bold = 0;
    public object Size = 12f;
    public object Color = 0;
    public object Underline = 0;
}
public class ROMaths : Rec
{
    RRange r; public ROMaths(RRange r) : base("omaths") { this.r = r; }
    public int Count { get { int n = r.D.EqList.Count(x => x.R.E > x.R.S && x.R.S >= r.S && x.R.E <= r.E); if (n > 0) return n; for (int i = r.S; i < r.E; i++) if (r.D.Eq[i]) return 1; return 0; } }
    public ROMath Add(RRange rg) { for (int i = rg.S; i < rg.E; i++) rg.D.Eq[i] = true; return new ROMath(); }
    public ROMath Item(int i)
    {
        var l = r.D.EqList.Where(x => x.R.E > x.R.S && x.R.S >= r.S && x.R.E <= r.E).OrderBy(x => x.R.S).ToList();
        return new ROMath(i >= 1 && i <= l.Count ? l[i - 1].R : null);
    }
}
public class ROMath : Rec { public int Type { get; set; } RRange rg; public ROMath() : base("omath") { } public ROMath(RRange r) : base("omath") { rg = r; } public RRange Range { get { return rg; } } public void BuildUp() { Rec.Log.Add("buildup"); } public void Remove() { Rec.Log.Add("omath.remove"); } }
public class RApp : Rec
{
    public RDoc D; public RApp(RDoc d) : base("rapp") { D = d; }
    public RDocs Documents { get { return new RDocs(); } }
    public RDoc ActiveDocument { get { return D; } }
    public RRange Sel;
    public RRange Selection { get { if (Sel == null) Sel = D.Mk(0, 0); return Sel; } }
    public Rec UndoRecord { get { return new Rec("undo"); } }
}
public class RDocs : Rec { public RDocs() : base("docs") { } public int Count { get { return 1; } } }

static class TRW
{
    static int fails = 0, total = 0;
    static void Eq(object a, object b, string name) { total++; bool ok = Equals(a, b); if (!ok) fails++; Console.WriteLine((ok ? "PASS " : "FAIL ") + name + (ok ? "" : "  (nhan [" + a + "], mong [" + b + "])")); }
    static int Main()
    {
        string text = "Sour corrosion follows Fe + HCl -> fecl2 + H2.\rThe H2S level is 12500 mm in 2023. See Fe2+.\r";
        var st = new AppSettings(); st.NumStyle = "us"; st.RvNumStyle = "us"; st.NumDec = "keep"; st.NumThou = true; st.EqMode = "fmt";
        var d = new RDoc(text); var app = new RApp(d); var rv = new WordReview(app);
        RvResult r = rv.Scan(st);
        Console.WriteLine("  -> " + r.Message);
        foreach (RvRow x in r.Rows) Console.WriteLine("     " + x.SevText + " | " + x.CatText + " | " + x.Before + " -> " + x.Show);
        Eq(r.Rows.Count, 4, "tìm thấy 4 mục: phương trình, H2S, 12500, Fe2+");
        RvRow eq = r.Rows.First(x => x.Before.StartsWith("Fe + HCl")); Eq(eq.Multi, true, "phương trình chưa cân bằng có hai cách sửa");
        // sửa ở chế độ định dạng ký tự: H2S -> H với 2 là chỉ số dưới
        RvRow h2s = r.Rows.First(x => x.Before == "H2S");
        rv.FixItem(h2s.Index, 0, st);
        int at = d.T.ToString().IndexOf("H2S");
        Eq(d.Sub[at + 1], true, "chữ số 2 của H2S là chỉ số dưới"); Eq(d.Sub[at], false, "chữ H không bị chỉ số dưới");
        // sửa số kỹ thuật (chữ thuần)
        RvRow num = r.Rows.First(x => x.Before == "12500"); rv.FixItem(num.Index, 0, st);
        Eq(d.T.ToString().Contains("12,500 mm"), true, "12500 -> 12,500");
        // vị trí các mục sau không lệch: sửa tiếp Fe2+ (nằm sau chỗ đã đổi độ dài)
        RvRow ion = r.Rows.First(x => x.Before == "Fe2+"); rv.FixItem(ion.Index, 0, st);
        string now = d.T.ToString(); int fe = now.IndexOf("Fe2+ ");
        if (fe < 0) fe = now.IndexOf("Fe2+.");
        Eq(fe >= 0, true, "văn bản Fe2+ vẫn còn đó sau các sửa đổi trước");
        Eq(d.Sup[fe + 2] && d.Sup[fe + 3], true, "2+ là chỉ số trên");
        // sửa tất cả (còn lại: phương trình mức Sai) và kiểm tra Pending
        RvResult p = rv.Pending(); Eq(p.Rows.Count, 1, "còn lại một mục (phương trình)");
        string fm = rv.FixAll(st); Console.WriteLine("  -> " + fm);
        Eq(d.T.ToString().Contains("FeCl2"), true, "fecl2 đã thành FeCl2 sau Sửa tất cả"); Eq(rv.Pending().Rows.Count, 0, "hết mục");
        // quét lại: các chỗ đã định dạng không bị báo lại
        r = rv.Scan(st); Console.WriteLine("  -> " + r.Message);
        Eq(r.Rows.Count, 0, "quét lại không còn mục nào (đã định dạng sẵn thì bỏ qua)");
        // công thức nằm riêng một dòng: luôn dùng Equation (OMath), công thức cùng dòng với chữ: theo EqMode
        st.EqMode = "fmt";
        d = new RDoc("Gas H2S is present.\rFe + H2S -> FeS + H2\rAnd x^2 + y^2 = r^2 inline.\r"); app = new RApp(d); rv = new WordReview(app);
        r = rv.Scan(st); Eq(r.Rows.Count, 3, "ba mục: H2S cùng dòng, phương trình riêng dòng, toán cùng dòng"); rv.FixAll(st);
        string all = d.T.ToString();
        int eqAt = all.IndexOf("Fe+H");
        Eq(eqAt >= 0 && d.Eq[eqAt], true, "phương trình nằm riêng một dòng được dựng bằng Equation (OMath)");
        int hAt = all.IndexOf("H2S");
        Eq(d.Eq[hAt], false, "H2S cùng dòng với chữ không dùng Equation"); Eq(d.Sub[hAt + 1], true, "H2S: chữ số 2 là chỉ số dưới (định dạng ký tự)");
        r = rv.Scan(st); Eq(r.Rows.Count, 0, "quét lại: công thức Word và chỉ số đã có không bị báo lại");
        // chế độ font chữ nhỏ và thấp (glyph)
        st.EqMode = "glyph";
        d = new RDoc("The H2O and SO4^2- ions, and x^2 + y^2 = r^2 here.\r"); app = new RApp(d); rv = new WordReview(app);
        r = rv.Scan(st); rv.FixAll(st); all = d.T.ToString();
        Eq(all.Contains("H₂O") && all.Contains("SO₄²⁻"), true, "dùng ký tự chỉ số nhỏ thấp: H₂O, SO₄²⁻");
        int gi = all.IndexOf('₂'); Eq(d.Fnt[gi], "Cambria Math", "ký tự chỉ số dùng font Cambria Math"); Eq(d.Fnt[all.IndexOf('H')], "", "chữ H giữ font gốc");
        Eq(d.Sub[gi], false, "không dùng thuộc tính chỉ số dưới khi đã có glyph"); Eq(all.Contains("x² + y² = r²") && d.Ita[all.IndexOf((char)120)] && d.Ita[all.IndexOf((char)114)], true, "biến x, r in nghiêng và số mũ dạng glyph"); Eq(d.Ita[all.IndexOf('H')], false, "chữ hóa học đứng thẳng");
        // công thức hóa học trong đoạn chữ nghiêng phải thẳng
        st.EqMode = "fmt";
        d = new RDoc("Water H_2O is vital.\r"); app = new RApp(d); rv = new WordReview(app);
        for (int i = 0; i < d.Ita.Count; i++) d.Ita[i] = true;
        r = rv.Scan(st); Eq(r.Rows.Count, 1, "H_2O nhận diện"); Eq(r.Rows[0].Cat, "chem", "H_2O là hóa học"); rv.FixAll(st);
        all = d.T.ToString(); int hh = all.IndexOf("H2O");
        Eq(hh >= 0 && d.Sub[hh + 1] && !d.Sub[hh] && !d.Sub[hh + 2], true, "H_2O thành H2O, chỉ số 2 là chỉ số dưới");
        Eq(!d.Ita[hh] && !d.Ita[hh + 1] && !d.Ita[hh + 2], true, "H, 2, O đều thẳng dù đoạn đang nghiêng");
        // tô sáng mục đang xem và trả lại
        d = new RDoc("Gas H2S is present and 12500 mm.\r"); app = new RApp(d); rv = new WordReview(app);
        r = rv.Scan(st); int hs0 = d.T.ToString().IndexOf("H2S");
        rv.Goto(r.Rows[0].Index); Eq(d.Hl[hs0], 7, "mục đang xem được tô vàng");
        int n12 = d.T.ToString().IndexOf("12500"); rv.Goto(r.Rows[1].Index);
        Eq(d.Hl[hs0], 0, "mục trước được bỏ tô khi chuyển mục"); Eq(d.Hl[n12], 7, "số đang xem được tô vàng");
        rv.FixItem(r.Rows[1].Index, 0, st); Eq(d.Hl.Any(x => x != 0), false, "sửa xong thì không còn tô sáng");
        rv.Goto(r.Rows[0].Index); rv.ClearMark(); Eq(d.Hl.Any(x => x != 0), false, "ClearMark trả lại như cũ");
        // màu tô sáng có sẵn được giữ lại
        d = new RDoc("Gas H2S is present.\r"); app = new RApp(d); rv = new WordReview(app);
        int hp = d.T.ToString().IndexOf("H2S"); for (int i = hp; i < hp + 3; i++) d.Hl[i] = 4;
        r = rv.Scan(st); rv.Goto(r.Rows[0].Index); Eq(d.Hl[hp], 7, "đang tô sáng"); rv.ClearMark(); Eq(d.Hl[hp], 4, "màu tô sáng ban đầu (4) được trả lại");
        // số 'năm' không bị đề xuất sửa
        d = new RDoc("Built in 1889 and 1839, size 1,234 mm.\r"); app = new RApp(d); rv = new WordReview(app);
        r = rv.Scan(st); Eq(r.Rows.Count, 0, "năm 1839, 1889 không bị nghi ngờ");
        // bỏ qua
        d = new RDoc("Gas H2S is present.\r"); app = new RApp(d); rv = new WordReview(app); st.EqMode = "fmt";
        r = rv.Scan(st); rv.Ignore(r.Rows[0].Index); r = rv.Scan(st); Eq(r.Rows.Count, 0, "mục đã bỏ qua không hiện lại"); rv.ClearIgnored(); r = rv.Scan(st); Eq(r.Rows.Count, 1, "Hiện lại sau khi xóa danh sách bỏ qua");
        // văn bản bị đổi sau khi quét
        rv.Scan(st); d.Replace(4, 7, "XYZ"); try { rv.FixItem(0, 0, st); Eq(false, true, "phải báo văn bản đã đổi"); } catch (UserMsg) { Eq(true, true, "văn bản đã đổi sau khi quét thì báo lỗi, không sửa nhầm"); }

        // === Rev0.3 & Rev0.9.5: chèn công thức đã soạn (không còn ô chọn alone, tự nhận theo ngữ cảnh)
        {
            var wc = new WordCompose(null);
            d = new RDoc("Before  after.\r"); app = new RApp(d); wc = new WordCompose(app); app.Sel = d.Mk(7, 7);
            ComposeResult cr = Compose.Build("H_2O", "chem", false);
            string im = wc.Insert(cr, false, "glyph", false); Console.WriteLine("  -> " + im);
            Eq(d.T.ToString().StartsWith("Before H₂O ") || d.T.ToString().Contains("H₂O"), true, "chèn H₂O cùng dòng bằng glyph");
            int gj = d.T.ToString().IndexOf('₂'); Eq(d.Fnt[gj], "Cambria Math", "glyph dùng Cambria Math");
            d = new RDoc("Before  after.\r"); app = new RApp(d); wc = new WordCompose(app); app.Sel = d.Mk(7, 7);
            wc.Insert(Compose.Build("SO_4^{2-}", "chem", false), false, "fmt", false);
            string t5 = d.T.ToString(); int so = t5.IndexOf("SO42−"); Eq(so >= 0 && d.Sub[so + 2] && d.Sup[so + 3] && d.Sup[so + 4], true, "SO4 2- : chỉ số dưới 4, chỉ số trên 2−");
            d = new RDoc("\r"); app = new RApp(d); wc = new WordCompose(app); app.Sel = d.Mk(0, 0);
            im = wc.Insert(Compose.Build("(a+b)/c", "math", false), false, "fmt", true); Console.WriteLine("  -> " + im);
            Eq(d.T.ToString().Contains("a+bc") && d.Eq[d.T.ToString().IndexOf("a+bc")], true, "phân số riêng dòng dùng Equation (OMML)");
            Eq(RRange.LastXml.Contains("<m:f>") && RRange.LastXml.Contains("m:val=\"bar\"") && RRange.LastXml.Contains("<m:oMathPara>"), true, "kiểu xếp chồng: m:f bar trong oMathPara");
            d = new RDoc("\r"); app = new RApp(d); wc = new WordCompose(app); app.Sel = d.Mk(0, 0);
            wc.Insert(Compose.Build("(a+b)/c", "math", false), false, "fmt", false);
            Eq(RRange.LastXml.Contains("<m:f>"), false, "kiểu tuyến tính: không dùng m:f");
            d = new RDoc("Use here now.\r"); app = new RApp(d); wc = new WordCompose(app); app.Sel = d.Mk(4, 8);
            wc.Insert(Compose.Build("x^2 + y^2 = r^2", "math", false), true, "fmt", false);
            t5 = d.T.ToString(); Eq(t5.StartsWith("Use x2 + y2 = r2") , true, "thay phần bôi đen bằng biểu thức cùng dòng"); Eq(d.Sup[t5.IndexOf("x2") + 1], true, "x2 có số mũ là chỉ số trên");
            try { wc.Insert(new ComposeResult(), false, "fmt", false); Eq(false, true, "phải báo lỗi khi chưa có công thức"); } catch (UserMsg) { Eq(true, true, "chưa có công thức hợp lệ thì báo lỗi"); }
        }
        // === Rev0.9.5: Kiểm thử tự động chèn cùng dòng vs riêng một dòng (bỏ ô chAlone)
        {
            var wcNew = new WordCompose(null);

            // 1. Chèn công thức vào đoạn đang có chữ -> CÙNG DÒNG (inline)
            // 1a. Biểu thức có phân số (complex) chèn vào giữa đoạn có chữ -> Equation cùng dòng (không bọc m:oMathPara)
            d = new RDoc("Gia tri cua f la  ket qua.\r"); app = new RApp(d); wcNew = new WordCompose(app);
            app.Sel = d.Mk(17, 17); // con trỏ giữa đoạn có chữ (sau "la ")
            string mInline1 = wcNew.Insert(Compose.Build("(a+b)/c", "math", false), false, "fmt", true);
            Eq(mInline1.Contains("Equation cùng dòng"), true, "chèn vào đoạn có chữ: Equation cùng dòng");
            Eq(RRange.LastXml.Contains("<m:oMathPara>"), false, "chèn vào đoạn có chữ: không có m:oMathPara");
            Eq(d.T.ToString().Contains("Gia tri cua f la a+bc ket qua.\r"), true, "chèn vào đoạn có chữ: giữ nguyên văn bản trước và sau");

            // 1b. Công thức đơn giản (H2O) chèn vào đoạn có chữ -> định dạng cùng dòng
            d = new RDoc("Chat phan ung:  trong dung dich.\r"); app = new RApp(d); wcNew = new WordCompose(app);
            app.Sel = d.Mk(15, 15);
            string mInline2 = wcNew.Insert(Compose.Build("H_2O", "chem", false), false, "fmt", false);
            Eq(mInline2.Contains("chỉ số"), true, "chèn vào đoạn có chữ: H2O dùng định dạng chỉ số cùng dòng");
            int hPos = d.T.ToString().IndexOf("H2O");
            Eq(hPos >= 0 && d.Sub[hPos + 1] && !d.Eq[hPos], true, "chèn vào đoạn có chữ: H2O không dùng Equation");

            // 2. Chèn vào đoạn trống -> RIÊNG MỘT DÒNG (Equation block)
            // 2a. Đoạn trống nằm giữa hai đoạn khác
            d = new RDoc("Đoạn trước.\r\rĐoạn sau.\r"); app = new RApp(d); wcNew = new WordCompose(app);
            int emptyPos = "Đoạn trước.\r".Length; // vị trí con trỏ tại đoạn thứ 2 hoàn toàn trống
            app.Sel = d.Mk(emptyPos, emptyPos);
            string mAlone1 = wcNew.Insert(Compose.Build("x^2 + y^2 = r^2", "math", false), false, "fmt", false);
            Eq(mAlone1.Contains("Equation riêng một dòng"), true, "chèn vào đoạn trống ở giữa: Equation riêng một dòng");
            Eq(RRange.LastXml.Contains("<m:oMathPara>"), true, "chèn vào đoạn trống ở giữa: có thẻ m:oMathPara");

            // 2b. Đoạn văn hoàn toàn trống (tài liệu mới chỉ có dấu hết đoạn \r)
            d = new RDoc("\r"); app = new RApp(d); wcNew = new WordCompose(app);
            app.Sel = d.Mk(0, 0);
            string mAlone2 = wcNew.Insert(Compose.Build("Fe + HCl -> FeCl2 + H2", "chem", false), false, "fmt", false);
            Eq(mAlone2.Contains("Equation riêng một dòng"), true, "chèn vào đoạn hoàn toàn trống: Equation riêng một dòng");
            Eq(RRange.LastXml.Contains("<m:oMathPara>"), true, "chèn vào đoạn hoàn toàn trống: có thẻ m:oMathPara");

            // 3. Chèn thay đúng một đoạn bôi đen cả đoạn -> RIÊNG MỘT DÒNG
            // 3a. Bôi đen cả đoạn bao gồm cả ký tự hết đoạn \r
            d = new RDoc("Đoạn mở đầu.\rPHUONG TRINH CAN THAY\rĐoạn kết thúc.\r"); app = new RApp(d); wcNew = new WordCompose(app);
            int p2Start = "Đoạn mở đầu.\r".Length;
            int p2End = p2Start + "PHUONG TRINH CAN THAY\r".Length;
            app.Sel = d.Mk(p2Start, p2End); // bôi đen đúng toàn bộ đoạn 2 (gồm \r)
            string mRep1 = wcNew.Insert(Compose.Build("t_{req} = (P D)/(2 S E)", "math", false), true, "fmt", true);
            Eq(mRep1.Contains("Equation riêng một dòng"), true, "bôi đen cả đoạn gồm \\r: chèn Equation riêng một dòng");
            Eq(RRange.LastXml.Contains("<m:oMathPara>"), true, "bôi đen cả đoạn gồm \\r: có thẻ m:oMathPara");
            Eq(d.T.ToString().Contains("PHUONG TRINH CAN THAY"), false, "bôi đen cả đoạn: đã thay thế hoàn toàn chữ cũ");

            // 3b. Bôi đen toàn bộ chữ của đoạn (không bao gồm ký tự \r cuối đoạn)
            d = new RDoc("Đoạn một.\rDOAN CAN THAY KHONG CHUA R\rĐoạn ba.\r"); app = new RApp(d); wcNew = new WordCompose(app);
            int pBStart = "Đoạn một.\r".Length;
            int pBEnd = pBStart + "DOAN CAN THAY KHONG CHUA R".Length;
            app.Sel = d.Mk(pBStart, pBEnd);
            string mRep2 = wcNew.Insert(Compose.Build("(a+b)/c", "math", false), true, "fmt", true);
            Eq(mRep2.Contains("Equation riêng một dòng"), true, "bôi đen cả chữ trong đoạn không gồm \\r: Equation riêng một dòng");
            Eq(RRange.LastXml.Contains("<m:oMathPara>"), true, "bôi đen cả chữ trong đoạn không gồm \\r: có thẻ m:oMathPara");

            // 3c. Bôi đen chỉ một phần đoạn văn -> CÙNG DÒNG (đối chứng)
            d = new RDoc("Đoạn văn có phan_can_thay o giua.\r"); app = new RApp(d); wcNew = new WordCompose(app);
            int partS = d.T.ToString().IndexOf("phan_can_thay");
            int partE = partS + "phan_can_thay".Length;
            app.Sel = d.Mk(partS, partE); // chỉ bôi đen phần giữa đoạn
            string mRepPart = wcNew.Insert(Compose.Build("x^2", "math", false), true, "fmt", false);
            Eq(!mRepPart.Contains("Equation riêng một dòng"), true, "bôi đen một phần đoạn: chèn CÙNG DÒNG");
            Eq(d.T.ToString().StartsWith("Đoạn văn có x2 o giua.\r"), true, "bôi đen một phần đoạn: thay đúng chữ trong dòng");
        }
        // === Rev0.6: Kiểu chỉ số luân phiên raw -> fmt -> eq -> fmt -> eq ...
        {
            d = new RDoc("Gas H2O and SO4^2- here.\r"); app = new RApp(d); app.Sel = d.Mk(4, 4);
            var wi = new WordIndexStyle(app);
            try { wi.Cycle(null); Eq(false, true, "con trỏ trống phải báo lỗi"); } catch (UserMsg) { Eq(true, true, "chưa bôi đen thì báo lỗi"); }
            app.Sel = d.Mk(0, d.T.Length - 1);
            string m1 = wi.Cycle(null); Console.WriteLine("  -> " + m1);
            string a1 = d.T.ToString(); int h = a1.IndexOf("H2O");
            Eq(h >= 0 && d.Sub[h + 1] && !d.Sub[h] && !d.Sub[h + 2], true, "bước 1: H2O có 2 là chỉ số dưới (định dạng)");
            int so = a1.IndexOf("SO4"); Eq(so >= 0 && d.Sub[so + 2] && d.Sup[so + 3], true, "bước 1: SO4 2-: chỉ số dưới 4, chỉ số trên 2-");
            string m2 = wi.Cycle(null); Console.WriteLine("  -> " + m2);
            Eq(d.Eq.Any(x => x), true, "bước 2: font thường chuyển sang Equation");
            Eq(RRange.LastXml.Contains("<m:sSub>") || RRange.LastXml.Contains("<m:sSup>"), true, "bước 2: Equation có chỉ số (OMML)");
            string m3 = wi.Cycle(null); Console.WriteLine("  -> " + m3 + " | " + d.T);
            string a3 = d.T.ToString(); int h3 = a3.IndexOf("H2O");
            Eq(d.Eq.Any(x => x), false, "bước 3: không còn Equation");
            Eq(h3 >= 0 && d.Sub[h3 + 1] && !d.Sub[h3] && !d.Sub[h3 + 2], true, "bước 3: Equation đổi lại thành font thường, H2O có 2 là chỉ số dưới");
            int so3 = a3.IndexOf("SO4"); Eq(so3 >= 0 && d.Sub[so3 + 2] && d.Sup[so3 + 3], true, "bước 3: SO4 2- khôi phục chỉ số dưới 4, chỉ số trên 2-");
            Eq(a3.Contains("Gas ") && a3.Contains(" and ") && a3.Contains(" here."), true, "bước 3: chữ xung quanh không đổi");
            string m4 = wi.Cycle(null); Console.WriteLine("  -> " + m4);
            Eq(d.Eq.Any(x => x), true, "bước 4: quay vòng, lại thành Equation");
            string m5 = wi.Cycle(null); Console.WriteLine("  -> " + m5);
            Eq(d.Eq.Any(x => x), false, "bước 5: lại về font thường");
            // Equation có phân số
            d = new RDoc("Ratio a/b ok.\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1); wi = new WordIndexStyle(app);
            // đặt thẳng một kiểu và đoạn không có công thức
            d = new RDoc("Plain words only here.\r"); app = new RApp(d); app.Sel = d.Mk(0, 10); wi = new WordIndexStyle(app);
            string mn = wi.Cycle(null); Console.WriteLine("  -> " + mn); Eq(mn.StartsWith("Không tìm thấy"), true, "không có công thức thì báo rõ");
            // LaTeX lộ mã
            d = new RDoc("Energy \\Delta G and P_CO2 here.\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1); wi = new WordIndexStyle(app);
            string ml = wi.Cycle("fmt"); Console.WriteLine("  -> " + ml + " | " + d.T);
            string al = d.T.ToString(); Eq(al.Contains("Δ"), true, "\\Delta thành Δ"); int pc = al.IndexOf("P"); 
            int co = al.IndexOf("CO2"); Eq(co >= 0 && d.Sub[co] && d.Sub[co + 1] && d.Sub[co + 2], true, "P_CO2: cả cụm CO2 là chỉ số dưới");
            Eq(d.Ita[al.IndexOf("P")], true, "P viết nghiêng");
        }
        // === Rev0.4: Định dạng số trong văn bản (nhiều số cùng lúc)
        {
            var st4 = new AppSettings(); st4.NumStyle = "us"; st4.NumThou = true;
            d = new RDoc("Flow 1,234.56 and 7,890.5 m3/h, ratio 0.75.\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1);
            var wn = new WordNumbers(app);
            string n1 = wn.Apply(st4, "cycle"); Console.WriteLine("  -> " + n1 + " | " + d.T);
            Eq(st4.NumStyle, "si", "vòng kiểu: Mỹ sang SI"); Eq(d.T.ToString().Contains("1\u00A0234,56") || d.T.ToString().Contains("1\u202F234,56") || d.T.ToString().Contains("1 234,56"), true, "1,234.56 thành kiểu SI");
            Eq(d.T.ToString().Contains("0,75"), true, "0.75 thành 0,75 cùng lúc");
            n1 = wn.Apply(st4, "cycle"); Console.WriteLine("  -> " + n1 + " | " + d.T);
            Eq(st4.NumStyle, "vn", "vòng kiểu: SI sang Việt"); Eq(d.T.ToString().Contains("1.234,56") && d.T.ToString().Contains("7.890,5"), true, "đủ hai số kiểu Việt");
            n1 = wn.Apply(st4, "inc"); Console.WriteLine("  -> " + n1 + " | " + d.T);
            Eq(d.T.ToString().Contains("1.234,560") && d.T.ToString().Contains("7.890,50") && d.T.ToString().Contains("0,750"), true, "tăng một chữ số thập phân cho từng số");
            n1 = wn.Apply(st4, "dec"); n1 = wn.Apply(st4, "dec"); Console.WriteLine("  -> " + n1 + " | " + d.T);
            Eq(d.T.ToString().Contains("1.234,6 and") && d.T.ToString().Contains("7.891 m3/h") && d.T.ToString().Contains("0,8."), true, "giảm hai lần: làm tròn 1.234,56 thành 1.234,6; 7.890,5 thành 7.891");
            // con trỏ trong một số: chỉ đổi số đó
            d = new RDoc("A 1,5 and 2,5 here.\r"); app = new RApp(d); st4.NumStyle = "vn"; app.Sel = d.Mk(4, 4); wn = new WordNumbers(app);
            wn.Apply(st4, "inc"); Eq(d.T.ToString().StartsWith("A 1,50 and 2,5"), true, "con trỏ trong 1,5: chỉ số đó tăng thập phân");
            // tối đa 5 chữ số
            d = new RDoc("x 0,12345 y\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1); wn = new WordNumbers(app);
            string mx = wn.Apply(st4, "inc"); Eq(d.T.ToString().Contains("0,12345"), true, "đã tối đa 5 chữ số thì không tăng thêm"); Console.WriteLine("  -> " + mx);
            // không có số
            d = new RDoc("no digits here\r"); app = new RApp(d); app.Sel = d.Mk(0, 5); wn = new WordNumbers(app);
            try { wn.Apply(st4, "inc"); Eq(false, true, "phải báo không có số"); } catch (UserMsg) { Eq(true, true, "không có số thì báo"); }
        }
        // === Rev0.7: Đọc số (viết số bằng chữ ngay sau số), đồng bộ 0 đến 5 số lẻ, kiểu kỹ thuật
        {
            var q7 = new AppSettings(); q7.NumStyle = "vn";
            // con trỏ trong số
            d = new RDoc("Áp suất 1.234 kPa.\r"); app = new RApp(d); app.Sel = d.Mk(11, 11); var w7 = new WordNumbers(app);
            string r7 = w7.Read(q7); Console.WriteLine("  -> " + r7 + " | " + d.T);
            Eq(d.T.ToString(), "Áp suất 1.234 kPa (một nghìn hai trăm ba mươi bốn kilôpascan).\r", "đọc số kèm đơn vị: chữ nằm ngay sau đơn vị, trong ngoặc");
            Eq(app.Sel.S == app.Sel.E && app.Sel.S == d.T.ToString().IndexOf("kilôpascan)") + "kilôpascan)".Length, true, "con trỏ nằm sau phần vừa viết");
            // bấm lần hai không thêm lần hai
            r7 = w7.Read(q7); Console.WriteLine("  -> " + r7);
            Eq(d.T.ToString().Split(new[] { "một nghìn" }, StringSplitOptions.None).Length - 1, 1, "bấm lại không viết thêm");
            // con trỏ ngay sau số
            d = new RDoc("Có 25 máy.\r"); app = new RApp(d); app.Sel = d.Mk(4, 4); w7 = new WordNumbers(app);
            w7.Read(q7); Eq(d.T.ToString(), "Có 25 (hai mươi lăm) máy.\r", "con trỏ sát cuối số cũng đọc");
            // bôi đen nhiều số
            d = new RDoc("A 3,05 và 100 và -7.\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1); w7 = new WordNumbers(app);
            r7 = w7.Read(q7); Console.WriteLine("  -> " + r7 + " | " + d.T);
            Eq(d.T.ToString(), "A 3,05 (ba phẩy không năm) và 100 (một trăm) và -7 (âm bảy).\r", "nhiều số cùng lúc, từ cuối lên đầu không lệch vị trí");
            // số kỹ thuật
            d = new RDoc("Tải 2,3 × 10⁶ N.\r"); app = new RApp(d); app.Sel = d.Mk(6, 6); w7 = new WordNumbers(app);
            w7.Read(q7); Eq(d.T.ToString(), "Tải 2,3 × 10⁶ N (hai phẩy ba nhân mười mũ sáu niutơn).\r", "đọc số kỹ thuật kèm đơn vị");
            // con trỏ không ở số
            d = new RDoc("Không có số\r"); app = new RApp(d); app.Sel = d.Mk(3, 3); w7 = new WordNumbers(app);
            try { w7.Read(q7); Eq(false, true, "phải báo không có số"); } catch (UserMsg) { Eq(true, true, "con trỏ ngoài số thì báo"); }
            // === Rev0.9: đọc kèm đơn vị
            {
                var qu = new AppSettings(); qu.NumStyle = "vn";
                d = new RDoc("Trả 1.000.000 VNĐ và 20 kPa.\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1); var wu = new WordNumbers(app);
                string ru = wu.Read(qu); Console.WriteLine("  -> " + ru + " | " + d.T);
                Eq(d.T.ToString(), "Trả 1.000.000 VNĐ (một triệu Việt Nam đồng) và 20 kPa (hai mươi kilôpascan).\r", "bôi đen: số kèm tiền và áp suất");
                string once = d.T.ToString(); app.Sel = d.Mk(0, d.T.Length - 1); wu.Read(qu);
                Eq(d.T.ToString(), once, "bấm lại không viết thêm (có đơn vị)");
                // tên thông dụng
                qu.UnitNames = "common";
                d = new RDoc("Áp suất 20 kPa, lực 5 N.\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1); wu = new WordNumbers(app);
                wu.Read(qu); Eq(d.T.ToString(), "Áp suất 20 kPa (hai mươi kilopascal), lực 5 N (năm newton).\r", "tên thông dụng: kilopascal, newton");
                qu.UnitNames = "law";
                // con trỏ trong số, trong đơn vị, và sau phần đọc
                d = new RDoc("p = 20 kPa tại đây\r"); app = new RApp(d); app.Sel = d.Mk(5, 5); wu = new WordNumbers(app);
                wu.Read(qu); Eq(d.T.ToString(), "p = 20 kPa (hai mươi kilôpascan) tại đây\r", "con trỏ trong số: chèn sau đơn vị");
                d = new RDoc("p = 20 kPa tại đây\r"); app = new RApp(d); app.Sel = d.Mk(8, 8); wu = new WordNumbers(app);
                wu.Read(qu); Eq(d.T.ToString(), "p = 20 kPa (hai mươi kilôpascan) tại đây\r", "con trỏ trong đơn vị cũng đọc số phía trước");
                d = new RDoc("p = 20 kPa tại đây\r"); app = new RApp(d); app.Sel = d.Mk(5, 5); wu = new WordNumbers(app);
                wu.Read(qu); wu.Read(qu); Eq(d.T.ToString(), "p = 20 kPa (hai mươi kilôpascan) tại đây\r", "bấm hai lần khi con trỏ đã ở sau phần đọc: không thêm");
                // số không có đơn vị vẫn đọc như cũ, từ thường không bị nhầm là đơn vị
                d = new RDoc("Có 25 máy và 3 người.\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1); wu = new WordNumbers(app);
                wu.Read(qu); Eq(d.T.ToString(), "Có 25 (hai mươi lăm) máy và 3 (ba) người.\r", "chữ thường sau số không bị coi là đơn vị");
                // đơn vị ghép và tiền ký hiệu đứng trước
                d = new RDoc("v = 12 m/s, giá $100.\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1); wu = new WordNumbers(app);
                wu.Read(qu); Eq(d.T.ToString(), "v = 12 m/s (mười hai mét trên giây), giá $100 (một trăm đô la Mỹ).\r", "đơn vị ghép và ký hiệu $");
                // bôi đen chỉ một phần: đơn vị ngoài vùng chọn không được đọc
                d = new RDoc("20 kPa\r"); app = new RApp(d); app.Sel = d.Mk(0, 2); wu = new WordNumbers(app);
                wu.Read(qu); Eq(d.T.ToString(), "20 (hai mươi) kPa\r", "bôi đen chỉ số 20 thì chỉ đọc số");
            }
            // đồng bộ 0 đến 5 số lẻ
            for (int k = 0; k <= 5; k++)
            {
                d = new RDoc("x 3,14159 y 2,5 z\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1); w7 = new WordNumbers(app);
                q7.SyncSci = false; q7.SyncDec = k; w7.Apply(q7, "sync");
                string[] a = { "3", "3,1", "3,14", "3,142", "3,1416", "3,14159" }, b = { "3", "2,5", "2,50", "2,500", "2,5000", "2,50000" };
                Eq(d.T.ToString(), "x " + a[k] + " y " + b[k] + " z\r", "đồng bộ " + k + " số lẻ trong Word");
                Eq(app.Sel.S == 0 && app.Sel.E == d.T.Length - 1, true, "giữ vùng bôi đen sau đồng bộ " + k);
            }
            // kiểu kỹ thuật
            d = new RDoc("Lực 2300000 N và 0,00367 m.\r"); app = new RApp(d); app.Sel = d.Mk(0, d.T.Length - 1); w7 = new WordNumbers(app);
            q7.SyncSci = true; q7.SciDec = 1; r7 = w7.Apply(q7, "sync"); Console.WriteLine("  -> " + r7 + " | " + d.T);
            Eq(d.T.ToString(), "Lực 2,3\u00A0×\u00A010⁶ N và 3,7\u00A0×\u00A010⁻³ m.\r", "kiểu kỹ thuật trong văn bản");
            Eq(app.Sel.S == 0 && app.Sel.E == d.T.Length - 1, true, "giữ vùng bôi đen sau kiểu kỹ thuật");
            q7.SciDec = 3; w7.Apply(q7, "sync"); Eq(d.T.ToString(), "Lực 2,300\u00A0×\u00A010⁶ N và 3,700\u00A0×\u00A010⁻³ m.\r", "chạy lại với số lẻ khác: viết lại, không lồng thêm ×10");
            q7.SyncSci = false; q7.SyncDec = 1; w7.Apply(q7, "sync"); Eq(d.T.ToString(), "Lực 2,3\u00A0×\u00A010⁶ N và 3,7\u00A0×\u00A010⁻³ m.\r", "quay lại đồng bộ thường: chỉ chỉnh phần đầu");
        }
        // === Rev0.8: lỗi nghiêm trọng làm đổi giá trị số (video 16:08). Con trỏ trong số, bấm liên tiếp nhiều lệnh định dạng, giá trị phải giữ nguyên.
        {
            var q8 = new AppSettings(); q8.NumStyle = "vn";
            d = new RDoc("làm t 48,5. Lưu chất\r"); app = new RApp(d); app.Sel = d.Mk(8, 8); var w8 = new WordNumbers(app);
            w8.Read(q8);
            Eq(d.T.ToString(), "làm t 48,5 (bốn mươi tám phẩy năm). Lưu chất\r", "tái hiện: Đọc số trên 48,5");
            // đặt con trỏ lại trong số (như người dùng bấm vào số) rồi bấm Thêm số lẻ nhiều lần
            app.Sel = d.Mk(8, 8);
            string[] want = { "48,50", "48,500", "48,5000", "48,50000" };
            for (int k = 0; k < 4; k++)
            {
                w8.Apply(q8, "inc");
                Eq(d.T.ToString().StartsWith("làm t " + want[k] + " (bốn"), true, "Thêm số lẻ lần " + (k + 1) + " giữ giá trị: " + d.T);
                Eq(app.Sel.S == app.Sel.E, true, "sau lệnh vẫn là con trỏ, không bôi đen một ký tự (lần " + (k + 1) + ")");
            }
            for (int k = 0; k < 4; k++) w8.Apply(q8, "dec");
            Eq(d.T.ToString().StartsWith("làm t 48,5 (bốn"), true, "Bớt số lẻ bốn lần về lại 48,5: " + d.T);
            // luân chuyển kiểu dấu nhiều lần không đổi giá trị
            string before = d.T.ToString();
            for (int k = 0; k < 6; k++) { w8.Apply(q8, "cycle"); Eq(app.Sel.S == app.Sel.E, true, "Dấu: vẫn là con trỏ"); }
            Eq(d.T.ToString(), before.Replace("48,5 ", "48,5 "), "Dấu qua đủ vòng về như cũ: " + d.T);
            // số nhóm nghìn
            d = new RDoc("tải 4850000 N\r"); app = new RApp(d); app.Sel = d.Mk(6, 6); w8 = new WordNumbers(app);
            q8.NumThou = true;
            for (int k = 0; k < 3; k++) w8.Apply(q8, "cycle");
            Eq(d.T.ToString().Contains("4850000") || d.T.ToString().Contains("4.850.000") || d.T.ToString().Contains("4,850,000") || d.T.ToString().Contains("4 850 000") || d.T.ToString().Contains("4\u00A0850\u00A0000"), true, "Dấu giữ giá trị 4850000: " + d.T);
            // nếu văn bản đã hỏng sẵn thì không được biến thành số khác
            d = new RDoc("làm t 48,5,000 mm\r"); app = new RApp(d); app.Sel = d.Mk(8, 8); w8 = new WordNumbers(app);
            try { w8.Apply(q8, "cycle"); } catch (UserMsg) { }
            Eq(d.T.ToString(), "làm t 48,5,000 mm\r", "chuỗi 48,5,000 không phải số nên không bị sửa");
        }
        // === Rev0.5: Chuẩn hóa nhanh tại con trỏ, sửa và bỏ qua mục tương tự
        {
            var q = new AppSettings(); q.EqMode = "fmt"; q.RvNumStyle = "us"; q.EqStacked = true;
            d = new RDoc("Gia tri R_cyl va R_{cyl} va t_req o day.\r"); app = new RApp(d); var wr = new WordReview(app);
            string t0 = d.T.ToString(); int a0 = t0.IndexOf("R_cyl") + 3;
            app.Sel = d.Mk(a0, a0);
            string qm = wr.QuickFix(q); Console.WriteLine("  -> " + qm);
            string t1 = d.T.ToString(); int r1 = t1.IndexOf("Rcyl");
            Eq(r1 >= 0 && d.Sub[r1 + 1] && d.Sub[r1 + 3] && !d.Sub[r1], true, "con trỏ trong R_cyl: bỏ _, cyl là chỉ số dưới, R không");
            Eq(t1.Contains("R_{cyl}") && t1.Contains("t_req"), true, "chỉ sửa chỗ có con trỏ, các chỗ khác giữ nguyên");
            // bôi đen cả đoạn: sửa mọi chỗ
            app.Sel = d.Mk(0, d.T.Length - 1);
            string qm2 = wr.QuickFix(q); Console.WriteLine("  -> " + qm2);
            Eq(d.T.ToString().Contains("_") || d.T.ToString().Contains("{"), false, "bôi đen cả đoạn: hết _ và { }");
            string qm3 = wr.QuickFix(q); Eq(qm3.Length > 0, true, "chạy lại không lỗi: " + qm3);
            // nhiều chữ: v_2 không thuộc thư viện nhưng nút Chuẩn hóa vẫn đọc như LaTeX
            d = new RDoc("Toc do v_2 la gi.\r"); app = new RApp(d); wr = new WordReview(app);
            int v0 = d.T.ToString().IndexOf("v_2") + 1; app.Sel = d.Mk(v0, v0);
            wr.QuickFix(q); string t2 = d.T.ToString(); int v1 = t2.IndexOf("v2");
            Eq(v1 >= 0 && d.Sub[v1 + 1] && !t2.Contains("_"), true, "v_2: chuyển thành v với chỉ số dưới 2");
            // không thấy gì tại con trỏ
            d = new RDoc("Chu thuong khong loi.\r"); app = new RApp(d); wr = new WordReview(app); app.Sel = d.Mk(3, 3);
            Eq(wr.QuickFix(q).Contains("Không thấy lỗi"), true, "chữ thường: báo không thấy lỗi");
            // công thức riêng một dòng có phân số: dựng Equation (OMML)
            d = new RDoc("Re_{cyl} = \\frac{a b}{c}\r"); app = new RApp(d); wr = new WordReview(app); app.Sel = d.Mk(2, 2);
            wr.QuickFix(q); Eq(RRange.LastXml.Contains("<m:f>") && d.Eq[0], true, "dòng riêng có \\frac: Equation OMML xếp chồng");
            // sửa tương tự
            d = new RDoc("A Fe2+ and Fe2+ again and H2S only.\r"); app = new RApp(d); wr = new WordReview(app);
            var q2 = new AppSettings(); q2.RvNumStyle = "us"; q2.EqMode = "fmt";
            RvResult rr = wr.Scan(q2);
            int ionIdx = rr.Rows.First(x => x.Before == "Fe2+").Index;
            Eq(wr.SimilarCount(ionIdx), 2, "hai chỗ Fe2+ giống nhau");
            wr.FixSimilar(ionIdx, 0, q2);
            Eq(d.T.ToString().Contains("fe2+") || d.T.ToString().Contains("Fe2+"), true, "đã sửa");
            Eq(wr.Pending().Rows.Count, rr.Rows.Count - 2, "sửa tương tự: hai mục cùng biến mất, mục khác còn lại");
            rr = wr.Scan(q2); Eq(rr.Rows.Count, 1, "quét lại: chỉ còn H2S chưa sửa, hai Fe2+ đã sửa không báo lại");
            d = new RDoc("x H2S and H2S and Fe2+.\r"); app = new RApp(d); wr = new WordReview(app);
            rr = wr.Scan(q2); int hs = rr.Rows.First(x => x.Before == "H2S").Index;
            wr.IgnoreSimilar(hs);
            Eq(wr.Pending().Rows.Any(x => x.Before == "H2S"), false, "bỏ qua tương tự: cả hai H2S không còn trong danh sách");
            Eq(wr.Pending().Rows.Any(x => x.Before == "Fe2+"), true, "mục khác vẫn còn");
            // Rev0.6: số cùng dạng nhưng khác chữ số (5.678 và 12.345) được gộp là "tương tự"; dạng khác thì không
            d = new RDoc("Mã A 5.678 và 12.345 và 7.5 mm xong.\r"); app = new RApp(d); wr = new WordReview(app);
            var qn = new AppSettings(); qn.NumStyle = "vn"; qn.RvNumStyle = "vn"; qn.NumDec = "keep"; qn.NumThou = true; qn.EqMode = "fmt";
            rr = wr.Scan(qn); int first = rr.Rows.First(x => x.Before == "5.678").Index;
            Eq(wr.SimilarCount(first), 2, "5.678 và 12.345 cùng dạng hàng nghìn: tương tự nhau");
            string im = wr.IgnoreSimilar(first); Console.WriteLine("  -> " + im);
            Eq(wr.Pending().Rows.Any(x => x.Before == "5.678" || x.Before == "12.345"), false, "bỏ qua tương tự: cả hai số khác chữ số cùng biến mất");
            Eq(im.Contains("cùng dạng"), true, "thông báo nói rõ là cùng dạng, không phải cùng chữ");
            d = new RDoc("Mã A 5.678 và 12.345 xong.\r"); app = new RApp(d); wr = new WordReview(app);
            rr = wr.Scan(qn); first = rr.Rows.First(x => x.Before == "5.678").Index;
            string fx = wr.FixSimilar(first, 0, qn); Console.WriteLine("  -> " + fx);
            Eq(d.T.ToString().Contains("5,678") && d.T.ToString().Contains("12,345"), true, "sửa tương tự: mỗi số được sửa theo giá trị riêng của nó");
        }
        // AI transport is mocked elsewhere. Verify exact Word ranges, stale snapshots and safe application.
        {
            var dAi = new RDoc("DO NOT SEND.\rThey believes.\rLast sentence.\r");
            var appAi = new RApp(dAi);
            var reviewer = new WordReview(appAi) { Mode = "gram" };
            appAi.Sel = dAi.Mk(13, 27);
            var selected = reviewer.CaptureAi(true);
            Eq(selected.Blocks.Count, 1, "AI quét đúng một đoạn bôi đen");
            Eq(selected.Blocks[0].Text, "They believes.", "AI không mở rộng phạm vi bôi đen");
            var whole = reviewer.CaptureAi(false);
            Eq(whole.Blocks.Count, 3, "AI quét cả ba đoạn phần thân");
            var dTable = new RDoc("First cell\rLast cell\rAfter table.\r");
            dTable.CellEnds.Add(10);
            dTable.CellEnds.Add(20);
            var tableReview = new WordReview(new RApp(dTable)) { Mode = "gram" };
            var table = tableReview.CaptureAi(false);
            Eq(table.Blocks[1].Text, "Last cell", "quét được các ô trong bảng Word");
            Eq(table.Blocks[1].Start, 11, "ký tự ẩn cuối ô không chiếm vị trí Word");
            var tableIssue = new Issue { Cat = "ai-gram", Sev = "err", Before = "Last cell", Key = "table", Reason = "kiểm thử" };
            var rows = tableReview.AcceptAi(table, new List<AiFinding> { new AiFinding { Start = 11, Issue = tableIssue } });
            Eq(rows.Rows.Count, 1, "định vị đúng mục gợi ý trong ô sau");
            var issue = new Issue { Before = "They believes", Cat = "ai-gram", Sev = "err", Key = "agree", Reason = "Chủ vị" };
            issue.Fixes.Add(Fix.Text("Sửa", "They believe", true));
            var findings = new List<AiFinding> { new AiFinding { Start = 13, Issue = issue } };
            var settings = new AppSettings();
            RvResult result = reviewer.AcceptAi(selected, findings);
            Eq(result.Rows.Count, 1, "nhận đúng lỗi trong vùng Word");
            reviewer.FixItem(result.Rows[0].Index, 0, settings);
            Eq(dAi.T.ToString().Contains("They believe."), true, "sửa đúng vùng Word");
            try { reviewer.AcceptAi(whole, findings); Eq(false, true, "chặn ảnh chụp cũ"); }
            catch (UserMsg) { Eq(true, true, "chặn ảnh chụp cũ"); }
            appAi.Sel = dAi.Mk(0, 0);
            try { reviewer.CaptureAi(true); Eq(false, true, "chặn quét vùng chọn rỗng"); }
            catch (UserMsg) { Eq(true, true, "chặn quét vùng chọn rỗng"); }

            var dRw = new RDoc("Prefix ORIGINAL suffix.\r");
            var appRw = new RApp(dRw);
            var words = new WordText(appRw);
            appRw.Sel = dRw.Mk(7, 15);
            var snap = words.GetRewriteSnapshot();
            appRw.Sel = dRw.Mk(0, 0);
            words.ApplyRewrite(snap, "RESULT");
            Eq(dRw.T.ToString(), "Prefix RESULT suffix.\r", "kết quả chỉ thay vùng từng bôi đen");
            try { words.ApplyRewrite(snap, "OTHER"); Eq(false, true, "chặn áp dụng kết quả cũ"); }
            catch (UserMsg) { Eq(true, true, "chặn áp dụng kết quả cũ"); }
            var multi = new RDoc("Before first\r\rsecond after.\r");
            var multiApp = new RApp(multi);
            var multiWords = new WordText(multiApp);
            multiApp.Sel = multi.Mk(7, 20);
            var multiSnap = multiWords.GetRewriteSnapshot();
            multiWords.ApplyRewrite(multiSnap, "A\n\nB");
            Eq(multi.T.ToString(), "Before A\r\rB after.\r", "giữ ngắt đoạn và đoạn trống");
        }
        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}
