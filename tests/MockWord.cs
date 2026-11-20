// Mô hình Word giả (chỉ để kiểm thử logic gọi Word, không thay thế việc thử trong Word thật).
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using TroLyKyThuat;

public class Rec : DynamicObject
{
    public static List<string> Log = new List<string>();
    public string RecPath; public Dictionary<string, object> P = new Dictionary<string, object>();
    public Rec(string n) { RecPath = n; }
    public override bool TryGetMember(GetMemberBinder b, out object r)
    {
        if (P.TryGetValue(b.Name, out r)) return true;
        r = new Rec(RecPath + "." + b.Name); P[b.Name] = r; return true;
    }
    public override bool TrySetMember(SetMemberBinder b, object v) { P[b.Name] = v; Log.Add(RecPath + "." + b.Name + "=" + v); return true; }
    public override bool TryInvokeMember(InvokeMemberBinder b, object[] a, out object r)
    {
        string key = b.Name + "(" + string.Join(",", a.Select(x => x == null ? "null" : x.ToString()).ToArray()) + ")";
        Log.Add(RecPath + "." + key);
        object memo;
        if (!P.TryGetValue("()" + key, out memo)) { memo = new Rec(RecPath + "." + key); P["()" + key] = memo; }
        r = memo; return true;
    }
}

public class MCell : Rec
{
    public MTable T; public int R, C;
    public MCell(MTable t, int r, int c) : base("cell") { T = t; R = r; C = c; }
    public int RowIndex { get { return R; } }
    public int ColumnIndex { get { return C; } }
    public double Width { get { return T.WidthOf(R, C); } }
    MRange rng;
    public MRange Range { get { if (rng == null) rng = new MRange(this); return rng; } }
    Rec shading;
    public Rec Shading { get { if (shading == null) shading = new Rec("shading"); return shading; } }
}
public class MRange : Rec
{
    MCell cell; public MTable tbl;
    public MRange(MCell c) : base("range") { cell = c; }
    public MRange(MTable t) : base("range") { tbl = t; }
    public string Text
    {
        get { return cell.T.D[cell.R - 1][cell.C - 1] + "\r\a"; }
        set { cell.T.D[cell.R - 1][cell.C - 1] = value; }
    }
    public int Start { get { return cell != null ? cell.R * 100 + cell.C : 0; } }
    public int End { get { return cell != null ? cell.R * 100 + cell.C + 1 : 99999; } }
    MCells cs; int csRows = -1, csCells = -1;
    public MCells Cells   // Rev0.9.5.13: làm mới khi bảng thêm hàng hoặc ô (bộ sưu tập sống như Word), giữ nguyên đối tượng khi bảng không đổi
    {
        get
        {
            int n = 0; foreach (var r in tbl.D) n += r.Count;
            if (cs == null || csRows != tbl.D.Count || csCells != n) { cs = new MCells(tbl, null); csRows = tbl.D.Count; csCells = n; }
            return cs;
        }
    }
}
public class MCells : Rec
{
    List<MCell> l;
    public MCells(MTable t, List<MCell> sel) : base("cells")
    {
        if (sel != null) l = sel;
        else { l = new List<MCell>(); for (int r = 1; r <= t.D.Count; r++) for (int c = 1; c <= t.D[r - 1].Count; c++) l.Add(t.Cell(r, c)); }
    }
    public int Count { get { return l.Count; } }
    public MCell Item(int i) { return l[i - 1]; }
    public Rec Shading = new Rec("cells.shading");
}
public class MRows : Rec
{
    MTable t; public MRows(MTable t) : base("rows") { this.t = t; }
    public int Count { get { return t.D.Count; } }
    readonly Dictionary<int, MRowRec> cache = new Dictionary<int, MRowRec>();
    public Rec Item(int i) { MRowRec r; if (!cache.TryGetValue(i, out r)) { r = new MRowRec(t, i); cache[i] = r; } return r; }
    public object Add(object before)
    {
        int idx = (before is MRowRec) ? ((MRowRec)before).I - 1 : t.D.Count;
        var model = (before is MRowRec) ? t.D[((MRowRec)before).I - 1] : t.D[t.D.Count - 1];
        t.D.Insert(idx, new List<string>(model.Select(x => "")));
        if (t.Wd != null) { var wm = (before is MRowRec) ? t.Wd[((MRowRec)before).I - 1] : t.Wd[t.Wd.Count - 1]; t.Wd.Insert(idx, new List<double>(wm)); }
        return new Rec("row");
    }
    public object Add() { var model = t.D[t.D.Count - 1]; t.D.Add(new List<string>(model.Select(x => ""))); if (t.Wd != null) t.Wd.Add(new List<double>(t.Wd[t.Wd.Count - 1])); return new Rec("row"); }
}
public class MRowRec : Rec
{
    public MTable T; public int I; public MRowRec(MTable t, int i) : base("row" + i) { T = t; I = i; }
    public void Delete() { T.D.RemoveAt(I - 1); }
}
public class MCols : Rec
{
    MTable t; public MCols(MTable t) : base("cols") { this.t = t; }
    public int Count { get { return t.D[0].Count; } }   // đã đo trên Word 16 thật: không báo lỗi với hàng gộp ngang, trả số ô của hàng đầu
    public Rec Item(int i) { if (t.Irregular) throw new InvalidOperationException("Cannot access individual columns because the table has mixed cell widths."); return new MColRec(t, i); }
    public object Add(object before)
    {
        int idx = (before is MColRec) ? ((MColRec)before).I - 1 : t.D[0].Count;
        foreach (var r in t.D) r.Insert(idx, "");
        return new Rec("col");
    }
    public object Add() { if (t.Irregular) throw new InvalidOperationException("Cannot access individual columns because the table has mixed cell widths."); foreach (var r in t.D) r.Add(""); return new Rec("col"); }
}
public class MColRec : Rec
{
    public MTable T; public int I; public MColRec(MTable t, int i) : base("col" + i) { T = t; I = i; }
    public void Delete() { foreach (var r in T.D) r.RemoveAt(I - 1); }
}
public class MTable : Rec
{
    public List<List<string>> D = new List<List<string>>();
    public MTable(string[][] rows) : base("table") { foreach (var r in rows) D.Add(new List<string>(r)); }
    // Rev0.9.5.13: độ rộng từng ô (pt) để mô phỏng ô gộp; mặc định mỗi ô rộng 100 pt
    public List<List<double>> Wd;
    public double WidthOf(int r, int c) { return (Wd != null && r <= Wd.Count && c <= Wd[r - 1].Count) ? Wd[r - 1][c - 1] : 100.0; }
    public bool Irregular { get { foreach (var r in D) if (r.Count != D[0].Count) return true; return false; } }
    public bool Uniform { get { return !Irregular; } }
    MRows rows; MRange rng;
    public MRows Rows { get { if (rows == null) rows = new MRows(this); return rows; } }
    object style;
    public object Style { get { return style; } set { Log.Add("table.Style=" + value); style = new StyleRec((string)value); } }
    public MCols Columns { get { return new MCols(this); } }
    readonly Dictionary<long, MCell> cc = new Dictionary<long, MCell>();
    public MCell Cell(int r, int c)
    {
        if (r < 1 || r > D.Count || c < 1 || c > D[r - 1].Count) throw new ArgumentException("khong co o");
        long k = (long)r * 1000 + c; MCell m;
        if (!cc.TryGetValue(k, out m)) { m = new MCell(this, r, c); cc[k] = m; }
        return m;
    }
    public MRange Range { get { if (rng == null) rng = new MRange(this); return rng; } }
}
public class StyleRec : Rec { public string NameLocal; public StyleRec(string n) : base("style") { NameLocal = n; } }
public class MSel : Rec
{
    public MTable T; public int R1, C1, R2, C2; public int[] RowList;
    public MSel(MTable t, int r1, int c1, int r2, int c2) : base("sel") { T = t; R1 = r1; C1 = c1; R2 = r2; C2 = c2; }
    public MSel(MTable t, int[] rows, int c1, int c2) : base("sel") { T = t; RowList = rows; R1 = rows[0]; R2 = rows[rows.Length - 1]; C1 = c1; C2 = c2; }
    public bool Information(int k) { return true; }
    public MSelRange Range { get { int a = (RowList != null ? RowList.Sum() * 1000 : 0) + R1 * 100 + C1; return new MSelRange(a, R2 * 100 + C2 + 1); } }
    public int Type { get { return 5; } }
    public Rec Tables { get { var r = new Rec("tables"); r.P["Count"] = 1; return new MTablesOne(T); } }
    public MCells Cells
    {
        get
        {
            var l = new List<MCell>();
            if (RowList != null) { foreach (int r in RowList) for (int c = C1; c <= C2; c++) l.Add(T.Cell(r, c)); }
            else for (int r = R1; r <= R2; r++) for (int c = C1; c <= C2; c++) l.Add(T.Cell(r, c));
            return new MCells(T, l);
        }
    }
}
public class MSelRange : Rec { public int Start, End; public MSelRange(int s, int e) : base("selrange") { Start = s; End = e; } }
public class MTablesOne : Rec
{
    MTable t; public MTablesOne(MTable t) : base("tables") { this.t = t; }
    public int Count { get { return 1; } }
    public MTable Item(int i) { return t; }
}
public class MDocs : Rec { public MDocs() : base("docs") { } public int Count { get { return 1; } } }
public class MDoc : Rec { MTable t; public MDoc(MTable t) : base("doc") { this.t = t; } public MTablesOne Tables { get { return new MTablesOne(t); } } }
public class MApp : Rec
{
    public MTable T; public MSel S;
    public MApp(MTable t, MSel s) : base("app") { T = t; S = s; }
    public MDocs Documents { get { return new MDocs(); } }
    MDoc doc;
    public MDoc ActiveDocument { get { if (doc == null) doc = new MDoc(T); return doc; } }
    public MSel Selection { get { return S; } }
    public Rec UndoRecord { get { return new Rec("undo"); } }
}

public class MTablesN : Rec
{
    List<MTable> l; public MTablesN(IEnumerable<MTable> t) : base("tables") { l = new List<MTable>(t); }
    public int Count { get { return l.Count; } }
    public MTable Item(int i) { return l[i - 1]; }
}
public class MDoc2 : Rec { MTable[] t; public MDoc2(MTable[] t) : base("doc") { this.t = t; } public MTablesN Tables { get { return new MTablesN(t); } } }
public class MSel2 : Rec { MTable[] t; public MSel2(MTable[] t) : base("sel") { this.t = t; } public MTablesN Tables { get { return new MTablesN(t); } } }
public class MApp2 : Rec
{
    MDoc2 d; MSel2 s;
    public MApp2(MTable[] all, MTable[] sel) : base("app") { d = new MDoc2(all); s = new MSel2(sel); }
    public MDocs Documents { get { return new MDocs(); } }
    public MDoc2 ActiveDocument { get { return d; } }
    public MSel2 Selection { get { return s; } }
    public Rec UndoRecord { get { return new Rec("undo"); } }
}

public class MPageSetup
{
    public int PaperSize;
    public int Orientation; // 0 dọc, 1 ngang
    public float TopMargin, BottomMargin, LeftMargin, RightMargin, HeaderDistance, FooterDistance;
}
public class MLayoutSection
{
    public int Index;
    public MPageSetup PageSetup = new MPageSetup();
    public MLayoutSection(int idx, int orient) { Index = idx; PageSetup.Orientation = orient; }
}
public class MLayoutSections
{
    public List<MLayoutSection> Items = new List<MLayoutSection>();
    public int Count { get { return Items.Count; } }
    public MLayoutSection Item(int i) { return Items[i - 1]; }
}
public class MParaFont
{
    public string Name;
    public float Size;
}
public class MListFormat
{
    public int ListType;
}
public class MParagraphStyle
{
    public bool AutomaticallyUpdate;
    public string NameLocal;
}
public class MParaRange
{
    public MParagraphStyle Style = new MParagraphStyle();
    public MParaFont Font = new MParaFont();
    public bool InTable;
    public MListFormat ListFormat = new MListFormat();
    public MLayoutParas Paragraphs = new MLayoutParas();
    public MParaRange Duplicate { get { return this; } }
    public void Collapse(int direction) { }
    public bool Information(int k) { return k == 12 && InTable; }
}
public class MParaFormat
{
    int alignment;
    public bool FailAlignment;
    public int Alignment
    {
        get { return alignment; }
        set { if (FailAlignment) throw new InvalidOperationException("Simulated paragraph formatting failure"); alignment = value; }
    }
    public int LineSpacingRule;
    public float LineSpacing, SpaceBefore, SpaceAfter;
    public int OutlineLevel = 10;
    public float LeftIndent;
    public float FirstLineIndent;
}
public class MLayoutPara
{
    public MParaRange Range = new MParaRange();
    public MParaFormat Format = new MParaFormat();
    public MLayoutPara(string fontName, float fontSize, bool inTable)
    {
        Range.Font.Name = fontName; Range.Font.Size = fontSize; Range.InTable = inTable;
    }
    public MLayoutPara(string fontName, float fontSize, bool inTable, int outlineLevel, int listType)
    {
        Range.Font.Name = fontName; Range.Font.Size = fontSize; Range.InTable = inTable;
        Format.OutlineLevel = outlineLevel;
        Range.ListFormat.ListType = listType;
    }
}
public class MLayoutParas
{
    public List<MLayoutPara> Items = new List<MLayoutPara>();
    public int Count { get { return Items.Count; } }
    public MLayoutPara Item(int i) { return Items[i - 1]; }
}
public class MLayoutDoc
{
    public MLayoutSections Sections = new MLayoutSections();
    public MLayoutParas Paragraphs = new MLayoutParas();
}
public class MLayoutSel
{
    public MLayoutDoc Doc;
    public int Start, End;
    public int CurPage = 1, TotalPages = 1;
    public int CurSecIdx = 1;
    public MLayoutParas Paragraphs = new MLayoutParas();
    public bool StartIsActive;
    public MParaRange Range
    {
        get
        {
            var target = new MParaRange();
            if (Paragraphs.Count > 0) target.Paragraphs.Items.Add(Paragraphs.Item(StartIsActive ? 1 : Paragraphs.Count));
            else if (Doc.Paragraphs.Count > 0) target.Paragraphs.Items.Add(Doc.Paragraphs.Item(1));
            return target;
        }
    }
    public MLayoutSel(MLayoutDoc d) { Doc = d; }
    public MLayoutSections Sections
    {
        get
        {
            var s = new MLayoutSections();
            s.Items.Add(Doc.Sections.Items[CurSecIdx - 1]);
            return s;
        }
    }
    public int Information(int k) { return k == 3 ? CurPage : (k == 4 ? TotalPages : 0); }
    public void GoTo(int what, int which, int count) { CurPage = count; Start += 100; End = Start; }
    public void EndKey(int unit) { Start += 500; End = Start; }
    public void MoveLeft(int unit, int count) { }
    public void InsertBreak(int kind)
    {
        if (kind == 2)
        {
            int curOrient = Doc.Sections.Items[CurSecIdx - 1].PageSetup.Orientation;
            var ns = new MLayoutSection(CurSecIdx + 1, curOrient);
            Doc.Sections.Items.Insert(CurSecIdx, ns);
            for (int i = 0; i < Doc.Sections.Items.Count; i++) Doc.Sections.Items[i].Index = i + 1;
            CurSecIdx++;
            Start += 10; End = Start;
        }
        else if (kind == 7) { TotalPages++; CurPage++; }
    }
    public void SetRange(int s, int e) { Start = s; End = e; }
}
public class MLayoutApp
{
    public MDocs Documents { get { return new MDocs(); } }
    public MLayoutDoc ActiveDocument = new MLayoutDoc();
    public MLayoutSel Selection;
    public Rec UndoRecord { get { return new Rec("undo"); } }
    public MLayoutApp() { Selection = new MLayoutSel(ActiveDocument); }
}
