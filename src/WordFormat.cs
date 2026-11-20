// Trợ lý Kỹ thuật - Rev0.4: hai lệnh làm việc trên vùng đang bôi đen: Định dạng số (kiểu dấu, chữ số thập phân) và Kiểu chỉ số (công thức toán, lý, hóa).
// Mọi hàm ở đây phải chạy trên luồng chính của Word.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    /// <summary>Định dạng số cho bất kỳ số nào đang bôi đen: dãy ô trong bảng, hoặc các số nằm trong văn bản. Bôi đen nhiều số thì đổi cùng lúc.</summary>
    public sealed class WordNumbers
    {
        readonly dynamic app;
        public WordNumbers(object application) { app = application; }

        static string StripEnd(string t)
        {
            if (t == null) return "";
            return t.TrimEnd('\r', '\n', '\a');
        }

        /// <summary>mode: "cycle" (đổi kiểu dấu theo vòng Việt, Mỹ, SI), "inc" (tăng một chữ số thập phân), "dec" (giảm một chữ số thập phân),
        /// "sync" (đồng bộ: mọi số có đúng st.SyncDec chữ số thập phân). Mọi chế độ giữ nguyên vùng bôi đen để bấm tiếp.</summary>
        public string Apply(AppSettings st, string mode)
        {
            try { if ((int)app.Documents.Count == 0) throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            catch (UserMsg) { throw; }
            catch { throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            var o = new NumEditOpts();
            o.Thousands = st.NumThou; o.FallbackStyle = st.NumStyle;
            string prev = st.NumStyle, newStyle = null;
            if (mode == "cycle") { newStyle = NumberEdit.NextStyle(prev); o.Style = newStyle; }
            else if (mode == "sync")
            {
                if (st.SyncSci) { o.Sci = true; o.SciDec = Math.Max(0, Math.Min(5, st.SciDec)); }
                else o.SetDec = Math.Max(0, Math.Min(5, st.SyncDec));
            }
            else o.DecDelta = mode == "inc" ? 1 : -1;
            var w = new WordTables(app);
            SelInfo s = w.GetSelection(false);
            string msg;
            if (s.InTable && CellMode(s)) msg = ApplyCells(w, s, o, mode);
            else msg = ApplyInline(o, mode);
            if (newStyle != null) { st.NumStyle = newStyle; st.Save(); }
            return msg;
        }

        /// <summary>Trong bảng: đổi theo ô khi bôi đen từ hai ô trở lên, hoặc cả một ô, hoặc chỉ đặt con trỏ trong ô. Bôi đen một phần chữ trong một ô thì đổi theo chữ.</summary>
        bool CellMode(SelInfo s)
        {
            if (s.Cells.Count != 1) return true;
            try
            {
                dynamic sel = app.Selection; dynamic rng = sel.Range;
                if ((int)rng.Start == (int)rng.End) return true;
                string selText = StripEnd((string)rng.Text);
                string cell = CsvUtil.CellText((string)s.Table.Cell(s.Cells[0][0], s.Cells[0][1]).Range.Text);
                return selText.Trim() == cell.Trim() || selText.Trim().Length == 0;
            }
            catch { return true; }
        }

        string ApplyCells(WordTables w, SelInfo s, NumEditOpts o, string mode)
        {
            dynamic tbl = s.Table;
            var texts = new List<string>();
            foreach (int[] rc in s.Cells) { string tx = ""; try { tx = CsvUtil.CellText((string)tbl.Cell(rc[0], rc[1]).Range.Text); } catch { } texts.Add(tx); }
            // dấu thập phân của cả vùng chọn
            var all = new List<string>();
            foreach (string tx in texts) foreach (Match m in NumberEdit.Find(tx, true)) all.Add(m.Value);
            char evidence = NumberUtil.DetectHint(all);
            int total = 0, changed = 0, same = 0, amb = 0;
            var todo = new List<int>(); var outText = new List<string>();
            for (int i = 0; i < texts.Count; i++)
            {
                if (texts[i].Trim().Length == 0) continue;
                var res = NumberEdit.Plan(texts[i], o, NumberEdit.IsPure(texts[i]));
                total += res.Total; same += res.Same; amb += res.Ambiguous;
                if (res.Changes.Count == 0) continue;
                todo.Add(i); outText.Add(res.Text); changed += res.Changes.Count;
            }
            if (total == 0) return "Không có số nào trong các ô đang chọn.";
            int origS = -1, origE = -1;
            try { dynamic r0 = app.Selection.Range; origS = (int)r0.Start; origE = (int)r0.End; } catch { }
            WordTables.Fast((object)app, "Định dạng số trong bảng", delegate
            {
                for (int k = 0; k < todo.Count; k++)
                {
                    int[] rc = s.Cells[todo[k]];
                    WordTables.SetCellText(tbl, rc[0], rc[1], outText[k]);
                }
            });
            RestoreCells(tbl, s, origS, origE);
            return Summary(o, mode, changed, same, total, amb, evidence, "ô") + " Số trong " + s.Cells.Count + " ô.";
        }

        /// <summary>Đặt lại vùng bôi đen sau khi ghi chữ vào ô (Word có thể bỏ vùng chọn khi nội dung ô bị thay), để bấm tiếp nhiều định dạng liên tiếp.
        /// Vùng chọn hình chữ nhật: chọn lại từ ô đầu đến ô cuối. Con trỏ nằm trong một ô: đặt lại con trỏ. Vùng rời nhau (giữ Ctrl) không dựng lại được.</summary>
        void RestoreCells(dynamic tbl, SelInfo s, int origS, int origE)
        {
            try
            {
                dynamic sel = app.Selection;
                if (s.Cells.Count == 1 && origS >= 0 && origS == origE)
                {
                    dynamic c = tbl.Cell(s.Cells[0][0], s.Cells[0][1]).Range;
                    int cs = (int)c.Start, ce = (int)c.End - 1;
                    sel.SetRange(Math.Max(cs, Math.Min(origS, ce)), Math.Max(cs, Math.Min(origS, ce)));
                    return;
                }
                int r1 = int.MaxValue, r2 = 0, c1 = int.MaxValue, c2 = 0;
                foreach (int[] rc in s.Cells) { r1 = Math.Min(r1, rc[0]); r2 = Math.Max(r2, rc[0]); c1 = Math.Min(c1, rc[1]); c2 = Math.Max(c2, rc[1]); }
                if (s.Cells.Count != (r2 - r1 + 1) * (c2 - c1 + 1)) { Log.W("RestoreCells: vung chon roi nhau, khong dung lai duoc"); return; }
                dynamic a = tbl.Cell(r1, c1).Range, b = tbl.Cell(r2, c2).Range;
                sel.SetRange((int)a.Start, (int)b.End);
            }
            catch (Exception ex) { Log.W("RestoreCells loi: " + ex.Message); }
        }

        static string Summary(NumEditOpts o, string mode, int changed, int same, int total, int amb, char evidence, string unit)
        {
            string what = mode == "sync" && o.Sci ? "Đã viết " + changed + " số theo kiểu kỹ thuật (a × 10ⁿ, a có " + o.SciDec + " chữ số thập phân)."
                : mode == "sync" ? "Đã đồng bộ " + changed + " số về đúng " + o.SetDec + " chữ số thập phân."
                : mode == "cycle" ? "Đã đổi " + changed + " số sang " + NumberEdit.StyleName(o.Style) + "."
                : (mode == "inc" ? "Đã tăng một chữ số thập phân cho " + changed + " số." : "Đã giảm một chữ số thập phân cho " + changed + " số.");
            if (changed == 0)
                what = mode == "sync" && o.Sci ? "Mọi số đã đúng kiểu kỹ thuật với " + o.SciDec + " chữ số thập phân."
                    : mode == "sync" ? "Mọi số đã có đúng " + o.SetDec + " chữ số thập phân."
                    : mode == "cycle" ? "Các số đã đúng " + NumberEdit.StyleName(o.Style) + "."
                    : (mode == "inc" ? "Không số nào tăng thêm được (tối đa " + o.DecMax + " chữ số thập phân)." : "Không số nào giảm thêm được (đã là số nguyên).");
            if (amb > 0) what += " Có " + amb + " số dạng 1,234 hoặc 1.234 mơ hồ: đã hiểu theo dấu thập phân '" + (evidence != '\0' ? evidence : NumberUtil.HintForStyle(o.FallbackStyle)) + "'.";
            return what;
        }

        /// <summary>Đọc số: viết số bằng chữ ngay sau số, trong ngoặc: 1.234 (một nghìn hai trăm ba mươi bốn). Con trỏ trong hoặc sát một số thì đọc số đó,
        /// bôi đen thì đọc mọi số trong vùng. Số đã có sẵn đúng phần đọc phía sau thì bỏ qua (bấm hai lần không thêm hai lần).</summary>
        public string Read(AppSettings st)
        {
            try { if ((int)app.Documents.Count == 0) throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            catch (UserMsg) { throw; }
            catch { throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            dynamic doc = app.ActiveDocument;
            dynamic sel = app.Selection; dynamic rng = sel.Range;
            int start = (int)rng.Start, end = (int)rng.End;
            string text; int baseAt; int fs = -1, fe = -1;
            if (start == end)
            {
                dynamic pr = null;
                foreach (dynamic p in (IEnumerable)rng.Paragraphs) { pr = p.Range; break; }
                if (pr == null) throw new UserMsg("Hãy đặt con trỏ trong hoặc sát một số.");
                text = (string)pr.Text; baseAt = (int)pr.Start;
                fs = fe = start - baseAt;
                if (fs < 0 || fs > text.Length) throw new UserMsg("Hãy đặt con trỏ trong hoặc sát một số.");
            }
            else { text = (string)rng.Text; baseAt = start; }
            bool allowSpace = fs >= 0 || NumberEdit.IsPure(text);
            bool law = st.UnitNames != "common";
            List<NumItem> items = NumberEdit.Items(text, allowSpace, fs, fe);
            if (items.Count == 0 && fs > 0)
            {
                // con trỏ nằm trong đơn vị đứng sau số ("20 kP|a") hoặc ngay sau phần đọc đã viết ("20 kPa (hai mươi kilôpascan)|"): coi như đang ở số phía trước
                foreach (NumItem c in NumberEdit.Items(text, allowSpace, -1, -1))
                {
                    UnitSpan us = UnitWords.Scan(text, c.E, law);
                    int ue = us != null ? us.End : c.E;
                    bool inUnit = us != null && fs > c.E && fs <= ue;
                    bool afterWords = fs > ue && text.Length > ue + 2 && text[ue] == ' ' && text[ue + 1] == '(' && text[fs - 1] == ')' && text.IndexOf(')', ue) == fs - 1;
                    if (inUnit || afterWords) { items = new List<NumItem>(); items.Add(c); break; }
                }
            }
            if (items.Count == 0) throw new UserMsg(fs >= 0 ? "Con trỏ không nằm trong hoặc sát số nào. Đặt con trỏ trong số hoặc ngay sau số rồi bấm lại." : "Không thấy số nào trong phần đang chọn.");
            var raws = new List<string>(); foreach (NumItem it in items) raws.Add(it.Raw);
            char hint = NumberUtil.DetectHint(raws);
            if (hint == '\0') hint = NumberUtil.HintForStyle(st.NumStyle);
            var todo = new List<NumItem>(); var words = new List<string>(); int same = 0, bad = 0; string firstWords = null;
            foreach (NumItem it in items)
            {
                int uend;
                string w = NumberWords.ReadItem(text, it, hint, law, out uend);
                if (w == null) { bad++; continue; }
                it.UEnd = uend;
                if (firstWords == null) firstWords = w;
                string after = uend < text.Length ? text.Substring(uend) : "";
                if (after.StartsWith(" (" + w + ")", StringComparison.Ordinal)) { same++; continue; }
                todo.Add(it); words.Add(w);
            }
            if (todo.Count == 0)
            {
                if (same > 0) return "Số đã có phần đọc bằng chữ ngay sau.";
                throw new UserMsg("Không đọc được số này (quá " + NumberWords.MaxDigits + " chữ số hoặc không đúng dạng số).");
            }
            int failed = 0, delta = 0, lastEnd = -1;
            WordTables.Fast((object)app, "Đọc số", delegate
            {
                for (int i = todo.Count - 1; i >= 0; i--)       // từ cuối lên đầu để vị trí phía trước không lệch
                {
                    NumItem c = todo[i];
                    string whole = text.Substring(c.S, c.UEnd - c.S);          // gồm cả đơn vị: bản đọc chèn ngay sau đơn vị
                    int bells = 0; for (int k = 0; k < c.S && k < text.Length; k++) if (text[k] == '\a') bells++;
                    int at = baseAt + c.S - bells;
                    dynamic r = doc.Range(at, at + whole.Length);
                    string actual = ""; try { actual = (string)r.Text; } catch { }
                    if (actual != whole) { failed++; continue; }
                    string ins = " (" + words[i] + ")";
                    r.InsertAfter(ins);
                    delta += ins.Length;
                    if (lastEnd < 0) lastEnd = at + whole.Length + ins.Length;
                }
            });
            try
            {
                if (fs >= 0 && lastEnd >= 0) sel.SetRange(lastEnd, lastEnd);              // con trỏ nằm ngay sau phần vừa viết, gõ tiếp được
                else sel.SetRange(start, Math.Max(start, end + delta));
            }
            catch (Exception ex) { Log.W("Doc so: dat lai vung chon loi: " + ex.Message); }
            string m = todo.Count - failed == 1 ? "Đã viết bằng chữ: " + firstWordsOf(todo, words, failed) : "Đã viết bằng chữ cho " + (todo.Count - failed) + " số.";
            if (same > 0) m += " " + same + " số đã có sẵn.";
            if (bad > 0) m += " " + bad + " số không đọc được.";
            if (failed > 0) m += " Không định vị được " + failed + " số (có thể nằm trong trường hoặc công thức).";
            return m;
        }
        static string firstWordsOf(List<NumItem> todo, List<string> words, int failed) { return words.Count > 0 ? words[0] : ""; }

        string ApplyInline(NumEditOpts o, string mode)
        {
            dynamic doc = app.ActiveDocument;
            dynamic sel = app.Selection; dynamic rng = sel.Range;
            int start = (int)rng.Start, end = (int)rng.End;
            string text; int baseAt; int fs = -1, fe = -1;
            if (start == end)
            {
                // chỉ đặt con trỏ: dùng số đang chứa con trỏ trong đoạn
                dynamic pr = null;
                foreach (dynamic p in (IEnumerable)rng.Paragraphs) { pr = p.Range; break; }
                if (pr == null) throw new UserMsg("Hãy bôi đen các số cần định dạng.");
                text = (string)pr.Text; baseAt = (int)pr.Start;
                fs = fe = start - baseAt;
                if (fs < 0 || fs > text.Length) throw new UserMsg("Hãy bôi đen các số cần định dạng.");
            }
            else { text = (string)rng.Text; baseAt = start; }
            bool pure = fs < 0 && NumberEdit.IsPure(text);
            if (fs >= 0) pure = false;
            NumEditResult res = NumberEdit.Plan(text, o, pure || fs >= 0, fs, fe);
            if (res.Total == 0) throw new UserMsg("Không thấy số nào trong phần đang chọn. Hãy bôi đen một hoặc nhiều số, hoặc đặt con trỏ trong một số.");
            if (res.Changes.Count == 0) return Summary(o, mode, 0, res.Same, res.Total, res.Ambiguous, res.Hint, "số");
            int failed = 0, delta = 0;
            WordTables.Fast((object)app, "Định dạng số", delegate
            {
                for (int i = res.Changes.Count - 1; i >= 0; i--)        // từ cuối lên đầu để vị trí phía trước không lệch
                {
                    NumChange c = res.Changes[i];
                    int bells = 0; for (int k = 0; k < c.S && k < text.Length; k++) if (text[k] == '\a') bells++;      // dấu hết ô của bảng chiếm 2 ký tự trong chữ nhưng 1 vị trí
                    int at = baseAt + c.S - bells;
                    dynamic r = doc.Range(at, at + c.Old.Length);
                    string actual = ""; try { actual = (string)r.Text; } catch { }
                    if (actual != c.Old) { failed++; continue; }
                    r.Text = c.New;
                    delta += c.New.Length - c.Old.Length;
                }
            });
            // Giữ vùng bôi đen để bấm tiếp: cùng điểm đầu, điểm cuối dịch theo tổng độ dài chữ thay đổi (tính bằng số học, không phụ thuộc cách Word co giãn vùng).
            // Rev0.8: chỉ đặt con trỏ (không bôi đen) thì sau khi sửa vẫn là con trỏ, đặt ngay cuối số vừa sửa. Trước đây vùng chọn bị dựng thành [đầu, đầu + số ký tự thêm vào],
            // tức là bôi đen đúng một ký tự giữa số, lần bấm sau sửa trúng ký tự đó và số bị phá ("48,5" thành "48,5,0" rồi "48,5,000").
            try
            {
                if (fs >= 0)
                {
                    int caret = start;
                    if (res.Changes.Count > 0)
                    {
                        NumChange c0 = res.Changes[0];
                        int b0 = 0; for (int k = 0; k < c0.S && k < text.Length; k++) if (text[k] == '\a') b0++;
                        caret = baseAt + c0.S - b0 + c0.New.Length;
                    }
                    sel.SetRange(caret, caret);
                }
                else sel.SetRange(start, Math.Max(start, end + delta));
            }
            catch (Exception ex) { Log.W("Khoi phuc vung chon loi: " + ex.Message); }
            string m = Summary(o, mode, res.Changes.Count - failed, res.Same, res.Total, res.Ambiguous, res.Hint, "số");
            if (failed > 0) m += " Không định vị được " + failed + " số (có thể nằm trong trường hoặc công thức).";
            return m;
        }
    }

    /// <summary>Kiểu chỉ số: tự tìm công thức toán, lý, hóa trong phần bôi đen và luân chuyển cách viết chỉ số:
    /// Luân phiên hai kiểu: Equation ⇄ font thường có chỉ số trên, dưới bằng định dạng. (Kiểu ký tự nhỏ thấp của font Cambria Math chỉ còn dùng khi chọn trong Cài đặt.)</summary>
    public sealed class WordIndexStyle
    {
        readonly dynamic app;
        public WordIndexStyle(object application) { app = application; }

        const string SubG = "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎";
        const string SupG = "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ";
        static readonly Regex GlyphRe = new Regex("[₀-₉₊₋₌₍₎⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ]");

        public static string ModeName(string m)
        {
            return m == "fmt" ? "font thường, chỉ số trên và dưới bằng định dạng" : (m == "glyph" ? "ký tự chỉ số nhỏ, thấp (font Cambria Math)" : "Equation cùng dòng");
        }

        /// <summary>Đoán kiểu hiện tại của phần bôi đen: raw (chữ phẳng), fmt, glyph, eq.</summary>
        public static string NextMode(string current) { return current == "fmt" ? "eq" : "fmt"; }     // Equation → font thường → Equation...; chữ phẳng hoặc ký tự nhỏ thấp thì sang font thường trước

        /// <summary>Đổi về dạng gõ phẳng để bộ nhận diện đọc được: ký tự chỉ số Unicode (H₂O thành H2O, SO₄²⁻ thành SO4^2-, x² thành x^2), đoạn đang định dạng chỉ số trên, dưới (thành _ và ^),
        /// dấu nháy của UnicodeMath. marks: các đoạn {bắt đầu, kết thúc, loại} theo vị trí trong chữ, loại 1 = chỉ số dưới, 2 = chỉ số trên (có thể null). map[i] là vị trí trong chữ gốc của ký tự thứ i.</summary>
        public static string Normalize(string text, List<int[]> marks, out int[] map)
        {
            var kindAt = new int[text.Length + 1];
            var runEnd = new int[text.Length + 1];
            if (marks != null)
                foreach (int[] m in marks)
                    for (int i = Math.Max(0, m[0]); i < Math.Min(text.Length, m[1]); i++) { kindAt[i] = m[2]; runEnd[i] = m[1]; }
            var sb = new StringBuilder(); var mp = new List<int>();
            var closeAt = new Stack<bool>();
            int fmtEnd = -1;                                              // vị trí hết đoạn đang định dạng, để đóng ngoặc
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (fmtEnd == i) { sb.Append('}'); mp.Add(i - 1 >= 0 ? i - 1 : 0); fmtEnd = -1; }
                if (kindAt[i] != 0 && (i == 0 || kindAt[i - 1] != kindAt[i]))
                {
                    int len = runEnd[i] - i;
                    bool brace = len > 1;
                    sb.Append(kindAt[i] == 1 ? '_' : '^'); mp.Add(i);
                    if (brace) { sb.Append('{'); mp.Add(i); fmtEnd = runEnd[i]; }
                    else fmtEnd = -1;
                }
                int si = SubG.IndexOf(c), pi = SupG.IndexOf(c);
                if (c == '"') continue;                                       // dấu nháy của chuỗi UnicodeMath
                if (si >= 0)
                {
                    char a = "0123456789+-=()"[si];
                    bool prevLower = sb.Length > 0 && sb[sb.Length - 1] >= 'a' && sb[sb.Length - 1] <= 'z';
                    bool startRun = i == 0 || SubG.IndexOf(text[i - 1]) < 0;
                    if (startRun && prevLower) { sb.Append('_'); mp.Add(i); }
                    sb.Append(a); mp.Add(i); continue;
                }
                if (pi >= 0)
                {
                    bool startRun = i == 0 || SupG.IndexOf(text[i - 1]) < 0;
                    if (startRun) { sb.Append('^'); mp.Add(i); }
                    sb.Append("0123456789+-=()ni"[pi]); mp.Add(i); continue;
                }
                if ((c == '^' || c == '_') && i + 1 < text.Length && text[i + 1] == '(')
                {
                    sb.Append(c).Append('{'); mp.Add(i); mp.Add(i); i++; closeAt.Push(true); continue;
                }
                if (c == '(') { closeAt.Push(false); sb.Append(c); mp.Add(i); continue; }
                if (c == ')' && closeAt.Count > 0)
                {
                    bool br = closeAt.Pop(); sb.Append(br ? '}' : ')'); mp.Add(i); continue;
                }
                sb.Append(c); mp.Add(i);
            }
            if (fmtEnd == text.Length) { sb.Append('}'); mp.Add(text.Length - 1); }
            mp.Add(text.Length);
            map = mp.ToArray();
            return sb.ToString();
        }
        public static string Normalize(string text, out int[] map) { return Normalize(text, null, out map); }

        /// <summary>Tìm các đoạn đang định dạng chỉ số trên, dưới trong [a, b) của tài liệu bằng cách chia đôi dần (ít lệnh gọi Word).</summary>
        static List<int[]> FormatRuns(dynamic doc, int a, int b)
        {
            var res = new List<int[]>();
            ScanFmt(doc, a, b, res, 0);
            var merged = new List<int[]>();
            foreach (int[] r in res.OrderBy(x => x[0]))
            {
                if (merged.Count > 0 && merged[merged.Count - 1][1] == r[0] && merged[merged.Count - 1][2] == r[2]) merged[merged.Count - 1][1] = r[1];
                else merged.Add(new int[] { r[0], r[1], r[2] });
            }
            return merged;
        }
        static void ScanFmt(dynamic doc, int a, int b, List<int[]> res, int depth)
        {
            if (b <= a || depth > 24) return;
            dynamic r = doc.Range(a, b);
            int sub = 0, sup = 0;
            try { sub = (int)r.Font.Subscript; sup = (int)r.Font.Superscript; } catch { return; }
            if (sub == 0 && sup == 0) return;
            if (sub == -1 && sup == 0) { res.Add(new int[] { a, b, 1 }); return; }
            if (sup == -1 && sub == 0) { res.Add(new int[] { a, b, 2 }); return; }
            if (b - a == 1) return;
            int m = (a + b) / 2;
            ScanFmt(doc, a, m, res, depth + 1); ScanFmt(doc, m, b, res, depth + 1);
        }

        /// <summary>Công thức cần xử lý trong chữ đã chuẩn hóa: chỉ mục chưa rõ sai (Sev fmt) và ký hiệu mã LaTeX; hóa học sai hóa trị thì để mục Chuẩn hóa xử lý.</summary>
        public static List<Hit> FindFormulas(string norm)
        {
            var o = new ReviewOpts(); o.Num = false;
            var hits = Detector.Detect(norm, o, null);
            return hits.Where(h => h.Issue.Fixes.Count > 0 && (h.Issue.Sev == "fmt" || h.Issue.Cat == "sym")).ToList();
        }

        string Current(dynamic rng, string text)
        {
            try { if ((int)rng.OMaths.Count > 0) return "eq"; } catch { }
            if (GlyphRe.IsMatch(text)) return "glyph";
            try { if ((int)rng.Font.Subscript != 0 || (int)rng.Font.Superscript != 0) return "fmt"; } catch { }
            return "raw";
        }

        /// <summary>mode = null: luân chuyển từ kiểu hiện tại. Hoặc "fmt", "glyph", "eq" để đặt thẳng.</summary>
        public string Cycle(string forceMode, bool stacked = false)
        {
            try { if ((int)app.Documents.Count == 0) throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            catch (UserMsg) { throw; }
            catch { throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            dynamic doc = app.ActiveDocument;
            dynamic sel = app.Selection; dynamic rng = sel.Range;
            int start = (int)rng.Start, end = (int)rng.End;
            if (start == end) throw new UserMsg("Hãy bôi đen công thức hoặc đoạn có công thức (ví dụ H2O, SO4^2-, x^2 + y^2 = r^2) rồi bấm lại.");
            dynamic live = null;
            try { live = doc.Range(start, end); } catch { live = null; }
            string result = "";
            WordTables.Fast((object)app, "Kiểu chỉ số", delegate
            {
                string text = (string)rng.Text;
                string cur = Current(rng, text);
                string target = forceMode ?? NextMode(cur);
                if (cur == "eq" && target != "eq")
                {
                    // Equation có sẵn → font thường: đọc thẳng OMML của từng Equation rồi thay bằng chữ có chỉ số trên, dưới
                    int okE = 0, failE = 0;
                    ConvertEquationsToText(doc, rng, ref okE, ref failE);
                    result = okE > 0 ? "Đã chuyển " + okE + " Equation sang " + ModeName("fmt") + "." + (failE > 0 ? " Không chuyển được " + failE + " chỗ." : "") + " Bấm lại để chuyển ngược thành Equation."
                        : "Không chuyển được Equation về chữ thường" + (failE > 0 ? " (" + failE + " chỗ lỗi, xem nhật ký)." : ".");
                    return;
                }
                List<int[]> marks = null;
                if (text.IndexOf('\a') < 0 && cur != "eq")
                {
                    try
                    {
                        marks = new List<int[]>();
                        foreach (int[] fr in FormatRuns(doc, start, start + text.Length)) marks.Add(new int[] { fr[0] - start, fr[1] - start, fr[2] });
                    }
                    catch { marks = null; }
                }
                int[] map; string norm = Normalize(text, marks, out map);
                List<Hit> hits = FindFormulas(norm);
                if (hits.Count == 0) { result = "Không tìm thấy công thức nào trong phần đang bôi đen. Thử bôi đen đúng công thức, ví dụ H2O hoặc x^2."; return; }
                int ok = 0, fail = 0;
                foreach (Hit h in hits.OrderByDescending(x => x.S))
                {
                    int s0 = map[h.S], e0 = map[h.E - 1] + 1;
                    while (s0 > 0 && text[s0 - 1] == '"') s0--;           // dấu nháy bao quanh chữ cái của UnicodeMath
                    while (e0 < text.Length && text[e0] == '"') e0++;
                    string orig = text.Substring(s0, e0 - s0);
                    dynamic r;
                    try { r = doc.Range(start + s0 - CountBells(text, s0), start + e0 - CountBells(text, e0)); } catch { fail++; continue; }
                    string actual = ""; try { actual = (string)r.Text; } catch { }
                    if (actual != orig) { fail++; continue; }
                    Fix f = h.Issue.Fixes[0];
                    try
                    {
                        if (target == "eq" && f.NeedsFormat)
                        {
                            // dựng thẳng bằng OMML: không còn dấu nháy, ngoặc thừa như khi nhờ Word BuildUp chuỗi tuyến tính
                            if (!WordEq.Insert(doc, r, f.GetNode(), false, stacked)) WordReview.ApplyRunsTo(doc, r, f, false);
                        }
                        else if (f.Runs != null && f.NeedsFormat) WordReview.ApplyRunsTo(doc, r, f, target == "glyph");
                        else { r.Text = f.Plain; }
                        ok++;
                    }
                    catch { fail++; }
                }
                result = "Đã chuyển " + ok + " công thức sang " + ModeName(target) + "." + (fail > 0 ? " Không chuyển được " + fail + " chỗ." : "")
                    + (ok > 0 ? " Bấm lại để sang kiểu kế tiếp." : "");
            });
            try { if (live != null) sel.SetRange((int)live.Start, (int)live.End); } catch { }
            return result;
        }

        /// <summary>Thay từng Equation (OMath) trong vùng bằng chữ thường có định dạng chỉ số trên, dưới. Đọc OMML qua WordOpenXML, xóa Equation, chèn chữ, kiểm tra lại.</summary>
        internal static void ConvertEquationsToText(dynamic doc, dynamic rng, ref int ok, ref int fail)
        {
            int n = 0;
            try { n = (int)rng.OMaths.Count; } catch { return; }
            for (int i = n; i >= 1; i--)
            {
                try
                {
                    dynamic om = rng.OMaths.Item(i);
                    dynamic orng = om.Range;
                    int os = (int)orng.Start;
                    string xml = "";
                    try { xml = (string)orng.WordOpenXML; } catch (Exception e1) { Log.W("WordOpenXML: " + e1.Message); }
                    MNode node = Omml.FromOmml(xml);
                    if (node == null) { Log.W("Equation " + i + ": khong doc duoc OMML"); fail++; continue; }
                    Fix f = Fix.FromNode("Chữ thường", node);
                    if (f.Plain.Length == 0) { fail++; continue; }
                    string prevFont = ""; double prevSize = 0;
                    try
                    {
                        if (os > 0) { dynamic pc = doc.Range(os - 1, os); prevFont = (string)pc.Font.Name; prevSize = (double)pc.Font.Size; }
                    }
                    catch { }
                    try { orng.Delete(); } catch (Exception e2) { Log.W("Xoa Equation loi: " + e2.Message); fail++; continue; }
                    try
                    {
                        dynamic chk = doc.Range(os, os);
                        if ((int)chk.OMaths.Count > 0) { try { chk.OMaths.Item(1).Remove(); } catch { } }      // còn sót vùng Equation rỗng
                    }
                    catch { }
                    dynamic ins = doc.Range(os, os);
                    WordReview.ApplyRunsTo(doc, ins, f, false);
                    string got = ""; try { got = (string)ins.Text; } catch { }
                    int inMath = 0; try { inMath = (int)ins.OMaths.Count; } catch { }
                    if (inMath > 0 || got != f.Plain) { Log.W("Equation " + i + ": chen chu khong dung (" + got + ", omath=" + inMath + ")"); fail++; continue; }
                    try
                    {
                        string fn = (string)ins.Font.Name;
                        if ((fn == "Cambria Math" || string.IsNullOrEmpty(fn)) && prevFont.Length > 0 && prevFont != "Cambria Math") { ins.Font.Name = prevFont; if (prevSize > 0) ins.Font.Size = (float)prevSize; }
                    }
                    catch { }
                    ok++;
                }
                catch (Exception ex) { Log.W("ConvertEquationsToText loi: " + ex.Message); fail++; }
            }
        }

        static int CountBells(string t, int upTo) { int n = 0; for (int i = 0; i < upTo && i < t.Length; i++) if (t[i] == '\a') n++; return n; }
    }
}
