// Trợ lý Kỹ thuật - Rev0.2: thao tác bảng trong Word bằng "late binding" (dynamic), không cần Office Interop.
// Điểm mới so với Rev0.1.3:
//  - Vùng chọn được chụp thành danh sách tọa độ (hàng, cột) trước khi sửa, vì tập ô của Word thay đổi khi ghi chữ.
//  - Mọi lệnh ghi theo kiểu "đọc, so sánh, chỉ ghi chỗ khác" nên chạy nhanh và không áp lại định dạng đã đúng.
//  - Tắt cập nhật màn hình, phân trang nền và kiểm tra chính tả khi chạy lệnh.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace TroLyKyThuat
{
    public sealed class UserMsg : Exception
    {
        public UserMsg(string m) : base(m) { }
    }

    public sealed class SelInfo
    {
        public bool InTable;
        public dynamic Table;
        public int R1, R2, C1, C2;
        public int RowCount, ColCount;
        public int TableNo, TableTotal;
        public int NumericCells, SelCells;
        public List<int> Rows = new List<int>();      // các hàng có ô được chọn (không trùng, tăng dần)
        public List<int> Cols = new List<int>();      // các cột có ô được chọn
        public List<int[]> Cells = new List<int[]>(); // {hàng, cột} của từng ô được chọn (bản chụp, không đổi)
        public string Diag = "";
    }

    /// <summary>Thông tin vùng chọn gọn nhẹ, an toàn để đưa sang luồng giao diện.</summary>
    public sealed class SelDto
    {
        public bool InTable;
        public string Text = "";
    }


    /// <summary>
    /// Word chỉ cho Selection.Cells của vùng chọn cuối cùng khi bạn giữ Ctrl để chọn nhiều vùng rời nhau.
    /// Bộ theo dõi này quan sát vùng chọn mỗi vài chục mili giây; mỗi lần vùng chọn đổi khi Ctrl đang (hoặc vừa) được giữ,
    /// vùng mới được cộng dồn vào các vùng trước, nhờ đó các lệnh biết đủ mọi ô đã chọn.
    /// </summary>
    public sealed class SelTracker
    {
        public string Key, LastSig;
        readonly HashSet<long> cells = new HashSet<long>();
        long lastCtrl = -100000;
        public const int CtrlGraceMs = 450;

        public int Count { get { return cells.Count; } }

        /// <summary>Ghi nhận một lần quan sát. tableKey = null nghĩa là con trỏ không nằm trong bảng. Trả về tập ô đầy đủ.</summary>
        public List<int[]> Observe(string sig, string tableKey, List<int[]> current, bool ctrlNow, long nowMs)
        {
            if (ctrlNow) lastCtrl = nowMs;
            if (sig != LastSig || tableKey != Key)
            {
                bool ctrl = ctrlNow || (nowMs - lastCtrl) < CtrlGraceMs;
                bool add = ctrl && LastSig != null && tableKey != null && tableKey == Key && cells.Count > 0;
                if (!add) cells.Clear();
                Key = tableKey; LastSig = sig;
                if (tableKey != null) foreach (int[] rc in current) cells.Add((long)rc[0] * 100000 + rc[1]);
            }
            var res = new List<int[]>();
            if (tableKey == null) return res;
            foreach (long k in cells) res.Add(new int[] { (int)(k / 100000), (int)(k % 100000) });
            res.Sort((a, b) => a[0] != b[0] ? a[0].CompareTo(b[0]) : a[1].CompareTo(b[1]));
            return res;
        }

        public void Reset() { cells.Clear(); Key = null; LastSig = null; }
    }

    public static class Keyboard
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern short GetAsyncKeyState(int vKey);
        public static Func<bool> CtrlDown = delegate { try { return (GetAsyncKeyState(0x11) & 0x8000) != 0; } catch { return false; } };
        public static Func<long> NowMs = delegate { return Environment.TickCount; };
    }

    public sealed class WordTables
    {
        public static readonly SelTracker Tracker = new SelTracker();
        public const string StyleName = "TroLy_Master";
        readonly dynamic app;
        int nSet, nSame;          // đếm số thuộc tính đã ghi, số thuộc tính đã đúng sẵn
        string only = null;
        bool Want(string k) { return only == null || only == k; }

        public WordTables(object application) { app = application; }

        // ------------------------------------------------------------------ tiện ích chung
        static string CellStr(dynamic cell)
        {
            return CsvUtil.CellText((string)cell.Range.Text);
        }

        /// <summary>Chạy một việc ở chế độ nhanh: một bước Undo, tắt cập nhật màn hình, phân trang nền, kiểm tra chính tả khi gõ.</summary>
        public static void Fast(dynamic app, string name, Action a)
        {
            dynamic ur = null;
            try { ur = app.UndoRecord; ur.StartCustomRecord(name); } catch { ur = null; }
            bool su = true, pg = true, sp = true, gr = true;
            bool suOk = false, pgOk = false, spOk = false, grOk = false;
            try { su = (bool)app.ScreenUpdating; app.ScreenUpdating = false; suOk = true; } catch { }
            try { pg = (bool)app.Options.Pagination; app.Options.Pagination = false; pgOk = true; } catch { }
            try { sp = (bool)app.Options.CheckSpellingAsYouType; app.Options.CheckSpellingAsYouType = false; spOk = true; } catch { }
            try { gr = (bool)app.Options.CheckGrammarAsYouType; app.Options.CheckGrammarAsYouType = false; grOk = true; } catch { }
            try { a(); }
            finally
            {
                if (grOk) { try { app.Options.CheckGrammarAsYouType = gr; } catch { } }
                if (spOk) { try { app.Options.CheckSpellingAsYouType = sp; } catch { } }
                if (pgOk) { try { app.Options.Pagination = pg; } catch { } }
                if (suOk) { try { app.ScreenUpdating = su; } catch { } }
                if (ur != null) { try { ur.EndCustomRecord(); } catch { } }
            }
        }

        void Undo(string name, Action a) { nSet = 0; nSame = 0; Fast((object)app, name, a); }

        // --- so sánh rồi mới ghi ---
        void SetI(Func<int> get, Action<int> set, int want)
        {
            int cur; bool known = true;
            try { cur = get(); } catch { cur = 0; known = false; }
            if (known && cur == want) { nSame++; return; }
            set(want); nSet++;
        }
        void SetF(Func<double> get, Action<double> set, double want)
        {
            double cur; bool known = true;
            try { cur = get(); } catch { cur = 0; known = false; }
            if (known && Math.Abs(cur - want) < 0.05) { nSame++; return; }
            set(want); nSet++;
        }
        void SetB(Func<bool> get, Action<bool> set, bool want)
        {
            bool cur; bool known = true;
            try { cur = get(); } catch { cur = false; known = false; }
            if (known && cur == want) { nSame++; return; }
            set(want); nSet++;
        }
        void SetS(Func<string> get, Action<string> set, string want)
        {
            string cur; bool known = true;
            try { cur = get(); } catch { cur = null; known = false; }
            if (known && cur == want) { nSame++; return; }
            set(want); nSet++;
        }

        public bool HasDocument()
        {
            try { return (int)app.Documents.Count > 0; } catch { return false; }
        }

        // ------------------------------------------------------------------ vùng chọn
        /// <summary>
        /// Chụp vùng chọn thành danh sách tọa độ. Word cho ô theo "bộ sưu tập sống", nên chụp ngay một lần,
        /// sau đó chỉ làm việc bằng tọa độ và tbl.Cell(hàng, cột).
        /// </summary>
        public SelInfo GetSelection(bool withText)
        {
            var info = new SelInfo();
            if (!HasDocument()) return info;
            dynamic sel = app.Selection;
            bool inT;
            try { inT = (bool)sel.Information(12); } catch { return info; }   // wdWithInTable
            string sig = null; int st = 0, en = 0, ty = -1;
            try { st = (int)sel.Range.Start; en = (int)sel.Range.End; ty = (int)sel.Type; sig = st + "-" + en + "-" + ty; } catch { }
            if (!inT) { Tracker.Observe(sig, null, null, Keyboard.CtrlDown(), Keyboard.NowMs()); return info; }
            dynamic tbl = sel.Tables.Item(1);
            info.InTable = true;
            info.Table = tbl;
            string tkey = null;
            try { tkey = ((int)tbl.Range.Start).ToString(CultureInfo.InvariantCulture); } catch { tkey = "?"; }
            var raw = new List<int[]>();
            var seen = new Dictionary<long, bool>();
            int nCells = -1, nRows = -1, nCols = -1;
            try
            {
                dynamic cells = sel.Cells; int n = (int)cells.Count; nCells = n;
                for (int i = 1; i <= n; i++)
                {
                    dynamic cell = cells.Item(i);
                    int r = (int)cell.RowIndex, c = (int)cell.ColumnIndex;
                    long key = (long)r * 100000 + c;
                    if (seen.ContainsKey(key)) continue;
                    seen[key] = true; raw.Add(new int[] { r, c });
                }
            }
            catch { }
            int tracked = Tracker.Count;
            info.Cells = Tracker.Observe(sig, tkey, raw, Keyboard.CtrlDown(), Keyboard.NowMs());
            var rowSet = new SortedDictionary<int, bool>(); var colSet = new SortedDictionary<int, bool>();
            foreach (int[] rc in info.Cells) { rowSet[rc[0]] = true; colSet[rc[1]] = true; }
            // Hợp với các hàng, cột mà Word báo (vùng chọn dạng cột hoặc hàng nguyên)
            try
            {
                dynamic rows = sel.Rows; nRows = (int)rows.Count;
                for (int i = 1; i <= nRows && i <= 500; i++) { int r = (int)rows.Item(i).Index; if (r > 0) rowSet[r] = true; }
            }
            catch { }
            try
            {
                dynamic cols = sel.Columns; nCols = (int)cols.Count;
                for (int i = 1; i <= nCols && i <= 200; i++) { int c = (int)cols.Item(i).Index; if (c > 0) colSet[c] = true; }
            }
            catch { }
            info.Rows = new List<int>(rowSet.Keys); info.Cols = new List<int>(colSet.Keys);
            int r1 = int.MaxValue, r2 = 0, c1 = int.MaxValue, c2 = 0;
            foreach (int r in info.Rows) { if (r < r1) r1 = r; if (r > r2) r2 = r; }
            foreach (int c in info.Cols) { if (c < c1) c1 = c; if (c > c2) c2 = c; }
            if (info.Rows.Count == 0) { r1 = 1; r2 = 1; }
            if (info.Cols.Count == 0) { c1 = 1; c2 = 1; }
            info.R1 = r1; info.R2 = r2; info.C1 = c1; info.C2 = c2;
            info.SelCells = info.Cells.Count;
            try { info.RowCount = (int)tbl.Rows.Count; } catch { info.RowCount = r2; }
            try { info.ColCount = (int)tbl.Columns.Count; } catch { info.ColCount = c2; }
            try
            {
                dynamic doc = app.ActiveDocument;
                int start = (int)tbl.Range.Start;
                info.TableTotal = (int)doc.Tables.Count;
                for (int i = 1; i <= info.TableTotal; i++)
                    if ((int)doc.Tables.Item(i).Range.Start == start) { info.TableNo = i; break; }
            }
            catch { }
            if (withText && info.Cells.Count <= 400)
            {
                int nums = 0;
                foreach (int[] rc in info.Cells)
                {
                    try { if (NumberUtil.Parse(CellStr(tbl.Cell(rc[0], rc[1]))) != null) nums++; } catch { }
                }
                info.NumericCells = nums;
            }
            info.Diag = string.Format("SEL type={0} range={1}-{2} cells={3} rowsApi={4} colsApi={5} gomCtrl={6} => hang [{7}] cot [{8}]",
                ty, st, en, nCells, nRows, nCols, tracked, IndexList.Format(info.Rows), IndexList.Format(info.Cols));
            return info;
        }

        /// <summary>Gọi định kỳ từ bộ đếm giờ trên luồng chính của Word: ghi nhận vùng chọn mới (gồm cả vùng chọn rời nhau bằng Ctrl).</summary>
        public void Poll()
        {
            try
            {
                if (!HasDocument()) return;
                dynamic sel = app.Selection;
                string sig = null;
                try { sig = (int)sel.Range.Start + "-" + (int)sel.Range.End + "-" + (int)sel.Type; } catch { return; }
                bool ctrl = Keyboard.CtrlDown();
                if (ctrl) { Tracker.Observe(Tracker.LastSig, Tracker.Key, new List<int[]>(), true, Keyboard.NowMs()); }   // chỉ cập nhật mốc thời gian Ctrl
                if (sig == Tracker.LastSig) return;
                GetSelection(false);
            }
            catch { }
        }

        public string Describe(SelInfo s)
        {
            if (!s.InTable) return "Chưa chọn bảng nào. Hãy đặt con trỏ vào một ô hoặc bôi đen các ô trong bảng của văn bản.";
            string num = s.NumericCells > 0 ? string.Format(", {0} ô có số", s.NumericCells) : "";
            return string.Format("Bảng thứ {0} trên {1}. Hàng đã chọn: {2}. Cột đã chọn: {3}. ({4} ô{5}). Bảng có {6} hàng, {7} cột.",
                s.TableNo, s.TableTotal, IndexList.Format(s.Rows), IndexList.Format(s.Cols), s.SelCells, num, s.RowCount, s.ColCount);
        }

        /// <summary>Gọn nhẹ cho bảng điều khiển: chạy trên luồng Word, trả về chữ.</summary>
        public SelDto Snapshot()
        {
            var d = new SelDto();
            try
            {
                SelInfo s = GetSelection(true);
                d.InTable = s.InTable; d.Text = Describe(s);
                if (s.InTable) Log.W(s.Diag);
            }
            catch (Exception ex) { d.Text = "Chưa đọc được vùng chọn trong Word: " + ex.Message; }
            return d;
        }

        SelInfo Need(bool withText)
        {
            var s = GetSelection(withText);
            if (!s.InTable) throw new UserMsg("Hãy đặt con trỏ vào một ô của bảng trước.");
            return s;
        }

        // ------------------------------------------------------------------ thêm, xóa hàng và cột
        // Mỗi hàng (cột) được chọn đều được xử lý riêng: thêm n hàng (cột) ngay cạnh từng hàng (cột) đó.
        // "manual" là danh sách số do người dùng nhập (ví dụ 2,4,6-8); để trống thì dùng vùng chọn trong Word.
        List<int> Targets(SelInfo s, string manual, bool rows)
        {
            if (manual != null && manual.Trim().Length > 0)
            {
                int max = rows ? s.RowCount : s.ColCount;
                List<int> l = IndexList.Parse(manual, max);
                if (l == null || l.Count == 0)
                    throw new UserMsg("Không hiểu danh sách số \"" + manual.Trim() + "\". Hãy nhập kiểu 2,4,6-8 với các số từ 1 đến " + max + ".");
                return l;
            }
            return rows ? s.Rows : s.Cols;
        }

        void InsertRows(dynamic tbl, int r, int n, bool above)
        {
            int done = 0;
            try
            {
                for (; done < n; done++)
                {
                    int rc = (int)tbl.Rows.Count;
                    if (above) tbl.Rows.Add(tbl.Rows.Item(r));
                    else if (r >= rc) tbl.Rows.Add();
                    else tbl.Rows.Add(tbl.Rows.Item(r + 1));
                }
            }
            catch
            {
                if (done >= n) return;
                dynamic sel = app.Selection;
                tbl.Cell(r, 1).Select();
                if (above) sel.InsertRowsAbove(n - done); else sel.InsertRowsBelow(n - done);
            }
        }

        void InsertCols(dynamic tbl, int c, int n, bool left)
        {
            int done = 0;
            try
            {
                for (; done < n; done++)
                {
                    int cc = (int)tbl.Columns.Count;
                    if (left) tbl.Columns.Add(tbl.Columns.Item(c));
                    else if (c >= cc) tbl.Columns.Add();
                    else tbl.Columns.Add(tbl.Columns.Item(c + 1));
                }
            }
            catch
            {
                if (done >= n) return;
                dynamic sel = app.Selection;
                tbl.Cell(1, c).Select();
                if (left) sel.InsertColumns(n - done); else sel.InsertColumnsRight(n - done);
            }
        }

        public string AddRows(int n, bool above, string manual)
        {
            if (n <= 0) return "Số lượng là 0: không thêm hàng nào.";
            var s = Need(false); dynamic tbl = s.Table;
            List<int> t = Targets(s, manual, true);
            Undo(above ? "Thêm hàng phía trên mỗi hàng đã chọn" : "Thêm hàng phía dưới mỗi hàng đã chọn", delegate
            {
                for (int i = t.Count - 1; i >= 0; i--) InsertRows(tbl, t[i], n, above);   // từ dưới lên để chỉ số không bị lệch
            });
            return string.Format("Đã thêm {0} hàng {1} mỗi hàng trong số: hàng {2}. Tổng cộng {3} hàng mới.",
                n, above ? "phía trên" : "phía dưới", IndexList.Format(t), n * t.Count);
        }

        public string AddCols(int n, bool left, string manual)
        {
            if (n <= 0) return "Số lượng là 0: không thêm cột nào.";
            var s = Need(false); dynamic tbl = s.Table;
            List<int> t = Targets(s, manual, false);
            Undo(left ? "Thêm cột bên trái mỗi cột đã chọn" : "Thêm cột bên phải mỗi cột đã chọn", delegate
            {
                for (int i = t.Count - 1; i >= 0; i--) InsertCols(tbl, t[i], n, left);
            });
            return string.Format("Đã thêm {0} cột {1} mỗi cột trong số: cột {2}. Tổng cộng {3} cột mới.",
                n, left ? "bên trái" : "bên phải", IndexList.Format(t), n * t.Count);
        }

        public string DeleteRows(string manual)
        {
            var s = Need(false); dynamic tbl = s.Table;
            List<int> t = Targets(s, manual, true);
            if (t.Count >= s.RowCount) throw new UserMsg("Không thể xóa toàn bộ các hàng của bảng bằng công cụ này.");
            Undo("Xóa hàng", delegate
            {
                for (int i = t.Count - 1; i >= 0; i--)
                {
                    int r = t[i];
                    try { tbl.Rows.Item(r).Delete(); }
                    catch { dynamic sel = app.Selection; tbl.Cell(r, 1).Select(); sel.Rows.Delete(); }
                }
            });
            return "Đã xóa " + t.Count + " hàng: " + IndexList.Format(t) + ".";
        }

        public string DeleteCols(string manual)
        {
            var s = Need(false); dynamic tbl = s.Table;
            List<int> t = Targets(s, manual, false);
            if (t.Count >= s.ColCount) throw new UserMsg("Không thể xóa toàn bộ các cột của bảng bằng công cụ này.");
            Undo("Xóa cột", delegate
            {
                for (int i = t.Count - 1; i >= 0; i--)
                {
                    int c = t[i];
                    try { tbl.Columns.Item(c).Delete(); }
                    catch { dynamic sel = app.Selection; tbl.Cell(1, c).Select(); sel.Columns.Delete(); }
                }
            });
            return "Đã xóa " + t.Count + " cột: " + IndexList.Format(t) + ".";
        }

        // ------------------------------------------------------------------ ghi chữ vào ô (giữ định dạng ký tự đầu)
        internal static void SetCellText(dynamic tbl, int r, int c, string text)
        {
            dynamic rg = tbl.Cell(r, c).Range;
            try { rg.MoveEnd(1, -1); } catch { }       // bỏ dấu kết thúc ô
            rg.Text = text;
        }

        /// <summary>Đọc chữ của các ô trong bản chụp, theo đúng thứ tự.</summary>
        static string[] ReadTexts(dynamic tbl, List<int[]> cells)
        {
            var res = new string[cells.Count];
            for (int i = 0; i < cells.Count; i++)
            {
                try { res[i] = CellStr(tbl.Cell(cells[i][0], cells[i][1])); } catch { res[i] = ""; }
            }
            return res;
        }

        // ------------------------------------------------------------------ tính toán, điền một lần
        public string Calc(AppSettings st)
        {
            var s = Need(false); dynamic tbl = s.Table;
            if (IsIrregular(tbl)) return CalcIrregular(st, s);   // Rev0.9.5.13: bảng có ô gộp (Word đánh số ô theo vị trí trong từng hàng)
            string dir = CalcPlanner.ResolveDirection("auto", s.R1, s.C1, s.R2, s.C2);   // tự động, không còn cài đặt tay
            string label = NumberUtil.OpLabel(st.CalcOp);
            string message = "";
            // Dấu thập phân của vùng: lấy từ các số không mơ hồ trong vùng chọn.
            var selTexts = new List<string>();
            foreach (int[] rc in s.Cells) { try { selTexts.Add(CellStr(tbl.Cell(rc[0], rc[1]))); } catch { } }
            char hint = NumberUtil.DetectHint(selTexts);
            if (hint == '\0') hint = NumberUtil.HintForStyle(st.NumStyle);
            // Rev0.9.1A: kiểm tra trước khi sửa bảng, nếu không ô nào khớp định dạng thì dừng để không thêm cột/hàng trống
            if (st.CalcFmtOn)
            {
                bool anyMatch = false;
                foreach (int[] rc in s.Cells) { try { if (CellFmtMatch(tbl.Cell(rc[0], rc[1]), st.CalcFmtKind)) { anyMatch = true; break; } } catch { } }
                if (!anyMatch) return "Không có ô nào khớp định dạng «" + FmtKindLabel(st.CalcFmtKind) + "» trong vùng chọn.";
            }
            Undo("Trợ lý: " + label, delegate
            {
                int rowCount = s.RowCount, colCount = s.ColCount;
                if (dir == "row" && s.C2 >= colCount)
                {
                    tbl.Columns.Add(); colCount++;
                    if (s.R1 > 1) { dynamic h = tbl.Cell(1, colCount); if (CellStr(h).Length == 0) h.Range.Text = label; }
                }
                if (dir == "col" && s.R2 >= rowCount)
                {
                    tbl.Rows.Add(); rowCount++;
                    if (s.C1 > 1) { dynamic h = tbl.Cell(rowCount, 1); if (CellStr(h).Length == 0) h.Range.Text = label; }
                }
                List<CalcJob> jobs = CalcPlanner.Plan(dir, s.R1, s.C1, s.R2, s.C2);
                int filled = 0, skipped = 0, excluded = 0;
                string last = "";
                foreach (CalcJob j in jobs)
                {
                    var nums = new List<double>();
                    int md = 0; bool vn = false, th = false;
                    foreach (int[] rc in j.Sources)
                    {
                        string tx;
                        dynamic srcCell;
                        try { srcCell = tbl.Cell(rc[0], rc[1]); tx = CellStr(srcCell); } catch { continue; }
                        if (st.CalcFmtOn && !CellFmtMatch(srcCell, st.CalcFmtKind)) { excluded++; continue; }
                        ParsedNumber p = NumberUtil.Parse(tx, hint);
                        if (p != null)
                        {
                            nums.Add(p.V); if (p.Decimals > md) md = p.Decimals;
                            if (p.Dec == ',' && p.Decimals > 0) vn = true;
                            if (p.Thou != '\0') th = true;
                        }
                        else if (tx.Length > 0) skipped++;
                    }
                    if (nums.Count == 0) continue;
                    double v = NumberUtil.Aggregate(st.CalcOp, nums);
                    int d = st.CalcDec == "auto" ? NumberUtil.AutoDecimals(st.CalcOp, md) : int.Parse(st.CalcDec, CultureInfo.InvariantCulture);
                    string style = vn ? (st.NumStyle == "si" ? "si" : "vn") : (md > 0 ? "us" : st.NumStyle);
                    last = NumberUtil.Format(v, d, style, th || st.NumThou);
                    try { SetCellText(tbl, j.Target[0], j.Target[1], last); filled++; } catch { }
                }
                if (filled == 0)
                {
                    message = st.CalcFmtOn && excluded > 0
                        ? "Không có ô nào khớp định dạng «" + FmtKindLabel(st.CalcFmtKind) + "» trong vùng chọn."
                        : "Không có ô số nào trong vùng chọn để tính.";
                    return;
                }
                message = "Đã điền " + filled + " kết quả (" + label + (filled == 1 ? " = " + last : "") + ") vào "
                    + (dir == "row" ? "cột bên phải" : "hàng bên dưới") + "."
                    + (skipped > 0 ? " Bỏ qua " + skipped + " ô không phải số." : "")
                    + (st.CalcFmtOn && excluded > 0 ? " Chỉ tính ô «" + FmtKindLabel(st.CalcFmtKind) + "», bỏ qua " + excluded + " ô khác." : "");
            });
            return message;
        }

        /// <summary>Bảng có ô gộp (hàng có số ô khác nhau hoặc gộp dọc): Cell(hàng, cột) đánh số theo vị trí trong từng hàng nên không còn khớp theo cột nhìn thấy.</summary>
        static bool IsIrregular(dynamic tbl)
        {
            // Đã đo trên Word 16 thật: Columns.Count không báo lỗi với hàng gộp ngang (trả số ô của hàng đầu), nên dùng Uniform và đối chiếu số ô.
            try { if (!(bool)tbl.Uniform) return true; } catch { }
            try { int rows = (int)tbl.Rows.Count, cols = (int)tbl.Columns.Count, cells = (int)tbl.Range.Cells.Count; return cells != rows * cols; }
            catch { return true; }   // gộp dọc: Rows.Count báo lỗi
        }

        /// <summary>Tính tổng, trung bình, nhỏ nhất, lớn nhất, đếm cho bảng có ô gộp. Ô được xác định theo vị trí ngang thật (cộng dồn độ rộng ô từng hàng),
        /// nên ô nằm thẳng cột bên dưới luôn đúng dù hàng đó có ô gộp. Cột ghi kết quả là ô ngay dưới ô cuối của cột được chọn, hàng ghi là ô ngay bên phải.</summary>
        string CalcIrregular(AppSettings st, SelInfo s)
        {
            dynamic tbl = s.Table;
            string label = NumberUtil.OpLabel(st.CalcOp);
            TableGrid grid = TableGrid.Build(tbl);
            var sel = new List<GridCell>();
            foreach (int[] rc in s.Cells) { GridCell gc = grid.At(rc[0], rc[1]); if (gc != null) sel.Add(gc); }
            if (sel.Count == 0) return "Không đọc được các ô đang chọn trong bảng.";
            var rowSet = new HashSet<int>(); var colKeys = new HashSet<string>();
            foreach (GridCell g in sel) { rowSet.Add(g.R); colKeys.Add(g.ColKey); }
            string dir = rowSet.Count == 1 ? "row" : (colKeys.Count == 1 ? "col" : (rowSet.Count > colKeys.Count ? "col" : "row"));

            var texts = new List<string>();
            foreach (GridCell g in sel) { try { texts.Add(CellStr(tbl.Cell(g.R, g.C))); } catch { } }
            char hint = NumberUtil.DetectHint(texts);
            if (hint == '\0') hint = NumberUtil.HintForStyle(st.NumStyle);
            if (st.CalcFmtOn)
            {
                bool anyMatch = false;
                foreach (GridCell g in sel) { try { if (CellFmtMatch(tbl.Cell(g.R, g.C), st.CalcFmtKind)) { anyMatch = true; break; } } catch { } }
                if (!anyMatch) return "Không có ô nào khớp định dạng «" + FmtKindLabel(st.CalcFmtKind) + "» trong vùng chọn.";
            }

            // nhóm nguồn và xác định ô đích
            var groups = new List<List<GridCell>>();
            if (dir == "col")
            {
                var by = new Dictionary<string, List<GridCell>>();
                foreach (GridCell g in sel) { List<GridCell> l; if (!by.TryGetValue(g.ColKey, out l)) { l = new List<GridCell>(); by[g.ColKey] = l; groups.Add(l); } l.Add(g); }
                foreach (List<GridCell> l in groups) l.Sort(delegate(GridCell a, GridCell b) { return a.R.CompareTo(b.R); });
            }
            else
            {
                var by = new Dictionary<int, List<GridCell>>();
                foreach (GridCell g in sel) { List<GridCell> l; if (!by.TryGetValue(g.R, out l)) { l = new List<GridCell>(); by[g.R] = l; groups.Add(l); } l.Add(g); }
                foreach (List<GridCell> l in groups) l.Sort(delegate(GridCell a, GridCell b) { return a.C.CompareTo(b.C); });
            }

            string message = "";
            Undo("Trợ lý: " + label, delegate
            {
                bool added = false;
                int filled = 0, skipped = 0, excluded = 0, noTarget = 0;
                string last = "";
                foreach (List<GridCell> src in groups)
                {
                    GridCell lastSrc = src[src.Count - 1];
                    GridCell target = dir == "col" ? grid.Below(lastSrc) : grid.Right(lastSrc);
                    if (target == null && !added)
                    {
                        // hết bảng: thêm một hàng (cột) mới; bảng có ô gộp thì Word có thể từ chối thêm cột
                        try { if (dir == "col") tbl.Rows.Add(); else tbl.Columns.Add(); added = true; grid = TableGrid.Build(tbl); }
                        catch { }
                        GridCell again = grid.At(lastSrc.R, lastSrc.C);
                        if (again != null) { lastSrc = again; target = dir == "col" ? grid.Below(again) : grid.Right(again); }
                    }
                    else if (target == null && added)
                    {
                        GridCell again = grid.At(lastSrc.R, lastSrc.C);
                        if (again != null) target = dir == "col" ? grid.Below(again) : grid.Right(again);
                    }
                    if (target == null) { noTarget++; continue; }
                    var nums = new List<double>();
                    int md = 0; bool vn = false, th = false;
                    foreach (GridCell g in src)
                    {
                        if (g.R == target.R && g.C == target.C) continue;
                        string tx; dynamic srcCell;
                        try { srcCell = tbl.Cell(g.R, g.C); tx = CellStr(srcCell); } catch { continue; }
                        if (st.CalcFmtOn && !CellFmtMatch(srcCell, st.CalcFmtKind)) { excluded++; continue; }
                        ParsedNumber pn = NumberUtil.Parse(tx, hint);
                        if (pn != null)
                        {
                            nums.Add(pn.V); if (pn.Decimals > md) md = pn.Decimals;
                            if (pn.Dec == ',' && pn.Decimals > 0) vn = true;
                            if (pn.Thou != '\0') th = true;
                        }
                        else if (tx.Length > 0) skipped++;
                    }
                    if (nums.Count == 0) continue;
                    double v = NumberUtil.Aggregate(st.CalcOp, nums);
                    int d = st.CalcDec == "auto" ? NumberUtil.AutoDecimals(st.CalcOp, md) : int.Parse(st.CalcDec, CultureInfo.InvariantCulture);
                    string style = vn ? (st.NumStyle == "si" ? "si" : "vn") : (md > 0 ? "us" : st.NumStyle);
                    last = NumberUtil.Format(v, d, style, th || st.NumThou);
                    try { SetCellText(tbl, target.R, target.C, last); filled++; } catch { }
                }
                if (filled == 0)
                {
                    message = noTarget > 0
                        ? "Không có ô để điền kết quả (bảng có ô gộp, không thêm được " + (dir == "col" ? "hàng" : "cột") + " mới). Hãy chọn thêm một ô trống " + (dir == "col" ? "bên dưới" : "bên phải") + " rồi tính lại."
                        : (st.CalcFmtOn && excluded > 0 ? "Không có ô nào khớp định dạng «" + FmtKindLabel(st.CalcFmtKind) + "» trong vùng chọn." : "Không có ô số nào trong vùng chọn để tính.");
                    return;
                }
                message = "Đã điền " + filled + " kết quả (" + label + (filled == 1 ? " = " + last : "") + ") vào "
                    + (dir == "row" ? "ô bên phải" : "ô thẳng cột bên dưới") + " (bảng có ô gộp)."
                    + (skipped > 0 ? " Bỏ qua " + skipped + " ô không phải số." : "")
                    + (noTarget > 0 ? " " + noTarget + " nhóm không có ô để điền." : "");
            });
            return message;
        }

        /// <summary>Ô có khớp kiểu định dạng cần lọc không (Rev0.9.1A, thử nghiệm). Word trả 9999999 khi định dạng lẫn lộn trong ô (ví dụ chỉ một phần chữ đậm): coi là không khớp.</summary>
        static bool CellFmtMatch(dynamic cell, string kind)
        {
            try
            {
                dynamic rg = cell.Range;
                if (kind == "bold") return (int)rg.Font.Bold == -1;
                if (kind == "highlight") { int h = (int)rg.HighlightColorIndex; return h != 0 && h != 9999999; }
                if (kind == "color") { int c = (int)rg.Font.Color; return c != -16777216 && c != 9999999; }
                if (kind == "fill") { int f = (int)cell.Shading.BackgroundPatternColor; return f != -16777216 && f != 9999999; }
            }
            catch { }
            return false;
        }

        static string FmtKindLabel(string kind)
        {
            if (kind == "highlight") return "tô sáng";
            if (kind == "color") return "màu chữ";
            if (kind == "fill") return "màu nền ô";
            return "chữ đậm";
        }

        struct CellUpdate
        {
            public int R, C;
            public string Text;
            public CellUpdate(int r, int c, string text) { R = r; C = c; Text = text; }
        }

        static bool TryCalcRowProd(dynamic tbl, int curRow, int curCol, int colCount, TableGrid grid, AppSettings st, out string resultStr, out int numCount)
        {
            resultStr = null;
            numCount = 0;

            var colList = new List<int>();
            if (grid != null) colList.AddRange(grid.RowCols(curRow)); else for (int cc = 1; cc <= colCount; cc++) colList.Add(cc);   // bảng có ô gộp: dùng đúng các ô có trong hàng
            var rowTexts = new List<string>();
            foreach (int c in colList)
            {
                if (c == curCol) continue;
                try { rowTexts.Add(CellStr(tbl.Cell(curRow, c))); } catch { }
            }
            char hint = NumberUtil.DetectHint(rowTexts);
            if (hint == '\0') hint = NumberUtil.HintForStyle(st.NumStyle);

            var nums = new List<double>();
            int sumDec = 0; int md = 0; bool vn = false, th = false;
            foreach (int c in colList)
            {
                if (c == curCol) continue;
                dynamic cell;
                try { cell = tbl.Cell(curRow, c); } catch { continue; }
                if (!CellFmtMatch(cell, st.CalcFmtKind)) continue;
                string tx = CellStr(cell);
                ParsedNumber p = NumberUtil.Parse(tx, hint);
                if (p != null)
                {
                    nums.Add(p.V);
                    sumDec += p.Decimals;
                    if (p.Decimals > md) md = p.Decimals;
                    if (p.Dec == ',' && p.Decimals > 0) vn = true;
                    if (p.Thou != '\0') th = true;
                }
            }

            if (nums.Count == 0) return false;

            numCount = nums.Count;
            double v = NumberUtil.Aggregate("prod", nums);
            int d = st.CalcDec == "auto" || string.IsNullOrEmpty(st.CalcDec) ? Math.Min(6, sumDec) : int.Parse(st.CalcDec, CultureInfo.InvariantCulture);
            string style = vn ? (st.NumStyle == "si" ? "si" : "vn") : (md > 0 ? "us" : st.NumStyle);
            resultStr = NumberUtil.Format(v, d, style, th || st.NumThou);
            return true;
        }

        public string CalcProdFmt(AppSettings st)
        {
            var s = Need(false); dynamic tbl = s.Table;
            if (s.Cells == null || s.Cells.Count == 0)
                return "Hãy bôi đen các ô trong cùng một cột (hoặc đặt con trỏ trong một ô).";

            int colCount = s.ColCount;
            TableGrid pgrid = null;
            try { colCount = (int)tbl.Columns.Count; } catch { pgrid = TableGrid.Build(tbl); }   // bảng có ô gộp
            if (colCount <= 0) colCount = s.C2;

            bool sameCol = true;
            if (pgrid == null)
            {
                int targetCol = s.Cells[0][1];
                for (int i = 1; i < s.Cells.Count; i++) if (s.Cells[i][1] != targetCol) sameCol = false;
            }
            else
            {
                string key0 = null;
                for (int i = 0; i < s.Cells.Count; i++)
                {
                    GridCell gcc = pgrid.At(s.Cells[i][0], s.Cells[i][1]);
                    string ck = gcc != null ? gcc.ColKey : "?" + i;
                    if (key0 == null) key0 = ck; else if (ck != key0) sameCol = false;
                }
            }
            if (!sameCol)
                return "Hãy bôi đen các ô trong cùng một cột (hoặc đặt con trỏ trong một ô).";

            if (s.Cells.Count == 1)
            {
                int curRow = s.Cells[0][0];
                int curCol = s.Cells[0][1];
                string resultStr;
                int numCount;
                if (!TryCalcRowProd(tbl, curRow, curCol, colCount, pgrid, st, out resultStr, out numCount))
                    return "Không có ô nào khớp định dạng «" + FmtKindLabel(st.CalcFmtKind) + "» trong hàng chứa con trỏ.";

                Undo("Trợ lý: Tích theo định dạng", delegate
                {
                    SetCellText(tbl, curRow, curCol, resultStr);
                });

                return string.Format("Tích của {0} số = {1}", numCount, resultStr);
            }

            var updates = new List<CellUpdate>();
            int k = 0, m = 0;
            for (int i = 0; i < s.Cells.Count; i++)
            {
                int curRow = s.Cells[i][0];
                int curCol = s.Cells[i][1];
                string resultStr;
                int numCount;
                if (TryCalcRowProd(tbl, curRow, curCol, colCount, pgrid, st, out resultStr, out numCount))
                {
                    updates.Add(new CellUpdate(curRow, curCol, resultStr));
                    k++;
                }
                else
                {
                    m++;
                }
            }

            if (k == 0)
                return "Không có ô nào khớp định dạng «" + FmtKindLabel(st.CalcFmtKind) + "» trong các hàng được bôi đen.";

            Undo("Trợ lý: Tích theo định dạng", delegate
            {
                for (int i = 0; i < updates.Count; i++)
                {
                    SetCellText(tbl, updates[i].R, updates[i].C, updates[i].Text);
                }
            });

            return string.Format("Đã tính tích cho {0} hàng (bỏ qua {1} hàng không có ô khớp).", k, m);
        }

        // ------------------------------------------------------------------ bố cục
        public string SetLayout(bool repeatHeader, bool noBreak)
        {
            var s = Need(false); dynamic tbl = s.Table;
            Undo("Bố cục bảng", delegate
            {
                SetB(delegate { return (bool)tbl.Rows.Item(1).HeadingFormat; }, delegate(bool v) { tbl.Rows.Item(1).HeadingFormat = v; }, repeatHeader);
                SetB(delegate { return (bool)tbl.Rows.AllowBreakAcrossPages; }, delegate(bool v) { tbl.Rows.AllowBreakAcrossPages = v; }, !noBreak);
            });
            return "Đã đặt: " + (repeatHeader ? "lặp dòng tiêu đề qua trang" : "không lặp dòng tiêu đề")
                + "; " + (noBreak ? "không ngắt hàng giữa hai trang" : "cho phép ngắt hàng giữa hai trang") + ".";
        }

        /// <summary>
        /// Bật hoặc tắt "lặp dòng tiêu đề" (kèm không ngắt hàng giữa hai trang) cho nhiều bảng một lúc.
        /// scope: "sel" = các bảng nằm trong phần đang chọn (bật nếu có bảng chưa bật, ngược lại tắt); "allOn", "allOff" = mọi bảng trong tài liệu.
        /// </summary>
        public string SetRepeatHeader(string scope)
        {
            if (!HasDocument()) throw new UserMsg("Chưa có tài liệu Word nào đang mở.");
            dynamic doc = app.ActiveDocument;
            var tbls = new List<object>();
            if (scope == "sel")
            {
                dynamic st = app.Selection.Tables; int n = (int)st.Count;
                for (int i = 1; i <= n; i++) tbls.Add(st.Item(i));
                if (tbls.Count == 0) throw new UserMsg("Hãy đặt con trỏ trong một bảng, hoặc bôi đen từ bảng này sang bảng khác (hoặc chọn Bật cho mọi bảng ở mũi tên bên cạnh).");
            }
            else
            {
                int n = (int)doc.Tables.Count;
                if (n == 0) throw new UserMsg("Tài liệu chưa có bảng nào.");
                for (int i = 1; i <= n; i++) tbls.Add(doc.Tables.Item(i));
            }
            bool target;
            if (scope == "allOn") target = true;
            else if (scope == "allOff") target = false;
            else
            {
                target = false;
                foreach (dynamic t in tbls) { try { if (!(bool)t.Rows.Item(1).HeadingFormat) { target = true; break; } } catch { target = true; break; } }
            }
            int ok = 0, fail = 0;
            Undo("Lặp tiêu đề", delegate
            {
                foreach (dynamic t in tbls)
                {
                    try
                    {
                        dynamic tb = t;
                        SetB(delegate { return (bool)tb.Rows.Item(1).HeadingFormat; }, delegate(bool v) { tb.Rows.Item(1).HeadingFormat = v; }, target);
                        if (target) SetB(delegate { return (bool)tb.Rows.AllowBreakAcrossPages; }, delegate(bool v) { tb.Rows.AllowBreakAcrossPages = v; }, false);
                        ok++;
                    }
                    catch (Exception ex) { fail++; Log.W("SetRepeatHeader: " + ex.Message); }
                }
            });
            string what = scope == "sel" ? (tbls.Count == 1 ? "bảng đang chọn" : tbls.Count + " bảng trong phần đang chọn") : "mọi bảng (" + tbls.Count + ")";
            return (target ? "Đã bật lặp dòng tiêu đề cho " : "Đã tắt lặp dòng tiêu đề cho ") + what + (target ? ", không ngắt hàng giữa hai trang" : "") + "."
                + (fail > 0 ? " " + fail + " bảng không đặt được." : "") + (scope == "sel" ? " Bấm lại để " + (target ? "tắt." : "bật.") : "");
        }

        // ------------------------------------------------------------------ định dạng số
        public string FormatSelectedNumbers(AppSettings st)
        {
            var s = Need(false); dynamic tbl = s.Table;
            string[] texts = ReadTexts(tbl, s.Cells);
            // Bằng chứng về dấu thập phân: các số không mơ hồ ngay trong vùng chọn; không có thì theo kiểu đích.
            char evidence = NumberUtil.DetectHint(texts);
            char hint = evidence != '\0' ? evidence : NumberUtil.HintForStyle(st.NumStyle);
            int total = 0, changed = 0, same = 0, ambig = 0;
            var todo = new List<int>(); var outText = new List<string>();
            for (int i = 0; i < texts.Length; i++)
            {
                ParsedNumber p0 = NumberUtil.Parse(texts[i], '\0');
                if (p0 == null) continue;
                total++;
                if (p0.Amb) ambig++;
                string f = NumberUtil.Reformat(texts[i], st.NumStyle, st.NumDec, st.NumThou, st.NumAddDecToInt, hint);
                if (f == texts[i]) { same++; continue; }
                todo.Add(i); outText.Add(f);
            }
            if (total == 0) return "Không có ô số nào trong vùng chọn.";
            Undo("Định dạng số trong bảng", delegate
            {
                for (int k = 0; k < todo.Count; k++)
                {
                    int[] rc = s.Cells[todo[k]];
                    SetCellText(tbl, rc[0], rc[1], outText[k]);
                    changed++;
                }
            });
            string msg = "Đã định dạng " + changed + " ô, " + same + " ô đã đúng sẵn, trên tổng " + total + " ô số.";
            if (ambig > 0)
                msg += " Có " + ambig + " ô dạng 1,234 hoặc 1.234 mơ hồ: "
                    + (evidence != '\0' ? "đã hiểu theo dấu thập phân '" + evidence + "' của các số rõ ràng khác trong vùng chọn." : "không có số rõ ràng nào để đối chiếu nên đã hiểu theo kiểu đích đang chọn.");
            return msg;
        }

        // ------------------------------------------------------------------ xuất bảng
        public string[][] ReadTable(dynamic tbl)
        {
            var rows = new List<List<string>>();
            dynamic cells = tbl.Range.Cells;
            int n = (int)cells.Count, lastRow = -1;
            for (int i = 1; i <= n; i++)
            {
                dynamic cell = cells.Item(i);
                int r = (int)cell.RowIndex;
                if (r != lastRow) { rows.Add(new List<string>()); lastRow = r; }
                rows[rows.Count - 1].Add(CellStr(cell));
            }
            var res = new string[rows.Count][];
            for (int i = 0; i < rows.Count; i++) res[i] = rows[i].ToArray();
            return res;
        }

        public string ToCsv()
        {
            var s = Need(false);
            var sb = new StringBuilder();
            foreach (string[] row in ReadTable(s.Table))
            {
                for (int i = 0; i < row.Length; i++) { if (i > 0) sb.Append(','); sb.Append(CsvUtil.Escape(row[i])); }
                sb.Append("\r\n");
            }
            return sb.ToString();
        }

        public string ToTabText()
        {
            var s = Need(false);
            var sb = new StringBuilder();
            foreach (string[] row in ReadTable(s.Table)) { sb.Append(string.Join("\t", row)); sb.Append("\r\n"); }
            return sb.ToString();
        }

        // ------------------------------------------------------------------ MASTER STYLE
        void SetBorders(dynamic borders, AppSettings st, string part = null)
        {
            int ls = WordMap.LineStyleEnum(st.BorderStyle), lw = WordMap.LineWidthEnum(st.BorderPt), col = AppSettings.HexToWord(st.BorderColor);
            int[] ids = { -1, -2, -3, -4, -5, -6 };   // trên, trái, dưới, phải, ngang trong, dọc trong
            foreach (int id in ids)
            {
                dynamic b = borders.Item(id);
                if (part == null || part == "style")
                    SetI(delegate { return (int)b.LineStyle; }, delegate(int v) { b.LineStyle = v; }, ls);
                if (part == null || part == "width")
                    SetI(delegate { return (int)b.LineWidth; }, delegate(int v) { b.LineWidth = v; }, lw);
                if (part == null || part == "color")
                    SetI(delegate { return (int)b.Color; }, delegate(int v) { b.Color = v; }, col);
            }
        }

        /// <summary>Tạo hoặc cập nhật Table Style "TroLy_Master" trong tài liệu theo cài đặt (chỉ ghi chỗ khác).</summary>
        /// <summary>update = false (áp dụng cho một bảng): chỉ tạo style nếu chưa có, KHÔNG sửa style đang có, vì style dùng chung cho mọi bảng đã gán master style
        /// (sửa màu nền dòng tiêu đề trong style sẽ làm đổi cả các bảng khác). Màu, viền... của bảng này đã được ghi trực tiếp vào từng ô.</summary>
        public void EnsureStyle(AppSettings st, bool update = true)
        {
            dynamic doc = app.ActiveDocument;
            dynamic style = null;
            try { style = doc.Styles.Item(StyleName); } catch { style = null; }
            if (style != null && !update) return;
            if (style == null) { style = doc.Styles.Add(StyleName, 3); nSet++; }   // wdStyleTypeTable
            SetS(delegate { return (string)style.Font.Name; }, delegate(string v) { style.Font.Name = v; }, st.FontName);
            SetF(delegate { return (double)style.Font.Size; }, delegate(double v) { style.Font.Size = (float)v; }, st.FontSize);
            dynamic ts = style.Table;
            SetBorders(ts.Borders, st);
            try { ts.RowStripe = 1; } catch { }
            try { ts.AllowBreakAcrossPage = !st.NoRowBreak; } catch { }
            dynamic hd = ts.Condition(0);                // wdFirstRow
            int hb = AppSettings.HexToWord(st.HdrBg), hf = AppSettings.HexToWord(st.HdrFg);
            SetI(delegate { return (int)hd.Shading.BackgroundPatternColor; }, delegate(int v) { hd.Shading.BackgroundPatternColor = v; }, hb);
            SetI(delegate { return (int)hd.Font.Color; }, delegate(int v) { hd.Font.Color = v; }, hf);
            SetI(delegate { return (int)hd.Font.Bold; }, delegate(int v) { hd.Font.Bold = v; }, -1);
            SetI(delegate { return (int)hd.ParagraphFormat.Alignment; }, delegate(int v) { hd.ParagraphFormat.Alignment = v; }, st.HdrAlign);
            dynamic ev = ts.Condition(3);                // wdEvenRowBanding
            int zc = AppSettings.HexToWord(st.ZebraColor);
            SetI(delegate { return (int)ev.Shading.BackgroundPatternColor; }, delegate(int v) { ev.Shading.BackgroundPatternColor = v; }, zc);
        }

        // Định dạng trực tiếp từng ô (không phụ thuộc xung đột giữa Table Style và Paragraph Style "Normal" của tài liệu).
        // Table Style TroLy_Master vẫn được gán để Word nhận biết bảng nào thuộc master style.
        // Mỗi thuộc tính được đọc trước, chỉ ghi khi khác giá trị mong muốn.
        void ApplyToTable(dynamic tbl, AppSettings st)
        {
            dynamic doc = app.ActiveDocument;
            if (only == null)
            {
                string curStyle = "";
                try { curStyle = (string)tbl.Style.NameLocal; } catch { curStyle = ""; }
                if (curStyle != StyleName) { tbl.Style = StyleName; nSet++; } else nSame++;
                SetB(delegate { return (bool)tbl.ApplyStyleHeadingRows; }, delegate(bool v) { tbl.ApplyStyleHeadingRows = v; }, true);
                SetB(delegate { return (bool)tbl.ApplyStyleLastRow; }, delegate(bool v) { tbl.ApplyStyleLastRow = v; }, false);
                SetB(delegate { return (bool)tbl.ApplyStyleFirstColumn; }, delegate(bool v) { tbl.ApplyStyleFirstColumn = v; }, false);
                SetB(delegate { return (bool)tbl.ApplyStyleLastColumn; }, delegate(bool v) { tbl.ApplyStyleLastColumn = v; }, false);
                SetB(delegate { return (bool)tbl.ApplyStyleRowBands; }, delegate(bool v) { tbl.ApplyStyleRowBands = v; }, false);       // nền xen kẽ do Trợ lý tô trực tiếp
                SetB(delegate { return (bool)tbl.ApplyStyleColumnBands; }, delegate(bool v) { tbl.ApplyStyleColumnBands = v; }, false);
            }
            dynamic range = tbl.Range;
            if (only == null)
            {
                SetBorders(tbl.Borders, st);
            }
            else
            {
                if (Want("borderClr")) SetBorders(tbl.Borders, st, "color");
                if (Want("borderSt")) SetBorders(tbl.Borders, st, "style");
                if (Want("borderW")) SetBorders(tbl.Borders, st, "width");
            }
            if (Want("font"))
            {
                SetS(delegate { return (string)range.Font.Name; }, delegate(string v) { range.Font.Name = v; }, st.FontName);
                SetF(delegate { return (double)range.Font.Size; }, delegate(double v) { range.Font.Size = (float)v; }, st.FontSize);
            }
            dynamic pf = range.ParagraphFormat;
            if (Want("space"))
            {
                SetF(delegate { return (double)pf.SpaceBefore; }, delegate(double v) { pf.SpaceBefore = (float)v; }, st.CellSpace);
                SetF(delegate { return (double)pf.SpaceAfter; }, delegate(double v) { pf.SpaceAfter = (float)v; }, st.CellSpace);
                SetI(delegate { return (int)pf.LineSpacingRule; }, delegate(int v) { pf.LineSpacingRule = v; }, 0);   // đơn
            }
            if (only == null && st.ClearManual)
            {
                // Thay cho ParagraphFormat.Reset (nặng và xóa cả căn lề): chỉ đưa các thụt lề về 0
                SetF(delegate { return (double)pf.LeftIndent; }, delegate(double v) { pf.LeftIndent = (float)v; }, 0);
                SetF(delegate { return (double)pf.RightIndent; }, delegate(double v) { pf.RightIndent = (float)v; }, 0);
                SetF(delegate { return (double)pf.FirstLineIndent; }, delegate(double v) { pf.FirstLineIndent = (float)v; }, 0);
            }
            if (only == null)
            {
                dynamic rc = range.Cells;
                SetI(delegate { return (int)rc.VerticalAlignment; }, delegate(int v) { rc.VerticalAlignment = v; }, 1);   // giữa theo chiều dọc
            }
            if (only == null || Want("hdrBg") || Want("hdrFg") || Want("hdrAl") || Want("align") || Want("zebra"))
            {
                int hdrBg = AppSettings.HexToWord(st.HdrBg), hdrFg = AppSettings.HexToWord(st.HdrFg), zeb = AppSettings.HexToWord(st.ZebraColor);
                dynamic cells = range.Cells;
                int n = (int)cells.Count;
                dynamic firstHdr = null, lastHdr = null, firstBody = null;
                for (int i = 1; i <= n; i++)
                {
                    dynamic cell = cells.Item(i);
                    int ri = (int)cell.RowIndex;
                    if (ri == 1)
                    {
                        if (firstHdr == null) firstHdr = cell;
                        lastHdr = cell;
                        if (Want("hdrBg"))
                            SetI(delegate { return (int)cell.Shading.BackgroundPatternColor; }, delegate(int v) { cell.Shading.BackgroundPatternColor = v; }, hdrBg);
                    }
                    else
                    {
                        if (firstBody == null) firstBody = cell;
                        if (Want("align"))
                        {
                            string tx = CellStr(cell);
                            int al = NumberUtil.Parse(tx) != null ? st.NumAlign : st.TextAlign;
                            SetI(delegate { return (int)cell.Range.ParagraphFormat.Alignment; }, delegate(int v) { cell.Range.ParagraphFormat.Alignment = v; }, al);
                        }
                        if ((Want("zebra") && (only != null || st.Zebra)) || (only == null && st.ClearManual))
                        {
                            int want = (st.Zebra && (ri - 1) % 2 == 0) ? zeb : -16777216;   // dòng dữ liệu thứ 2, 4, 6 tô màu; còn lại không tô
                            SetI(delegate { return (int)cell.Shading.BackgroundPatternColor; }, delegate(int v) { cell.Shading.BackgroundPatternColor = v; }, want);
                        }
                    }
                }
                if (firstHdr != null)
                {
                    dynamic hr = doc.Range((int)firstHdr.Range.Start, (int)lastHdr.Range.End);
                    if (Want("hdrFg"))
                        SetI(delegate { return (int)hr.Font.Color; }, delegate(int v) { hr.Font.Color = v; }, hdrFg);
                    if (only == null)
                        SetI(delegate { return (int)hr.Font.Bold; }, delegate(int v) { hr.Font.Bold = v; }, -1);
                    if (Want("hdrAl"))
                        SetI(delegate { return (int)hr.ParagraphFormat.Alignment; }, delegate(int v) { hr.ParagraphFormat.Alignment = v; }, st.HdrAlign);
                }
                if (only == null && firstBody != null && st.ClearManual)
                {
                    dynamic br = doc.Range((int)firstBody.Range.Start, (int)range.End);
                    SetI(delegate { return (int)br.Font.Bold; }, delegate(int v) { br.Font.Bold = v; }, 0);
                    SetI(delegate { return (int)br.Font.Color; }, delegate(int v) { br.Font.Color = v; }, -16777216);   // wdColorAutomatic
                    // Giữ nguyên nghiêng, gạch chân, chỉ số trên và dưới (H2S, p_H2S...)
                }
            }
            if (only == null)
            {
                try { SetB(delegate { return (bool)tbl.Rows.Item(1).HeadingFormat; }, delegate(bool v) { tbl.Rows.Item(1).HeadingFormat = v; }, st.RepeatHeader); } catch { }
                try { SetB(delegate { return (bool)tbl.Rows.AllowBreakAcrossPages; }, delegate(bool v) { tbl.Rows.AllowBreakAcrossPages = v; }, !st.NoRowBreak); } catch { }
            }
        }

        string Counts() { return nSet == 0 ? "Không có gì cần đổi, mọi thuộc tính đã đúng sẵn" : "Đã ghi " + nSet + " thuộc tính khác, " + nSame + " thuộc tính đã đúng sẵn nên bỏ qua"; }

        public string ApplyAll(AppSettings st)
        {
            if (!HasDocument()) throw new UserMsg("Chưa có tài liệu Word nào đang mở.");
            dynamic doc = app.ActiveDocument;
            int total = (int)doc.Tables.Count;
            if (total == 0) throw new UserMsg("Tài liệu chưa có bảng nào.");
            int ok = 0; var errs = new List<string>();
            only = null;
            Undo("Áp dụng master style cho mọi bảng", delegate
            {
                EnsureStyle(st);
                for (int i = 1; i <= total; i++)
                {
                    try { ApplyToTable(doc.Tables.Item(i), st); ok++; }
                    catch (Exception ex) { errs.Add("Bảng " + i + ": " + ex.Message); }
                }
            });
            return "Đã áp dụng master style cho " + ok + " trên " + total + " bảng. " + Counts() + "."
                + (errs.Count > 0 ? " Lỗi: " + string.Join(" | ", errs.ToArray()) : "");
        }

        public string ApplyToSelected(AppSettings st)
        {
            var s = Need(false);
            only = null;
            Undo("Áp dụng master style cho bảng đang chọn", delegate { EnsureStyle(st, false); ApplyToTable(s.Table, st); });
            return "Đã áp dụng master style cho bảng thứ " + s.TableNo + ". " + Counts() + ".";
        }

        public string ApplyNow(AppSettings st, string group)
        {
            if (!HasDocument()) return "Đã lưu thông số. Đặt con trỏ trong bảng để áp dụng ngay.";
            var tbls = new List<object>();
            try
            {
                dynamic selTables = app.Selection.Tables;
                int n = (int)selTables.Count;
                for (int i = 1; i <= n; i++) tbls.Add(selTables.Item(i));
            }
            catch { }
            if (tbls.Count == 0) return "Đã lưu thông số. Đặt con trỏ trong bảng để áp dụng ngay.";

            nSet = 0; nSame = 0;
            try
            {
                Undo("Trợ lý: áp dụng ngay", delegate
                {
                    try
                    {
                        foreach (dynamic tbl in tbls)
                        {
                            only = group;
                            EnsureStyle(st, false);
                            ApplyToTable(tbl, st);
                        }
                    }
                    finally
                    {
                        only = null;
                    }
                });
            }
            finally
            {
                only = null;
            }
            return "Đã áp dụng ngay cho " + tbls.Count + " bảng. " + Counts() + ".";
        }

        /// <summary>Canh toàn bảng sang trái (0), giữa (1) hoặc phải (2) so với lề trang giấy. Áp dụng ngay.</summary>
        public string AlignTables(int align)
        {
            if (!HasDocument()) throw new UserMsg("Chưa có tài liệu Word nào đang mở.");
            var tbls = new List<object>();
            try
            {
                dynamic selTables = app.Selection.Tables;
                int n = (int)selTables.Count;
                for (int i = 1; i <= n; i++) tbls.Add(selTables.Item(i));
            }
            catch { }
            if (tbls.Count == 0) throw new UserMsg("Hãy đặt con trỏ vào một ô của bảng trước.");
            string[] names = { "trái", "giữa", "phải" };
            int a = align < 0 ? 0 : (align > 2 ? 2 : align);
            nSet = 0; nSame = 0;
            Undo("Canh toàn bảng " + names[a], delegate
            {
                foreach (dynamic tbl in tbls)
                {
                    SetI(delegate { return (int)tbl.Rows.Alignment; }, delegate(int v) { tbl.Rows.Alignment = v; }, a);
                }
            });
            return "Đã canh " + tbls.Count + " bảng sang " + names[a] + " so với lề trang.";
        }
    }
    /// <summary>Một ô trong bản đồ bảng: R, C là chỉ số hàng và chỉ số ô trong hàng (đúng như Word đánh số), X0 đến X1 là khoảng ngang (pt).</summary>
    internal sealed class GridCell
    {
        public int R, C;
        public double X0, X1;
        public string ColKey { get { return Math.Round(X0).ToString(CultureInfo.InvariantCulture) + "|" + Math.Round(X1).ToString(CultureInfo.InvariantCulture); } }
    }

    /// <summary>Bản đồ ô của bảng có ô gộp. Word đánh số ô theo vị trí trong từng hàng, nên hàng có ô gộp có ít ô hơn và Cell(hàng, cột) không còn
    /// khớp theo cột nhìn thấy. Bản đồ cộng dồn độ rộng ô trong từng hàng để biết ô nào nằm thẳng cột với ô nào.</summary>
    internal sealed class TableGrid
    {
        readonly Dictionary<int, List<GridCell>> byRow = new Dictionary<int, List<GridCell>>();
        public int MaxRow;

        public static TableGrid Build(dynamic tbl)
        {
            var g = new TableGrid();
            dynamic cells = tbl.Range.Cells;
            int n = (int)cells.Count;
            var widths = new Dictionary<GridCell, double>();
            for (int i = 1; i <= n; i++)
            {
                dynamic cell = cells.Item(i);
                var gc = new GridCell { R = (int)cell.RowIndex, C = (int)cell.ColumnIndex };
                double w = 0;
                try { w = Convert.ToDouble((object)cell.Width, CultureInfo.InvariantCulture); } catch { w = 0; }
                widths[gc] = w;
                List<GridCell> l;
                if (!g.byRow.TryGetValue(gc.R, out l)) { l = new List<GridCell>(); g.byRow[gc.R] = l; }
                l.Add(gc);
                if (gc.R > g.MaxRow) g.MaxRow = gc.R;
            }
            foreach (KeyValuePair<int, List<GridCell>> kv in g.byRow)
            {
                kv.Value.Sort(delegate(GridCell a, GridCell b) { return a.C.CompareTo(b.C); });
                double x = 0;
                foreach (GridCell gc in kv.Value) { gc.X0 = x; x += widths[gc]; gc.X1 = x; }
            }
            return g;
        }

        public GridCell At(int r, int c)
        {
            List<GridCell> l;
            if (!byRow.TryGetValue(r, out l)) return null;
            foreach (GridCell g in l) if (g.C == c) return g;
            return null;
        }

        public List<int> RowCols(int r)
        {
            var res = new List<int>();
            List<GridCell> l;
            if (byRow.TryGetValue(r, out l)) foreach (GridCell g in l) res.Add(g.C);
            return res;
        }

        /// <summary>Ô thẳng cột ngay bên dưới (hàng kế tiếp, ô chồng lên khoảng ngang của ô này nhiều nhất); null nếu không có.</summary>
        public GridCell Below(GridCell src)
        {
            List<GridCell> l;
            if (!byRow.TryGetValue(src.R + 1, out l)) return null;
            GridCell best = null; double bestOv = 0.5;
            foreach (GridCell g in l)
            {
                double ov = Math.Min(src.X1, g.X1) - Math.Max(src.X0, g.X0);
                if (ov > bestOv) { bestOv = ov; best = g; }
            }
            return best;
        }

        /// <summary>Ô ngay bên phải trong cùng hàng.</summary>
        public GridCell Right(GridCell src) { return At(src.R, src.C + 1); }
    }
}
