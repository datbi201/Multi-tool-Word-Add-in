// Trợ lý Kỹ thuật - Rev0.4: rà soát toàn văn bản trong Word (quét, xem, sửa, bỏ qua, sửa tất cả).
// Mọi hàm ở đây phải chạy trên luồng chính của Word (qua Host.OnWord). Kết quả trả về là các lớp chữ thuần (DTO), không có đối tượng COM.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    public sealed class RvRow
    {
        public int Index;                 // số thứ tự ổn định của mục trong lần quét
        public string Cat, Sev, CatText, SevText, Before, Show, Reason;
        public string[] FixLabels, FixShow;
        public bool Multi;                // có nhiều cách sửa, cần người dùng chọn
    }

    public sealed class RvResult
    {
        public List<RvRow> Rows = new List<RvRow>();
        public string Message = "";
    }

    public sealed class WordReview
    {
        readonly dynamic app;
        dynamic doc;
        string docName = "";

        sealed class Item
        {
            public Hit H;
            public dynamic Rng;           // Range của Word: tự dịch chuyển theo các chỉnh sửa khác
            public string Before;
            public bool Done;
            public int Start;
            public bool Alone;            // công thức nằm riêng một dòng (cả đoạn): viết bằng Equation chuẩn
        }

        readonly List<Item> items = new List<Item>();
        readonly HashSet<string> ignored = new HashSet<string>();
        int markIdx = -1, markPrev = 0;   // mục đang được tô sáng và màu tô sáng ban đầu của nó
        public const int MarkColor = 7;    // wdYellow
        static readonly Regex AloneTail = new Regex(@"^\s*[.;,:]?\s*(?:\(\d+(?:[.\-]\d+)?\))?\s*$");

        public WordReview(object application) { app = application; }

        /// <summary>fmt = chuẩn hóa công thức, đơn vị, số (thẻ Chuẩn hóa); term = thuật ngữ; style = văn phong; gram = ngữ pháp.</summary>
        public string Mode = "fmt";

        static readonly Regex NumericCellRe = new Regex(@"^[\s]*[-+−]?\d[\d.,\s ]*$");

        // ------------------------------------------------------------------ quét
        public RvResult Scan(AppSettings st)
        {
            var res = new RvResult();
            ClearMark();
            items.Clear();
            try { if ((int)app.Documents.Count == 0) throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            catch (UserMsg) { throw; }
            catch { throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            doc = app.ActiveDocument;
            try { docName = (string)doc.FullName; } catch { docName = ""; }

            if (Mode != "fmt") throw new UserMsg("Hãy dùng lệnh Quét bằng AI.");

            var o = new ReviewOpts();
            o.Chem = st.RvChem; o.Math = st.RvMath; o.Sym = st.RvSym; o.Num = st.RvNum; o.NumTable = st.RvNumTbl;
            o.NumStyle = st.RvNumStyle; o.NumDec = st.NumDec; o.NumThou = st.NumThou; o.NumAddDec = st.NumAddDecToInt;

            dynamic scope = doc.Content;
            int lo = int.MinValue, hi = int.MaxValue; string scopeName = "toàn văn bản";
            if (st.RvScope == "sel")
            {
                try
                {
                    dynamic sel = app.Selection;
                    int a = (int)sel.Start, b = (int)sel.End;
                    if (b > a) { scope = doc.Range(a, b); lo = a; hi = b; scopeName = "phần đang chọn"; }
                    else throw new UserMsg("Chưa bôi đen đoạn nào. Hãy bôi đen phần cần rà soát, hoặc chọn phạm vi toàn văn bản.");
                }
                catch (UserMsg) { throw; }
                catch { }
            }

            int paras = 0, unlocated = 0, already = 0;
            Fast(delegate
            {
                foreach (dynamic p in (IEnumerable)scope.Paragraphs)
                {
                    dynamic pr = p.Range;
                    string text = (string)pr.Text;
                    if (text == null || text.Length < 2) continue;
                    paras++;
                    int ps = (int)pr.Start;
                    bool ctxDone = false; NumCtx ctx = new NumCtx();
                    Func<NumCtx> ctxf = delegate
                    {
                        if (!ctxDone)
                        {
                            ctxDone = true;
                            try { ctx.InTable = (bool)pr.Information(12); } catch { }
                            try { ctx.Heading = (int)p.OutlineLevel < 10; } catch { }
                            if (ctx.InTable) ctx.NumericCell = NumericCellRe.IsMatch(CsvUtil.CellText(text));
                            ctx.Language = (left, right) => LanguageAt((object)doc, ps + left, ps + right);
                        }
                        return ctx;
                    };
                    List<Hit> hits = Detector.Detect(text, o, ctxf);
                    if (hits.Count == 0) continue;
                    bool hasFields = false;
                    try { hasFields = (int)pr.Fields.Count > 0; } catch { }
                    foreach (Hit h in hits)
                    {
                        int s = ps + h.S, e = ps + h.E;
                        if (s < lo || e > hi) continue;
                        if (ignored.Contains(h.Issue.Key)) continue;
                        dynamic rng = doc.Range(s, e);
                        string before = h.Issue.Before;
                        if (hasFields || h.S + before.Length != h.E)
                        {
                            string actual = "";
                            try { actual = (string)rng.Text; } catch { }
                            if (actual != before) { unlocated++; continue; }
                        }
                        if (h.Issue.Fixes.Count > 0 && h.Issue.Fixes[0].NeedsFormat)
                        {
                            // đã là công thức Word hoặc đã có chỉ số trên, dưới: không cần làm lại
                            bool done = false;
                            try { if ((int)rng.OMaths.Count > 0) done = true; } catch { }
                            if (!done)
                            {
                                try { int sub = (int)rng.Font.Subscript, sup = (int)rng.Font.Superscript; if (sub != 0 || sup != 0) done = true; } catch { }
                            }
                            if (done) { already++; continue; }
                        }
                        bool alone = false;
                        if (h.Issue.Cat != "num" && h.Issue.Fixes.Count > 0 && h.Issue.Fixes[0].NeedsFormat && !ctxf().InTable)
                        {
                            string core = text.TrimEnd('\r', '\n', '\a', ' ');
                            if (h.E <= core.Length && core.Substring(0, h.S).Trim().Length == 0 && AloneTail.IsMatch(core.Substring(h.E))) alone = true;
                        }
                        items.Add(new Item { H = h, Rng = rng, Before = before, Start = s, Alone = alone });
                    }
                }
            });

            for (int i = 0; i < items.Count; i++) res.Rows.Add(ToRow(i));
            int err = items.Count(x => x.H.Issue.Sev == "err"), sus = items.Count(x => x.H.Issue.Sev == "sus");
            res.Message = "Đã quét " + paras + " đoạn (" + scopeName + "). Tìm thấy " + items.Count + " mục cần xem: "
                + err + " sai, " + sus + " nghi ngờ, " + (items.Count - err - sus) + " chưa chuẩn."
                + (already > 0 ? " Bỏ qua " + already + " chỗ đã được định dạng sẵn." : "")
                + (o.Stats.AutoText > 0 ? " Tự nhận " + o.Stats.AutoText + " số mang tính văn bản (năm, số trang, số hiệu, ngày) và không đụng tới." : "")
                + (o.Stats.OkNum > 0 ? " " + o.Stats.OkNum + " số kỹ thuật đã đúng quy ước." : "")
                + (unlocated > 0 ? " Không định vị được " + unlocated + " chỗ (có thể do trường hoặc công thức), hãy kiểm tra thủ công." : "");
            return res;
        }

        internal sealed class AiSnapshot
        {
            public object Document;
            public int Start, End, DocumentEnd;
            public string Text;
            public List<AiReviewBlock> Blocks;
            public string Scope;
        }

        internal AiSnapshot CaptureAi(bool selection)
        {
            dynamic active = app.ActiveDocument;
            dynamic content = active.Content;
            dynamic range = content;
            if (selection)
            {
                range = app.Selection.Range;
                if ((int)range.End <= (int)range.Start)
                    throw new UserMsg("Hãy bôi đen phần cần rà soát hoặc chọn Toàn văn bản.");
                WordText.RequireMainStory((object)range);
            }
            int start = (int)range.Start, end = (int)range.End;
            string text = (string)range.Text ?? "";
            if (string.IsNullOrWhiteSpace(text)) throw new UserMsg("Phạm vi rà soát không có nội dung.");
            // Word exposes end-of-cell markers as CR + BEL in Range.Text, but counts only the CR.
            string indexedText = text.Length == end - start ? text : text.Replace("\a", "");
            if (indexedText.Length != end - start)
                throw new UserMsg("Word báo vị trí không khớp với văn bản. Hãy thử chọn phần văn bản thuần.");
            return new AiSnapshot {
                Document = (object)active, Start = start, End = end, DocumentEnd = (int)content.End,
                Text = text, Blocks = AiReview.Blocks(start, indexedText), Scope = selection ? "vùng bôi đen" : "toàn phần nội dung chính"
            };
        }

        internal RvResult AcceptAi(AiSnapshot snapshot, List<AiFinding> findings)
        {
            dynamic active = app.ActiveDocument;
            if (!object.ReferenceEquals(snapshot.Document, (object)active) ||
                (int)active.Content.End != snapshot.DocumentEnd ||
                (string)active.Range(snapshot.Start, snapshot.End).Text != snapshot.Text)
                throw new UserMsg("Tài liệu hoặc nội dung đã thay đổi trong lúc quét. Hãy quét lại.");

            ClearMark();
            items.Clear();
            doc = active;
            try { docName = (string)doc.FullName; } catch { docName = ""; }
            var result = new RvResult();
            foreach (AiFinding finding in findings)
            {
                Issue issue = finding.Issue;
                int end = finding.Start + issue.Before.Length;
                if (finding.Start < snapshot.Start || end > snapshot.End || ignored.Contains(issue.Key)) continue;
                dynamic range = doc.Range(finding.Start, end);
                if ((string)range.Text != issue.Before) continue;
                try { WordText.RequireEditableText((object)range); }
                catch (UserMsg ex) { issue.Fixes.Clear(); issue.Reason += " " + ex.Message; }
                items.Add(new Item { H = new Hit { S = finding.Start, E = end, Issue = issue },
                    Rng = range, Before = issue.Before, Start = finding.Start });
                result.Rows.Add(ToRow(items.Count - 1));
            }
            result.Message = "AI đã rà soát " + snapshot.Scope + ". Có " + items.Count + " mục để xem.";
            return result;
        }

        void Fast(Action a) { WordTables.Fast((object)app, "Chuẩn hóa: quét", a); }

        RvRow ToRow(int i)
        {
            Issue iss = items[i].H.Issue;
            var r = new RvRow();
            r.Index = i; r.Cat = iss.Cat; r.Sev = iss.Sev;
            r.CatText = Detector.CatLabel(iss.Cat); r.SevText = Detector.SevLabel(iss.Sev);
            r.Before = iss.Disp ?? iss.Before; r.Reason = iss.Reason;
            r.FixLabels = iss.Fixes.Select(f => f.Label).ToArray();
            r.FixShow = iss.Fixes.Select(f => f.Show).ToArray();
            r.Show = iss.Fixes.Count > 0 ? iss.Fixes[0].Show : "";
            r.Multi = iss.Fixes.Count > 1;
            return r;
        }

        // ------------------------------------------------------------------ xem, bỏ qua
        Item Get(int idx)
        {
            if (idx < 0 || idx >= items.Count) throw new UserMsg("Mục này không còn trong danh sách. Hãy quét lại.");
            return items[idx];
        }

        void CheckDoc()
        {
            string now = "";
            try { now = (string)app.ActiveDocument.FullName; } catch { }
            if (now != docName || !object.ReferenceEquals((object)doc, (object)app.ActiveDocument)) throw new UserMsg("Tài liệu đang mở đã khác với tài liệu lúc quét. Hãy bấm Quét lại.");
        }

        /// <summary>Chuyển tới mục, tô sáng nó (màu vàng) để dễ nhìn. Màu tô sáng ban đầu được ghi lại để trả về khi sửa, bỏ qua hoặc đóng bảng điều khiển.</summary>
        public string Goto(int idx)
        {
            Item it = Get(idx); CheckDoc();
            WordTables.Fast((object)app, "Chuẩn hóa: đánh dấu mục", delegate
            {
                ClearMarkCore();
                dynamic rng = it.Rng;
                try { app.Selection.SetRange((int)rng.Start, (int)rng.Start); } catch { try { rng.Select(); } catch { } }
                try { markPrev = (int)rng.HighlightColorIndex; if (markPrev == 9999999) markPrev = 0; } catch { markPrev = 0; }
                try { rng.HighlightColorIndex = MarkColor; markIdx = idx; } catch { markIdx = -1; }
                try { app.ActiveWindow.ScrollIntoView(rng, true); } catch { }
            });
            return "Đã chuyển tới vị trí của mục " + (idx + 1) + " (đang tô sáng màu vàng).";
        }

        void ClearMarkCore()
        {
            if (markIdx >= 0 && markIdx < items.Count)
            {
                try { items[markIdx].Rng.HighlightColorIndex = markPrev; } catch { }
            }
            markIdx = -1;
        }

        public bool HasMark { get { return markIdx >= 0; } }

        /// <summary>Trả màu tô sáng về như cũ (gọi khi đóng bảng điều khiển hoặc trước khi quét lại).</summary>
        public void ClearMark()
        {
            if (markIdx < 0) return;
            try { WordTables.Fast((object)app, "Chuẩn hóa: bỏ đánh dấu", delegate { ClearMarkCore(); }); } catch { markIdx = -1; }
        }

        public string Ignore(int idx)
        {
            Item it = Get(idx);
            ClearMark();
            ignored.Add(it.H.Issue.Key);
            it.Done = true;
            return "Đã bỏ qua \"" + it.Before + "\". Mục này sẽ không hiện lại cho tới khi bấm Hiện lại các mục đã bỏ qua.";
        }

        public string GetLongSentence(int idx)
        {
            Item it = Get(idx); CheckDoc();
            if (it.Done || it.H.Issue.Cat != "ai-long" || (string)it.Rng.Text != it.Before)
                throw new UserMsg("Câu dài đã thay đổi. Hãy quét lại và chọn mục Câu dài AI.");
            WordText.RequireEditableText((object)it.Rng);
            return it.Before;
        }

        /// <summary>Thay nội dung mục bằng chuỗi sửa tùy chỉnh (vd kết quả sửa câu dài từ AI) trong 1 bước Undo.</summary>
        public string ReplaceCustom(int idx, string newText)
        {
            if (string.IsNullOrWhiteSpace(newText)) throw new UserMsg("Nội dung thay thế đang trống.");
            Item it = Get(idx); CheckDoc();
            string actual = ""; try { actual = (string)it.Rng.Text; } catch { }
            if (actual != it.Before) throw new UserMsg("Chữ tại vị trí này đã đổi sau khi quét. Hãy bấm Quét lại.");
            WordTables.Fast((object)app, "Chuẩn hóa: sửa bằng AI", delegate
            {
                ClearMarkCore();
                it.Rng.Text = newText.Trim();
            });
            it.Done = true;
            return "Đã áp dụng phương án AI cho mục đã chọn.";
        }

        /// <summary>Dạng của một số để so "giống nhau": số đánh số mục, bảng, hình (1.1, 2.3, 10.2.1: mỗi đoạn một hoặc hai chữ số) cùng dạng theo số đoạn;
        /// số khác thì theo số chữ số (1.234 khác 1.2).</summary>
        static string Shape(string t)
        {
            t = t ?? "";
            Match m = Regex.Match(t, @"^\d{1,2}(?:[.,]\d{1,2})+$");
            if (m.Success) return "#sec" + Regex.Matches(t, @"[.,]").Count;
            if (Regex.IsMatch(t, @"^\d{1,3}(?:[.,]\d{3})+$")) return "#thou" + Regex.Matches(t, @"[.,]").Count;     // 5.678 và 12.345 cùng dạng hàng nghìn
            return Regex.Replace(t, @"\d", "9");
        }

        /// <summary>Hai mục có cùng dạng không. Số (kể cả số đánh số mục, số hiệu bảng, hình): khác chữ số vẫn cùng dạng nếu cùng loại, cùng mức, cùng cách sửa đề xuất.
        /// Các loại khác: phải cùng chữ gốc.</summary>
        static string LanguageAt(object document, int start, int end)
        {
            try
            {
                dynamic wordDoc = document;
                int language = (int)wordDoc.Range(start, end).LanguageID;
                switch (language & 0x3ff)
                {
                    case 9: return "en";
                    case 42: return "vi";
                }
            }
            catch { }
            return "";
        }

        static bool SameKind(Issue a, string aBefore, Issue b, string bBefore)
        {
            if (a.Cat != b.Cat) return false;
            if (a.Cat.StartsWith("ai-", StringComparison.Ordinal)) return a.Key == b.Key && a.Sev == b.Sev;
            if (a.Cat != "num") return aBefore == bBefore;
            if (a.NumStyle != b.NumStyle) return false;
            if (a.Sev == b.Sev && (aBefore == bBefore || Shape(aBefore) == Shape(bBefore))
                && string.Join("|", a.Fixes.Select(q => q.Label).ToArray()) == string.Join("|", b.Fixes.Select(q => q.Label).ToArray())) return true;
            return false;
        }

        /// <summary>Các mục chưa xử lý giống mục này: cùng loại và cùng chữ gốc; riêng số (kể cả đánh số mục) thì cùng dạng là đủ.</summary>
        List<Item> SimilarItems(int idx)
        {
            Item it = Get(idx);
            return items.Where(x => !x.Done && SameKind(x.H.Issue, x.Before, it.H.Issue, it.Before)).ToList();
        }

        public int SimilarCount(int idx) { return SimilarItems(idx).Count; }

        public string IgnoreSimilar(int idx)
        {
            var list = SimilarItems(idx); string before = Get(idx).Before;
            ClearMark();
            foreach (Item x in list) { ignored.Add(x.H.Issue.Key); x.Done = true; }
            bool same = list.All(x => x.Before == before);
            return "Đã bỏ qua " + list.Count + " chỗ " + (same ? "giống \"" + before + "\"" : "cùng dạng với \"" + before + "\" (ví dụ " + string.Join(", ", list.Select(x => x.Before).Distinct().Take(4).ToArray()) + ")") + ". Bấm Hiện lại để thấy chúng.";
        }

        public string FixSimilar(int idx, int fixIdx, AppSettings st)
        {
            Item it0 = Get(idx); CheckDoc();
            Issue issue = it0.H.Issue;
            if (issue.Cat.StartsWith("ai-", StringComparison.Ordinal) && issue.Sev != "err")
                throw new UserMsg("Hãy xem và áp dụng từng đề xuất AI để giữ đúng ngữ cảnh.");
            if (issue.Fixes.Count == 0) throw new UserMsg("Mục này chỉ có thông tin để xem.");
            var list = SimilarItems(idx);
            Fix chosen = it0.H.Issue.Fixes[Math.Max(0, Math.Min(fixIdx, it0.H.Issue.Fixes.Count - 1))];
            int ok = 0, fail = 0, fb = 0; string firstErr = "";
            WordTables.Fast((object)app, "Chuẩn hóa: sửa các chỗ tương tự", delegate
            {
                ClearMarkCore();
                foreach (Item x in list.OrderByDescending(t => t.Start))
                {
                    if (x.H.Issue.Fixes.Count == 0) continue;
                    Fix f = x.H.Issue.Fixes.FirstOrDefault(q => q.Label == chosen.Label) ?? x.H.Issue.Fixes[0];
                    try { if (ApplyOne(x, f, st.EqMode, st.EqStacked)) fb++; ok++; }
                    catch (Exception ex) { fail++; if (firstErr.Length == 0) firstErr = ex.Message; }
                }
            });
            bool sameAll = list.All(x => x.Before == it0.Before);
            return "Đã sửa " + ok + " chỗ " + (sameAll ? "giống \"" + it0.Before + "\" thành \"" + chosen.Show + "\"" : "cùng dạng với \"" + it0.Before + "\" theo cách \"" + chosen.Label + "\"") + "." + (fail > 0 ? " " + fail + " chỗ không sửa được: " + firstErr : "")
                + (fb > 0 ? " " + fb + " chỗ dùng chỉ số trên, dưới thay cho công thức Word." : "");
        }

        public string ClearIgnored()
        {
            int n = ignored.Count; ignored.Clear();
            return n == 0 ? "Chưa có mục nào bị bỏ qua." : "Đã bỏ danh sách bỏ qua (" + n + " mục). Bấm Quét để hiện lại.";
        }

        /// <summary>Gắn kết quả gợi ý câu mới từ AI vào dòng câu dài để hiện dấu ✔ và sửa trực tiếp.</summary>
        public bool AttachAiFix(int idx, string newSentence)
        {
            if (idx < 0 || idx >= items.Count) return false;
            Item it = items[idx];
            if (it == null || it.H == null || it.H.Issue == null) return false;
            CheckDoc();
            if (it.Done || it.H.Issue.Cat != "ai-long" || (string)it.Rng.Text != it.Before || !AiReview.CanReplace(it.Before, newSentence)) return false;
            try { WordText.RequireEditableText((object)it.Rng); }
            catch (UserMsg) { return false; }
            it.H.Issue.Sev = "sus";
            it.H.Issue.Key = it.H.Issue.Key + "|sentence|" + newSentence;
            it.H.Issue.Fixes.Clear();
            it.H.Issue.Fixes.Add(Fix.Text("Dùng câu AI đề xuất", (newSentence ?? "").Trim(), true));
            return true;
        }

        public Issue GetIssue(int idx)
        {
            Item it = Get(idx);
            return it != null && it.H != null ? it.H.Issue : null;
        }

        // ------------------------------------------------------------------ sửa
        /// <summary>
        /// Sửa một mục. Công thức nằm riêng một dòng (Alone) luôn dùng Equation chuẩn của Word. Công thức cùng dòng với văn bản thì theo eqMode:
        /// "glyph" = ký tự chỉ số nhỏ và thấp có sẵn trong font Cambria, "fmt" = định dạng chỉ số trên, dưới.
        /// Trả về true nếu không tạo được Equation và phải dùng định dạng ký tự.
        /// </summary>
        bool ApplyOne(Item it, Fix f, string eqMode, bool stacked)
        {
            dynamic rng = it.Rng;
            string actual = "";
            try { actual = (string)rng.Text; } catch { }
            if (actual != it.Before)
                throw new UserMsg("Văn bản tại vị trí \"" + it.Before + "\" đã bị thay đổi sau lần quét. Hãy bấm Quét lại.");
            if (it.H.Issue.Cat.StartsWith("ai-", StringComparison.Ordinal)) WordText.RequireEditableText((object)rng);
            if (markIdx >= 0 && markIdx < items.Count && items[markIdx] == it) ClearMarkCore();

            string fontName = null;
            float? fontSize = null;
            int? bold = null, italic = null, underline = null, color = null;
            try
            {
                dynamic font = rng.Font;
                object oName = font.Name; if (oName is string && ((string)oName).Length > 0) fontName = (string)oName;
                object oSize = font.Size; if (oSize is float || oSize is double || oSize is int) fontSize = Convert.ToSingle(oSize);
                object oBold = font.Bold; if (oBold is int && (int)oBold != 9999999) bold = (int)oBold;
                object oItalic = font.Italic; if (oItalic is int && (int)oItalic != 9999999) italic = (int)oItalic;
                object oUnder = font.Underline; if (oUnder is int && (int)oUnder != 9999999) underline = (int)oUnder;
                object oColor = font.Color; if (oColor is int && (int)oColor != 9999999) color = (int)oColor;
            }
            catch { }

            bool fb = ApplyFix(doc, rng, f, it.Alone, eqMode, stacked);

            if (!f.NeedsFormat && f.Plain != null)
            {
                try
                {
                    int s = (int)rng.Start;
                    dynamic nr = doc.Range(s, s + f.Plain.Length);
                    dynamic nf = nr.Font;
                    if (fontName != null) nf.Name = fontName;
                    if (fontSize.HasValue && fontSize.Value > 0 && fontSize.Value < 1000) nf.Size = fontSize.Value;
                    if (bold.HasValue) nf.Bold = bold.Value;
                    if (italic.HasValue) nf.Italic = italic.Value;
                    if (underline.HasValue) nf.Underline = underline.Value;
                    if (color.HasValue) nf.Color = color.Value;
                }
                catch { }
            }

            it.Done = true;
            return fb;
        }

        /// <summary>
        /// Thay chữ trong rng bằng công thức đã sửa. Công thức riêng một dòng (alone) hoặc có phân số, căn: dựng Equation của Word bằng OMML.
        /// Còn lại theo eqMode: "glyph" = ký tự chỉ số nhỏ, thấp (Cambria), "fmt" = định dạng chỉ số trên, dưới.
        /// Trả về true nếu không dựng được Equation và phải dùng định dạng ký tự.
        /// </summary>
        public static bool ApplyFix(dynamic doc, dynamic rng, Fix f, bool alone, string eqMode, bool stacked)
        {
            bool fallback = false;
            if (!f.NeedsFormat) { rng.Text = f.Plain; return false; }
            MNode node = f.GetNode();
            bool useEq = node != null && (alone || Omml.IsComplex(f.Node, stacked));
            bool done = false;
            if (useEq) done = WordEq.Insert(doc, rng, node, alone, stacked);
            if (!done)
            {
                fallback = useEq;
                ApplyRunsTo(doc, rng, f, eqMode == "glyph");
            }
            return fallback;
        }

        const string GlyphFont = "Cambria Math";

        /// <summary>Đổi một đoạn chỉ số sang ký tự chỉ số nhỏ, thấp (Unicode). Null nếu có ký tự không có sẵn dạng chỉ số.</summary>
        public static string GlyphOf(string t, int kind)
        {
            const string sub = "₀₁₂₃₄₅₆₇₈₉", sup = "⁰¹²³⁴⁵⁶⁷⁸⁹";
            var sb = new StringBuilder();
            foreach (char c in t)
            {
                if (c >= '0' && c <= '9') sb.Append(kind == 1 ? sub[c - '0'] : sup[c - '0']);
                else if (c == '+') sb.Append(kind == 1 ? '₊' : '⁺');
                else if (c == '-' || c == '−') sb.Append(kind == 1 ? '₋' : '⁻');
                else if (c == '=') sb.Append(kind == 1 ? '₌' : '⁼');
                else if (c == '(') sb.Append(kind == 1 ? '₍' : '⁽');
                else if (c == ')') sb.Append(kind == 1 ? '₎' : '⁾');
                else if (kind == 2 && c == 'n') sb.Append('ⁿ');
                else if (kind == 2 && c == 'i') sb.Append('ⁱ');
                else return null;
            }
            return sb.ToString();
        }

        void ApplyRuns(dynamic rng, Fix f, bool glyph) { ApplyRunsTo(doc, rng, f, glyph); }

        /// <summary>Thay chữ trong rng bằng f.Plain (hoặc ký tự chỉ số nếu glyph) rồi áp dụng chỉ số trên, dưới, nghiêng, thẳng theo từng đoạn.</summary>
        public static void ApplyRunsTo(dynamic doc, dynamic rng, Fix f, bool glyph)
        {
            var parts = new List<string>(); var isGlyph = new List<bool>();
            foreach (Run r in f.Runs)
            {
                string g = glyph && (r.Kind == 1 || r.Kind == 2) ? GlyphOf(r.T, r.Kind) : null;
                parts.Add(g ?? r.T); isGlyph.Add(g != null);
            }
            rng.Text = string.Concat(parts.ToArray());
            int rs = (int)rng.Start;
            try { rng.Font.Subscript = 0; rng.Font.Superscript = 0; } catch { }
            if (f.IsChem) { try { rng.Font.Italic = 0; } catch { } }       // công thức hóa học: mọi chữ cái đứng thẳng
            int off = 0;
            for (int i = 0; i < f.Runs.Count; i++)
            {
                Run r = f.Runs[i]; int len = parts[i].Length;
                if (r.Kind != 0 && len > 0)
                {
                    dynamic sub = doc.Range(rs + off, rs + off + len);
                    if (r.Kind == 1 || r.Kind == 2)
                    {
                        try { sub.Font.Italic = 0; } catch { }               // chỉ số đứng thẳng
                        if (isGlyph[i])
                        {
                            sub.Font.Name = GlyphFont;
                            try { sub.Font.NameOther = GlyphFont; } catch { }
                        }
                        else if (r.Kind == 1) sub.Font.Subscript = -1;
                        else sub.Font.Superscript = -1;
                    }
                    else if (r.Kind == 3) sub.Font.Italic = -1;
                }
                off += len;
            }
        }

        public string FixItem(int idx, int fixIdx, AppSettings st)
        {
            Item it = Get(idx); CheckDoc();
            if (it.Done) throw new UserMsg("Mục này đã được xử lý.");
            Issue iss = it.H.Issue;
            if (iss.Fixes.Count == 0) throw new UserMsg("Mục này chỉ để xem, không có cách sửa tự động. Hãy sửa tay trong văn bản rồi bấm dấu ✖ để bỏ khỏi danh sách.");
            if (fixIdx < 0 || fixIdx >= iss.Fixes.Count) fixIdx = 0;
            Fix f = iss.Fixes[fixIdx];
            bool fb = false;
            WordTables.Fast((object)app, "Chuẩn hóa: sửa một mục", delegate { fb = ApplyOne(it, f, st.EqMode, st.EqStacked); });
            return "Đã sửa \"" + (iss.Disp ?? iss.Before) + "\" thành \"" + f.Show + "\"." + (it.Alone && !fb ? " Công thức nằm riêng một dòng nên đã dựng bằng Equation của Word." : "")
                + (fb ? " Không tạo được công thức Word nên đã dùng chỉ số trên, dưới bằng định dạng ký tự." : "");
        }

        public string FixAll(AppSettings st)
        {
            CheckDoc();
            var todo = items.Select((it, i) => new { it, i }).Where(x => !x.it.Done && x.it.H.Issue.Sev != "sus" && x.it.H.Issue.Fixes.Count > 0).ToList();
            int skippedSus = items.Count(x => !x.Done && x.H.Issue.Sev == "sus");
            if (todo.Count == 0)
                return skippedSus > 0 ? "Chỉ còn " + skippedSus + " mục nghi ngờ, cần bạn xem và quyết định từng mục." : "Không còn mục nào để sửa.";
            int ok = 0, fail = 0, fb = 0; string firstErr = "";
            WordTables.Fast((object)app, "Chuẩn hóa: sửa tất cả", delegate
            {
                ClearMarkCore();
                foreach (var x in todo.OrderByDescending(t => t.it.Start))     // từ cuối lên đầu để vị trí không lệch
                {
                    try { if (ApplyOne(x.it, x.it.H.Issue.Fixes[0], st.EqMode, st.EqStacked)) fb++; ok++; }
                    catch (Exception ex) { fail++; if (firstErr.Length == 0) firstErr = ex.Message; }
                }
            });
            return "Đã sửa " + ok + " mục." + (fail > 0 ? " " + fail + " mục không sửa được: " + firstErr : "")
                + (skippedSus > 0 ? " Còn " + skippedSus + " mục nghi ngờ chưa sửa, cần bạn xem từng mục." : "")
                + (fb > 0 ? " " + fb + " mục dùng chỉ số trên, dưới thay cho công thức Word." : "");
        }

        // ------------------------------------------------------------------ sửa nhanh tại con trỏ hoặc phần đang chọn
        static readonly Regex TokenCut = new Regex(@"[\s]");

        /// <summary>
        /// Nút Chuẩn hóa trên Ribbon: sửa ngay chữ tại con trỏ (hoặc phần đang bôi đen) về dạng chuẩn, không cần quét cả văn bản.
        /// Dùng phương án đầu tiên của mọi mục tìm thấy, kể cả mục nghi ngờ. Không tìm thấy lỗi quen thuộc mà chữ có _ ^ { } \ thì đọc như LaTeX.
        /// </summary>
        public string QuickFix(AppSettings st)
        {
            try { if ((int)app.Documents.Count == 0) throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            catch (UserMsg) { throw; }
            catch { throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            doc = app.ActiveDocument;
            dynamic selRange = app.Selection.Range;
            int a = (int)selRange.Start, b = (int)selRange.End;
            var o = new ReviewOpts();
            o.NumStyle = st.RvNumStyle; o.NumDec = st.NumDec; o.NumThou = st.NumThou; o.NumAddDec = st.NumAddDecToInt;
            int ok = 0, fb = 0, already = 0; string firstErr = ""; var shown = new List<string>();
            WordTables.Fast((object)app, "Chuẩn hóa: sửa nhanh", delegate
            {
                var todo = new List<Item>();
                foreach (dynamic p in (IEnumerable)selRange.Paragraphs)
                {
                    dynamic pr = p.Range;
                    string text = (string)pr.Text;
                    if (text == null || text.Length < 1) continue;
                    int ps = (int)pr.Start;
                    bool inTable = false, heading = false;
                    try { inTable = (bool)pr.Information(12); } catch { }
                    try { heading = (int)p.OutlineLevel < 10; } catch { }
                    var ctx = new NumCtx { InTable = inTable, Heading = heading,
                        Language = (left, right) => LanguageAt((object)doc, ps + left, ps + right) };
                    if (inTable) ctx.NumericCell = NumericCellRe.IsMatch(CsvUtil.CellText(text));
                    List<Hit> hits = Detector.Detect(text, o, delegate { return ctx; });
                    string core = text.TrimEnd('\r', '\n', '\a', ' ');
                    bool any = false;
                    foreach (Hit h in hits)
                    {
                        int s = ps + h.S, e = ps + h.E;
                        bool touch = a == b ? (s <= a && a <= e) : (s < b && e > a);
                        if (!touch || h.Issue.Fixes.Count == 0) continue;
                        dynamic rng = doc.Range(s, e);
                        string actual = ""; try { actual = (string)rng.Text; } catch { }
                        if (actual != h.Issue.Before) continue;
                        any = true;
                        if (h.Issue.Fixes[0].NeedsFormat)
                        {
                            bool done = false;
                            try { if ((int)rng.OMaths.Count > 0) done = true; } catch { }
                            if (!done) { try { if ((int)rng.Font.Subscript != 0 || (int)rng.Font.Superscript != 0) done = true; } catch { } }
                            if (done) { already++; continue; }
                        }
                        bool alone = h.Issue.Cat != "num" && h.Issue.Fixes[0].NeedsFormat && !inTable && h.E <= core.Length
                            && core.Substring(0, h.S).Trim().Length == 0 && AloneTail.IsMatch(core.Substring(h.E));
                        todo.Add(new Item { H = h, Rng = rng, Before = h.Issue.Before, Start = s, Alone = alone });
                    }
                    if (!any && a == b && ps <= a && a <= ps + text.Length)
                    {
                        // không có lỗi quen thuộc: đọc chữ quanh con trỏ như mã LaTeX (R_cyl, k_{a1}, v_2, x^{n+1})
                        int c = a - ps; if (c > text.Length) c = text.Length;
                        int ts = c, te = c;
                        while (ts > 0 && !char.IsWhiteSpace(text[ts - 1])) ts--;
                        while (te < text.Length && !char.IsWhiteSpace(text[te]) && text[te] != '\r' && text[te] != '\a') te++;
                        string tok = text.Substring(ts, te - ts);
                        string trimmed = tok.TrimEnd('.', ',', ';', ':');
                        te = ts + trimmed.Length; tok = trimmed;
                        if (tok.IndexOfAny(new[] { '_', '^', '{', '}', '\\' }) >= 0)
                        {
                            MNode node = MathParser.Parse(tok);
                            Fix nf = Fix.FromNode("Sửa", node);
                            if (nf.Plain != tok || nf.NeedsFormat)
                            {
                                var iss = new Issue { Cat = "sym", Sev = "sus", Before = tok, Key = "q|" + tok };
                                iss.Fixes.Add(nf);
                                dynamic rng = doc.Range(ps + ts, ps + te);
                                todo.Add(new Item { H = new Hit { S = ts, E = te, Issue = iss }, Rng = rng, Before = tok, Start = ps + ts, Alone = false });
                            }
                        }
                    }
                }
                foreach (Item it in todo.OrderByDescending(t => t.Start))
                {
                    try { if (ApplyOne(it, it.H.Issue.Fixes[0], st.EqMode, st.EqStacked)) fb++; ok++; if (shown.Count < 3) shown.Add(it.Before); }
                    catch (Exception ex) { if (firstErr.Length == 0) firstErr = ex.Message; }
                }
            });
            if (ok == 0)
                return already > 0 ? "Chữ này đã được định dạng chuẩn rồi."
                    : firstErr.Length > 0 ? firstErr
                    : a == b ? "Không thấy lỗi cần chuẩn hóa tại con trỏ. Hãy đặt con trỏ vào hoặc bôi đen từ, công thức cần sửa (ví dụ R_cyl, R_{cyl}, t_req)."
                    : "Không thấy lỗi cần chuẩn hóa trong phần đang bôi đen.";
            return "Đã chuẩn hóa " + ok + " chỗ (" + string.Join(", ", shown.ToArray()) + (ok > shown.Count ? ", ..." : "") + "): chữ biến in nghiêng, chỉ số đứng thẳng, bỏ các dấu thừa."
                + (fb > 0 ? " " + fb + " chỗ không dựng được công thức Word nên dùng chỉ số trên, dưới." : "");
        }

        /// <summary>Danh sách các mục chưa xử lý (sau khi sửa hoặc bỏ qua), để bảng điều khiển vẽ lại.</summary>
        public RvResult Pending()
        {
            var r = new RvResult();
            for (int i = 0; i < items.Count; i++) if (!items[i].Done) r.Rows.Add(ToRow(i));
            return r;
        }
    }

    /// <summary>Chèn công thức đã soạn vào Word tại con trỏ hoặc thay cho phần đang bôi đen.</summary>
    public sealed class WordCompose
    {
        readonly dynamic app;
        public WordCompose(object application) { app = application; }

        /// <summary>
        /// Xác định công thức chèn riêng một dòng (Equation block) hay cùng dòng (inline):
        /// Luôn cùng dòng, TRỪ KHI đoạn văn chứa con trỏ hoàn toàn trống (không có ký tự nào ngoài dấu hết đoạn)
        /// hoặc vùng bôi đen là cả một đoạn (khi đó riêng một dòng).
        /// </summary>
        internal static bool IsAlone(dynamic rng, bool replaceSel)
        {
            try
            {
                int start = (int)rng.Start;
                int end = (int)rng.End;
                bool isSelection = replaceSel && (start != end);

                var paras = new List<dynamic>();
                foreach (dynamic p in (IEnumerable)rng.Paragraphs) paras.Add(p);
                if (paras.Count == 0) return false;

                if (!isSelection)
                {
                    // Con trỏ: kiểm tra đoạn văn chứa con trỏ có hoàn toàn trống (không có ký tự nào ngoài dấu hết đoạn) không
                    dynamic p = paras[0];
                    string text = (string)p.Range.Text;
                    string core = (text ?? "").TrimEnd('\r', '\n', '\a');
                    return core.Trim().Length == 0;
                }
                else
                {
                    // Vùng bôi đen: chỉ riêng một dòng khi bôi đen là đúng cả một đoạn
                    if (paras.Count == 1)
                    {
                        dynamic pr = paras[0].Range;
                        int ps = (int)pr.Start;
                        int pe = (int)pr.End;
                        string text = (string)pr.Text;
                        string core = (text ?? "").TrimEnd('\r', '\n', '\a');
                        if (core.Trim().Length == 0) return true;
                        if (start == ps && (end == pe || end == ps + core.Length || end == pe - 1)) return true;
                    }
                    return false;
                }
            }
            catch { return false; }
        }

        public string Insert(ComposeResult r, bool replaceSel, string eqMode, bool stacked)
        {
            return Insert(r, null, replaceSel, eqMode, stacked);
        }

        public string Insert(ComposeResult r, bool alone, bool replaceSel, string eqMode, bool stacked)
        {
            return Insert(r, (bool?)alone, replaceSel, eqMode, stacked);
        }

        public string InsertPreview(ComposeResult formula, bool display)
        {
            MNode node = formula != null && formula.Ok && formula.Fix != null ? formula.Fix.GetNode() : null;
            if (node == null) throw new UserMsg("Chưa có công thức hợp lệ để chèn.");
            if ((int)app.Documents.Count == 0) throw new UserMsg("Chưa có tài liệu Word nào đang mở.");
            WordTables.Fast((object)app, "Chèn phương trình", delegate
            {
                dynamic doc = app.ActiveDocument;
                dynamic range = app.Selection.Range;
                range.Collapse(0);
                if (!WordEq.Insert(doc, range, node, display, display))
                    throw new UserMsg("Word không chèn được Equation theo kiểu đã chọn.");
                dynamic equation = range.OMaths.Item(1);
                if (display && (int)equation.Type != 0) equation.Type = 0; // Word có thể ép về inline giữa đoạn.
                range = equation.Range;
                range.Collapse(0);
                range.Select();
            });
            WordEq.FocusWord(app);
            return display ? "Đã chèn phương trình riêng dòng." : "Đã chèn phương trình cùng dòng.";
        }

        string Insert(ComposeResult r, bool? aloneOpt, bool replaceSel, string eqMode, bool stacked)
        {
            if (r == null || !r.Ok || r.Fix == null) throw new UserMsg("Chưa có công thức hợp lệ để chèn.");
            try { if ((int)app.Documents.Count == 0) throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            catch (UserMsg) { throw; }
            catch { throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            string how = "";
            WordTables.Fast((object)app, "Soạn công thức: chèn", delegate
            {
                dynamic doc = app.ActiveDocument;
                dynamic rng = app.Selection.Range;
                bool alone = aloneOpt.HasValue ? aloneOpt.Value : IsAlone(rng, replaceSel);
                if (!replaceSel || (int)rng.Start == (int)rng.End) rng.Collapse(0);      // wdCollapseEnd: chèn sau con trỏ
                Fix f = r.Fix;
                MNode node = f.GetNode();
                bool useEq = node != null && f.NeedsFormat && (alone || r.Complex || Omml.IsComplex(f.Node, stacked));
                bool done = false;
                if (useEq)
                {
                    done = WordEq.Insert(doc, rng, node, alone, stacked);
                    if (done) how = alone ? "Equation riêng một dòng" : "Equation cùng dòng";
                }
                if (!done)
                {
                    if (f.NeedsFormat) { WordReview.ApplyRunsTo(doc, rng, f, eqMode == "glyph"); how = eqMode == "glyph" ? "ký tự chỉ số nhỏ, thấp (font Cambria)" : "định dạng chỉ số trên, dưới"; }
                    else { rng.Text = f.Plain; how = "chữ thường"; }
                    if (useEq) how += " (không tạo được Equation nên dùng cách này)";
                }
                try { rng.Collapse(0); rng.Select(); } catch { }                         // đặt con trỏ sau công thức để thấy nó nhấp nháy
            });
            WordEq.FocusWord(app);
            return "Đã chèn công thức \"" + r.Show + "\" bằng " + how + ".";
        }
    }
}

namespace TroLyKyThuat
{
    /// <summary>Chèn công thức vào Word bằng OMML (Range.InsertXML) và đưa con trỏ về Word.</summary>
    public static class WordEq
    {
        public static bool Insert(dynamic doc, dynamic rng, MNode node, bool display, bool stacked)
        {
            try
            {
                string xml = Omml.Package(Omml.Build(node, stacked), display);
                rng.InsertXML(xml);
                try
                {
                    // Word có thể thêm một dấu xuống dòng ngay sau phần vừa chèn: bỏ nó để không tách đoạn
                    string t = (string)rng.Text;
                    if (t.Length > 0 && t[t.Length - 1] == '\r') { int e = (int)rng.End; doc.Range(e - 1, e).Delete(); }
                }
                catch { }
                bool ok = false;
                try { ok = (int)rng.OMaths.Count > 0; } catch { }
                return ok;
            }
            catch (Exception ex) { Log.W("InsertXML: " + ex.Message); return false; }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr h);

        /// <summary>Trả tiêu điểm bàn phím về cửa sổ Word để con trỏ soạn thảo hiện và nhấp nháy.</summary>
        public static void FocusWord(dynamic app)
        {
            try { app.Activate(); } catch { }
            try { int h = (int)app.ActiveWindow.Hwnd; if (h != 0) SetForegroundWindow((IntPtr)h); } catch { }
            try { app.ActiveWindow.Activate(); } catch { }
        }
    }
}
