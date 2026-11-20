// Trợ lý Kỹ thuật - Rev0.9.5: tra thuật ngữ và gợi ý từ đồng nghĩa tại con trỏ hoặc phần bôi đen trong Word.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    public sealed class SynResult
    {
        public int Start, End;
        public string Text, SentenceContext;
        internal object Document;
    }

    public sealed class RewriteParaSlice
    {
        public int Start, End;
        public string Text;
        public string FontName;
        public float? FontSize;
        public int? Bold, Italic, Color;
    }

    public sealed class RewriteSnapshot
    {
        public int Start, End, WordCount;
        public string Text = "";
        internal object Document;
        public List<RewriteParaSlice> Slices = new List<RewriteParaSlice>();
        internal int DocumentEnd;
    }

    public sealed class WordText
    {
        readonly dynamic app;
        RewriteSnapshot applied;
        string appliedText;
        public WordText(object application) { app = application; }

        sealed class Cand { public string Text; public int Start, End; }
        static readonly Regex TokenRe = new Regex(@"[\p{L}\p{N}]+(?:[-'’][\p{L}\p{N}]+)*", RegexOptions.CultureInvariant);

        dynamic Doc()
        {
            try { if ((int)app.Documents.Count == 0) throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            catch (UserMsg) { throw; }
            catch { throw new UserMsg("Chưa có tài liệu Word nào đang mở."); }
            return app.ActiveDocument;
        }

        /// <summary>Cụm cần tra: chữ đang bôi đen, hoặc các cụm 4 đến 1 từ quanh con trỏ (dài trước, ngắn sau).</summary>
        List<Cand> Candidates()
        {
            dynamic doc = Doc();
            dynamic sel = app.Selection; dynamic rng = sel.Range;
            int start = (int)rng.Start, end = (int)rng.End;
            var res = new List<Cand>();
            if (end > start)
            {
                string t = (string)rng.Text ?? "";
                int a = 0, b = t.Length;
                while (a < b && !char.IsLetterOrDigit(t[a])) a++;
                while (b > a && !char.IsLetterOrDigit(t[b - 1])) b--;
                if (b <= a) throw new UserMsg("Phần bôi đen không có chữ nào. Hãy bôi đen một từ hoặc cụm từ.");
                res.Add(new Cand { Text = t.Substring(a, b - a), Start = start + a, End = start + b });
                return res;
            }
            dynamic pr = null;
            foreach (dynamic p in (IEnumerable)rng.Paragraphs) { pr = p.Range; break; }
            if (pr == null) throw new UserMsg("Hãy đặt con trỏ trong hoặc sát một từ, hoặc bôi đen từ cần tra.");
            string text = (string)pr.Text ?? ""; int baseAt = (int)pr.Start;
            int caret = start - baseAt;
            var toks = TokenRe.Matches(text).Cast<Match>().ToList();
            int k = -1;
            for (int i = 0; i < toks.Count; i++) if (toks[i].Index <= caret && caret <= toks[i].Index + toks[i].Length) { k = i; break; }
            if (k < 0) throw new UserMsg("Con trỏ không nằm trong hoặc sát từ nào. Đặt con trỏ trong từ hoặc bôi đen từ rồi bấm lại.");
            for (int len = 4; len >= 1; len--)
                for (int first = k - len + 1; first <= k; first++)
                {
                    int last = first + len - 1;
                    if (first < 0 || last >= toks.Count) continue;
                    bool ok = true;
                    for (int j = first; j < last && ok; j++)
                    {
                        string gap = text.Substring(toks[j].Index + toks[j].Length, toks[j + 1].Index - (toks[j].Index + toks[j].Length));
                        ok = gap.Length > 0 && gap.All(c => c == ' ' || c == ' ');
                    }
                    if (!ok) continue;
                    int s = toks[first].Index, e = toks[last].Index + toks[last].Length;
                    res.Add(new Cand { Text = text.Substring(s, e - s), Start = baseAt + s, End = baseAt + e });
                }
            return res;
        }

        internal static void RequireMainStory(object range)
        {
            // Word indexes headers/footnotes separately from the main body.
            dynamic wordRange = range;
            int story = 1;
            try { story = (int)wordRange.StoryType; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            if (story != 1) throw new UserMsg("Hãy chọn văn bản trong phần nội dung chính của tài liệu.");
        }

        public SynResult Suggest()
        {
            dynamic active = Doc();
            RequireMainStory((object)app.Selection.Range);
            List<Cand> candidates = Candidates();
            Cand candidate = candidates[candidates.Count - 1];
            if (candidate.Text.Length > 200) throw new UserMsg("Hãy chọn một từ hoặc cụm ngắn.");
            string context = ReadCurrentParaText();
            if (context.Length > 1600)
            {
                int where = context.IndexOf(candidate.Text, StringComparison.Ordinal);
                int start = Math.Max(0, where - 700);
                context = context.Substring(start, Math.Min(1600, context.Length - start));
            }
            return new SynResult { Document = (object)active, Start = candidate.Start, End = candidate.End,
                Text = candidate.Text, SentenceContext = context };
        }

        string ReadCurrentParaText()
        {
            try
            {
                dynamic rng = app.Selection.Range;
                foreach (dynamic p in (IEnumerable)rng.Paragraphs)
                    return ((string)p.Range.Text ?? "").TrimEnd('\r', '\n', '\a');
            }
            catch { }
            return "";
        }

        /// <summary>Thay cụm đã gợi ý bằng alt, nếu chữ tại vị trí đó chưa bị đổi, đồng thời giữ nguyên định dạng ký tự gốc.</summary>
        public string Replace(SynResult r, string alt)
        {
            if (r == null || string.IsNullOrEmpty(alt)) throw new UserMsg("Chưa chọn từ để thay. Bấm Gợi ý từ trước.");
            dynamic doc = Doc();
            if (!object.ReferenceEquals(r.Document, (object)doc)) throw new UserMsg("Tài liệu đã đổi. Hãy yêu cầu gợi ý lại.");
            dynamic rng = doc.Range(r.Start, r.End);
            string actual = ""; try { actual = (string)rng.Text; } catch { }
            if (actual != r.Text) throw new UserMsg("Chữ tại vị trí \"" + r.Text + "\" đã đổi. Hãy bấm Gợi ý từ lại.");

            string fontName = null;
            float? fontSize = null;
            int? bold = null, italic = null, underline = null, color = null, sub = null, sup = null;
            try
            {
                dynamic f = rng.Font;
                object oName = f.Name; if (oName is string && ((string)oName).Length > 0) fontName = (string)oName;
                object oSize = f.Size; if (oSize is float || oSize is double || oSize is int) fontSize = Convert.ToSingle(oSize);
                object oBold = f.Bold; if (oBold is int) bold = (int)oBold;
                object oItalic = f.Italic; if (oItalic is int) italic = (int)oItalic;
                object oUnder = f.Underline; if (oUnder is int) underline = (int)oUnder;
                object oColor = f.Color; if (oColor is int) color = (int)oColor;
                object oSub = f.Subscript; if (oSub is int) sub = (int)oSub;
                object oSup = f.Superscript; if (oSup is int) sup = (int)oSup;
            }
            catch { }

            WordTables.Fast((object)app, "Văn phong: thay từ", delegate
            {
                rng.Text = alt;
                try
                {
                    dynamic newRng = doc.Range(r.Start, r.Start + alt.Length);
                    dynamic nf = newRng.Font;
                    if (fontName != null) nf.Name = fontName;
                    if (fontSize.HasValue && fontSize.Value > 0 && fontSize.Value < 1000) nf.Size = fontSize.Value;
                    if (bold.HasValue) nf.Bold = bold.Value;
                    if (italic.HasValue) nf.Italic = italic.Value;
                    if (underline.HasValue) nf.Underline = underline.Value;
                    if (color.HasValue) nf.Color = color.Value;
                    if (sub.HasValue) nf.Subscript = sub.Value;
                    if (sup.HasValue) nf.Superscript = sup.Value;
                }
                catch { }
            });
            int at = r.Start + alt.Length;
            try { app.Selection.SetRange(at, at); } catch { }
            return "Đã thay \"" + r.Text + "\" bằng \"" + alt + "\".";
        }

        /// <summary>Đọc đoạn văn đang bôi đen (trừ dấu ngắt đoạn cuối), số từ, từng lát đoạn và ngữ cảnh toàn tài liệu cho tab Viết lại.</summary>
        public RewriteSnapshot GetRewriteSnapshot()
        {
            dynamic doc = Doc();
            dynamic sel = app.Selection.Range;
            RequireMainStory((object)sel);
            int a = (int)sel.Start, b = (int)sel.End;
            if (b <= a) return new RewriteSnapshot { Start = a, End = a, WordCount = 0, Text = "" };

            RequireEditableText((object)sel);
            string rawSel = "";
            try { rawSel = (string)doc.Range(a, b).Text ?? ""; }
            catch { try { rawSel = (string)sel.Text ?? ""; } catch { } }
            while (b > a && rawSel.Length > 0 && (rawSel[rawSel.Length - 1] == '\r' || rawSel[rawSel.Length - 1] == '\n' || rawSel[rawSel.Length - 1] == '\a'))
            {
                rawSel = rawSel.Substring(0, rawSel.Length - 1);
                b--;
            }

            if (rawSel.IndexOf('\a') >= 0 || rawSel.Length != b - a)
                throw new UserMsg("Hãy chọn văn bản thuần trong cùng một ô hoặc đoạn.");

            var snap = new RewriteSnapshot
            {
                Document = (object)doc,
                DocumentEnd = (int)doc.Content.End,
                Start = a,
                End = b,
                Text = rawSel.Replace('\r', '\n').TrimEnd('\n'),
                WordCount = AiClient.CountWords(rawSel)
            };

            // Ghi nhận định dạng ký tự trên từng đoạn thuộc vùng đang chọn để giữ nguyên kiểu đoạn, font, màu...
            try
            {
                foreach (dynamic p in (IEnumerable)doc.Range(a, b).Paragraphs)
                {
                    dynamic pr = p.Range;
                    int ps = Math.Max(a, (int)pr.Start);
                    int pe = Math.Min(b, (int)pr.End);
                    string pTxt = "";
                    try { pTxt = ((string)doc.Range(ps, pe).Text ?? "").TrimEnd('\r', '\n', '\a'); } catch { }
                    if (pe <= ps) continue;
                    var sl = new RewriteParaSlice { Start = ps, End = pe, Text = pTxt };
                    try
                    {
                        dynamic f = doc.Range(ps, pe).Font;
                        object oName = f.Name; if (oName is string && ((string)oName).Length > 0) sl.FontName = (string)oName;
                        object oSize = f.Size; if (oSize is float || oSize is double || oSize is int) sl.FontSize = Convert.ToSingle(oSize);
                        object oBold = f.Bold; if (oBold is int && (int)oBold != 9999999) sl.Bold = (int)oBold;
                        object oItalic = f.Italic; if (oItalic is int && (int)oItalic != 9999999) sl.Italic = (int)oItalic;
                        object oColor = f.Color; if (oColor is int && (int)oColor != 9999999) sl.Color = (int)oColor;
                    }
                    catch { }
                    snap.Slices.Add(sl);
                }
            }
            catch { }

            return snap;
        }

        /// <summary>Áp dụng văn bản AI vào vùng đã chọn và giữ ngắt đoạn trong một bước Undo.</summary>
        public string ApplyRewrite(RewriteSnapshot snap, string rewritten, string translationTarget = null)
        {
            if (snap == null || snap.End <= snap.Start) throw new UserMsg("Chưa có vùng văn bản hợp lệ để áp dụng.");
            if (string.IsNullOrWhiteSpace(rewritten)) throw new UserMsg("Nội dung viết lại đang trống.");
            dynamic doc = Doc();

            if (!object.ReferenceEquals(snap.Document, (object)doc))
                throw new UserMsg("Tài liệu đã đổi. Hãy chạy Công cụ AI lại.");
            dynamic target = doc.Range(snap.Start, snap.End);
            RequireEditableText((object)target);
            string actual = ((string)target.Text ?? "").Replace('\r', '\n');
            if ((int)doc.Content.End != snap.DocumentEnd || actual != snap.Text)
                throw new UserMsg("Văn bản đã thay đổi. Hãy chạy Công cụ AI lại.");

            string[] output = rewritten.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            string[] input = snap.Text.Split('\n');
            if (input.Length != output.Length || output.Any(line => line.Any(c => char.IsControl(c) && c != '\t')))
                throw new UserMsg("Bản xem trước phải giữ số đoạn và các dấu ngắt đoạn của vùng đã chọn.");
            for (int i = 0; i < input.Length; i++)
                if ((input[i].Length == 0) != (output[i].Length == 0))
                    throw new UserMsg("Bản xem trước phải giữ nguyên vị trí đoạn trống.");
            string normalized = string.Join("\n", output);
            if (TranslationNumbers.Supports(translationTarget))
            {
                string localized;
                if (!TranslationNumbers.TryNormalize(snap.Text, normalized, translationTarget, out localized))
                    throw new UserMsg("Bản xem trước không khớp số liệu, đơn vị, số hiệu hoặc trích dẫn của bản gốc.");
                normalized = localized;
                output = normalized.Split('\n');
            }
            else if (!AiReview.CanReplace(snap.Text, normalized, true))
                throw new UserMsg("Bản xem trước đã đổi số liệu hoặc trích dẫn. Hãy kiểm tra và chỉnh lại.");

            if (normalized == snap.Text) throw new UserMsg("Bản xem trước giống vùng đã chọn.");
            applied = null;
            var offsets = new int[input.Length];
            for (int i = 1; i < offsets.Length; i++) offsets[i] = offsets[i - 1] + input[i - 1].Length + 1;
            WordTables.Fast((object)app, "Áp dụng kết quả AI", delegate
            {
                for (int i = input.Length - 1; i >= 0; i--)
                {
                    if (input[i] == output[i]) continue;
                    int start = snap.Start + offsets[i];
                    dynamic range = doc.Range(start, start + input[i].Length);
                    range.Text = output[i];
                    RewriteParaSlice format = snap.Slices.FirstOrDefault(slice => slice.Start == start);
                    if (format != null) ApplySliceFont(doc.Range(start, start + output[i].Length), format);
                }
            });
            applied = snap;
            appliedText = normalized;
            return "Đã áp dụng vào vùng đã bôi đen. Có thể hoàn tác bằng Ctrl+Z.";
        }

        internal static void RequireEditableText(object range)
        {
            dynamic value = range;
            try { if ((int)value.Fields.Count > 0) throw new UserMsg("Vùng có trường Word. Hãy chọn phần văn bản thuần."); }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            try { if ((int)value.OMaths.Count > 0) throw new UserMsg("Vùng có công thức Word. Hãy chọn phần văn bản thuần."); }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            try { if ((int)value.ContentControls.Count > 0) throw new UserMsg("Vùng có điều khiển nội dung của Word."); }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            try { if ((int)value.Hyperlinks.Count > 0) throw new UserMsg("Vùng có liên kết Word. Hãy chọn phần văn bản thuần."); }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
        }

        /// <summary>Hoàn tác một bước thay đổi vừa áp dụng trong Word.</summary>
        public string UndoLast()
        {
            dynamic doc = Doc();
            if (applied == null || !object.ReferenceEquals(applied.Document, (object)doc) ||
                (int)doc.Content.End != applied.DocumentEnd - applied.Text.Length + appliedText.Length ||
                ((string)doc.Range(applied.Start, applied.Start + appliedText.Length).Text ?? "").Replace('\r', '\n') != appliedText)
                throw new UserMsg("Không xác định được thao tác AI gần nhất trong tài liệu này.");
            try
            {
                if (!(bool)doc.Undo(1)) throw new UserMsg("Word không có thao tác để hoàn tác.");
                applied = null;
                return "Đã hoàn tác thao tác gần nhất trong Word.";
            }
            catch (Exception ex) { throw new UserMsg("Không hoàn tác được: " + ex.Message); }
        }

        static void ApplySliceFont(dynamic rng, RewriteParaSlice sl)
        {
            if (sl == null) return;
            try
            {
                dynamic nf = rng.Font;
                if (sl.FontName != null) nf.Name = sl.FontName;
                if (sl.FontSize.HasValue && sl.FontSize.Value > 0 && sl.FontSize.Value < 1000) nf.Size = sl.FontSize.Value;
                if (sl.Bold.HasValue) nf.Bold = sl.Bold.Value;
                if (sl.Italic.HasValue) nf.Italic = sl.Italic.Value;
                if (sl.Color.HasValue) nf.Color = sl.Color.Value;
            }
            catch { }
        }
    }
}
