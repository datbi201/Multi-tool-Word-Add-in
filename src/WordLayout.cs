// Trợ lý Kỹ thuật - Rev0.9.3: module Dàn trang (chuẩn hóa trang, chuẩn hóa đoạn văn, thêm trang ngang/dọc xen giữa văn bản).
using System;
using System.Collections.Generic;
using System.Globalization;

namespace TroLyKyThuat
{
    public sealed class WordLayout
    {
        readonly dynamic app;

        public WordLayout(object application) { app = application; }

        bool HasDocument() { try { return (int)app.Documents.Count > 0; } catch { return false; } }

        public static double CmToPt(double cm) { return cm * 72.0 / 2.54; }

        /// <summary>Mã WdPaperSize của Word: A4 = 7, Letter = 2, A3 = 6, Legal = 4.</summary>
        public static int PaperSizeEnum(string paper)
        {
            if (paper == "Letter") return 2;
            if (paper == "A3") return 6;
            if (paper == "Legal") return 4;
            return 7; // A4
        }

        void ApplyMargins(dynamic ps, AppSettings st, int orient)
        {
            ps.PaperSize = PaperSizeEnum(st.PagePaper);
            if (orient == 1)
            {
                ps.TopMargin = (float)CmToPt(st.PgLandTop);
                ps.BottomMargin = (float)CmToPt(st.PgLandBot);
                ps.LeftMargin = (float)CmToPt(st.PgLandLeft);
                ps.RightMargin = (float)CmToPt(st.PgLandRight);
            }
            else
            {
                ps.TopMargin = (float)CmToPt(st.PgPortTop);
                ps.BottomMargin = (float)CmToPt(st.PgPortBot);
                ps.LeftMargin = (float)CmToPt(st.PgPortLeft);
                ps.RightMargin = (float)CmToPt(st.PgPortRight);
            }
            ps.HeaderDistance = (float)CmToPt(st.PgHeaderDist);
            ps.FooterDistance = (float)CmToPt(st.PgFooterDist);
        }

        /// <summary>1. Chuẩn hóa dàn trang: đặt khổ giấy, lề trang dọc/ngang, đầu/chân trang cho mọi section theo Cài đặt.</summary>
        public string StandardizePageSetup(AppSettings st)
        {
            if (!HasDocument()) throw new UserMsg("Chưa có tài liệu Word nào đang mở.");
            dynamic doc = app.ActiveDocument;
            int n = (int)doc.Sections.Count;
            if (n == 0) return "Tài liệu không có phần (section) nào.";
            int port = 0, land = 0;
            WordTables.Fast((object)app, "Chuẩn hóa dàn trang", delegate
            {
                for (int i = 1; i <= n; i++)
                {
                    dynamic sec = doc.Sections.Item(i);
                    dynamic ps = sec.PageSetup;
                    int orient = 0;
                    try { orient = (int)ps.Orientation; } catch { orient = 0; }
                    ApplyMargins(ps, st, orient);
                    if (orient == 1) land++; else port++;
                }
            });
            return string.Format("Đã chuẩn hóa dàn trang cho {0} phần ({1} dọc, {2} ngang): giấy {3}, lề dọc {4}-{5}-{6}-{7} cm, lề ngang {8}-{9}-{10}-{11} cm.",
                n, port, land, st.PagePaper,
                Fmt(st.PgPortTop), Fmt(st.PgPortBot), Fmt(st.PgPortLeft), Fmt(st.PgPortRight),
                Fmt(st.PgLandTop), Fmt(st.PgLandBot), Fmt(st.PgLandLeft), Fmt(st.PgLandRight));
        }

        /// <summary>Hằng số chuyển đổi: 1 cm = 28,3464567 điểm (pt).</summary>
        public const double CmToPtFactor = 28.3464567;

        /// <summary>Chuẩn hóa trực tiếp đoạn tại đầu hoạt động của vùng chọn, căn đều hai bên.</summary>
        static int AlignCode(string a)
        {
            switch (a) { case "left": return 0; case "center": return 1; case "right": return 2; default: return 3; }   // wdAlignParagraph Left, Center, Right, Justify
        }

        static string AlignName(string a)
        {
            switch (a) { case "left": return "căn trái"; case "center": return "căn giữa"; case "right": return "căn phải"; default: return "căn đều hai bên"; }
        }

        static string StyleKey(dynamic style)
        {
            string name = null;
            try { name = (string)style.NameLocal; } catch { }
            return string.IsNullOrEmpty(name) ? "#" + System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode((object)style) : name;
        }

        /// <summary>Chuẩn hóa đoạn văn: cỡ chữ theo font, canh lề (Cài đặt), giãn dòng, khoảng cách đoạn, thụt dòng đầu hoặc thụt treo.
        /// Không bôi đen: chỉ đoạn tại con trỏ. Bôi đen nhiều đoạn: áp cho tất cả các đoạn đó cùng lúc (bỏ qua tiêu đề và đoạn trong ô bảng).</summary>
        public string StandardizeParagraphs(AppSettings st)
        {
            if (!HasDocument()) throw new UserMsg("Chưa có tài liệu Word nào đang mở.");
            dynamic selection = app.Selection;
            bool multi = false;
            try { multi = (int)selection.End != (int)selection.Start; } catch { multi = false; }
            var paras = new List<dynamic>();
            if (multi)
            {
                try
                {
                    dynamic ps = selection.Paragraphs;
                    int n = (int)ps.Count;
                    for (int i = 1; i <= n; i++) paras.Add(ps.Item(i));
                }
                catch { paras.Clear(); }
            }
            if (paras.Count == 0)
            {
                multi = false;
                dynamic cursor = selection.Range.Duplicate;
                cursor.Collapse((bool)selection.StartIsActive ? 1 : 0);
                paras.Add(cursor.Paragraphs.Item(1));
            }
            int alignCode = AlignCode(st.ParaAlign);
            int done = 0, skipped = 0, indented = 0;

            // Style bật AutomaticallyUpdate có thể biến định dạng trực tiếp thành định nghĩa style. Tạm tắt cho từng style gặp phải và khôi phục trong finally.
            var suspended = new Dictionary<string, dynamic>();
            foreach (dynamic para in paras)
            {
                try
                {
                    dynamic style = para.Range.Style;
                    string key = StyleKey(style);
                    if (!suspended.ContainsKey(key) && (bool)style.AutomaticallyUpdate) suspended[key] = style;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                catch (System.Runtime.InteropServices.COMException) { }
            }

            WordTables.Fast((object)app, multi ? "Chuẩn hóa các đoạn đã chọn" : "Chuẩn hóa đoạn tại con trỏ", delegate
            {
                foreach (KeyValuePair<string, dynamic> kv in suspended) kv.Value.AutomaticallyUpdate = false;
                try
                {
                    foreach (dynamic paragraph in paras)
                    {
                        dynamic range = paragraph.Range;
                        dynamic format = paragraph.Format;
                        bool inTable = false, heading = false, list = false;
                        try { inTable = (bool)range.Information(12); } catch { }
                        try { heading = (int)format.OutlineLevel < 10; } catch { }
                        try { list = (int)range.ListFormat.ListType != 0; } catch { }
                        if (multi && (inTable || heading)) { skipped++; continue; }

                        string name = "";
                        try { name = (string)range.Font.Name ?? ""; } catch { }
                        if (string.Equals(name, "Times New Roman", StringComparison.OrdinalIgnoreCase))
                            range.Font.Size = (float)st.ParaSizeTnr;
                        else if (string.Equals(name, "Arial", StringComparison.OrdinalIgnoreCase))
                            range.Font.Size = (float)st.ParaSizeArial;

                        format.Alignment = alignCode;
                        format.LineSpacingRule = 5; // wdLineSpaceMultiple
                        format.LineSpacing = (float)(st.ParaLineSpace * 12.0);
                        format.SpaceBefore = (float)st.ParaSpaceBefore;
                        format.SpaceAfter = (float)st.ParaSpaceAfter;

                        if (!inTable && !heading && !list)
                        {
                            if (st.ParaHangingCm > 0)
                            {
                                float indent = (float)CmToPt(st.ParaHangingCm);
                                format.LeftIndent = indent;
                                format.FirstLineIndent = -indent;
                            }
                            else format.FirstLineIndent = (float)CmToPt(st.ParaFirstLineCm);
                            indented++;
                        }
                        done++;
                    }
                }
                finally
                {
                    foreach (KeyValuePair<string, dynamic> kv in suspended) { try { kv.Value.AutomaticallyUpdate = true; } catch { } }
                }
            });
            string tail = AlignName(st.ParaAlign) + "; giãn dòng " + Fmt(st.ParaLineSpace) + ", trước " + Fmt(st.ParaSpaceBefore) + " pt, sau " + Fmt(st.ParaSpaceAfter) + " pt.";
            if (!multi) return "Đã chuẩn hóa đoạn tại con trỏ và " + tail;
            return "Đã chuẩn hóa " + done + " đoạn trong vùng chọn" + (skipped > 0 ? " (bỏ qua " + skipped + " tiêu đề hoặc đoạn trong ô bảng)" : "") + " và " + tail;
        }

        /// <summary>3 &amp; 4. Thêm 1 trang ngang (hoặc dọc) ngay dưới trang đặt con trỏ, giữ nguyên quy cách trang ngay dưới trang mới thêm.</summary>
        public string InsertOrientedPage(AppSettings st, bool landscape)
        {
            if (!HasDocument()) throw new UserMsg("Chưa có tài liệu Word nào đang mở.");
            dynamic doc = app.ActiveDocument;
            dynamic sel = app.Selection;

            dynamic curSec = sel.Sections.Item(1);
            int curSecIdx = (int)curSec.Index;
            int curOrient = 0;
            try { curOrient = (int)curSec.PageSetup.Orientation; } catch { curOrient = 0; }
            int targetOrient = landscape ? 1 : 0;

            if (curOrient == targetOrient)
            {
                WordTables.Fast((object)app, landscape ? "Thêm trang ngang" : "Thêm trang dọc", delegate
                {
                    int curPage = 1, totalPages = 1;
                    try { curPage = (int)sel.Information(3); totalPages = (int)sel.Information(4); } catch { }
                    if (curPage < totalPages)
                    {
                        try { sel.GoTo(1, 1, curPage + 1); sel.MoveLeft(1, 1); } catch { }
                    }
                    else
                    {
                        try { sel.EndKey(6); } catch { }
                    }
                    sel.InsertBreak(7); // wdPageBreak
                });
                return "Trang hiện tại đã là khổ " + (landscape ? "ngang" : "dọc") + "; đã thêm 1 trang mới cùng khổ ngay bên dưới.";
            }

            WordTables.Fast((object)app, landscape ? "Thêm 1 trang ngang" : "Thêm 1 trang dọc", delegate
            {
                int curPage = 1, totalPages = 1;
                try { curPage = (int)sel.Information(3); totalPages = (int)sel.Information(4); } catch { }
                bool hasNextPage = curPage < totalPages;
                if (hasNextPage)
                {
                    sel.GoTo(1, 1, curPage + 1); // wdGoToPage, wdGoToAbsolute -> đầu trang tiếp theo
                }
                else
                {
                    sel.EndKey(6); // wdStory -> cuối tài liệu
                }

                // Chèn ngắt phần sang trang mới để tạo trang mới với khổ mong muốn
                sel.InsertBreak(2); // wdSectionBreakNextPage
                int newSecIdx = curSecIdx + 1;
                int newPageStart = (int)sel.Start;

                if (hasNextPage)
                {
                    // Chèn ngắt phần thứ hai để các trang ngay bên dưới giữ nguyên khổ ban đầu
                    sel.InsertBreak(2); // wdSectionBreakNextPage
                    dynamic nextSec = doc.Sections.Item(newSecIdx + 1);
                    nextSec.PageSetup.Orientation = curOrient;
                    ApplyMargins(nextSec.PageSetup, st, curOrient);
                }

                dynamic midSec = doc.Sections.Item(newSecIdx);
                midSec.PageSetup.Orientation = targetOrient;
                ApplyMargins(midSec.PageSetup, st, targetOrient);

                // Đặt con trỏ vào trang mới vừa thêm
                try { sel.SetRange(newPageStart, newPageStart); } catch { }
            });

            return "Đã thêm 1 trang khổ " + (landscape ? "ngang" : "dọc") + " ngay dưới trang hiện tại; các trang bên dưới giữ nguyên khổ " + (curOrient == 1 ? "ngang" : "dọc") + ".";
        }

        static string Fmt(double v) { return v.ToString("0.##", CultureInfo.CurrentCulture); }
    }
}
