// Trợ lý Kỹ thuật - Rev0.9.5.1: bộ icon Ribbon mới vẽ theo phong cách phẳng, line art nét xám kèm màu nhấn như thẻ Insert của Word.
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;

namespace TroLyKyThuat
{
    /// <summary>Chuyển Image của .NET thành IPictureDisp để Ribbon dùng được.</summary>
    sealed class PictureConverter : AxHost
    {
        PictureConverter() : base("59EE46BA-677D-4d20-BF10-8D8067CB8B33") { }
        public static object ToPicture(Image img) { return GetIPictureDispFromPicture(img); }
    }

    public static class Icons
    {
        // Màu chuẩn phương án A
        static readonly Color Dark = Color.FromArgb(68, 68, 68);         // Nét chính viền xám đậm
        static readonly Color Blue = Color.FromArgb(0, 120, 212);        // Màu nhấn chủ đạo
        static readonly Color Green = Color.FromArgb(46, 158, 91);       // Thêm / xác nhận
        static readonly Color Red = Color.FromArgb(209, 52, 56);         // Xóa / lỗi / giảm
        static readonly Color Orange = Color.FromArgb(232, 119, 46);     // Điểm nhấn phụ
        static readonly Color GrayLine = Color.FromArgb(138, 138, 138);  // Nét phụ / đứt

        // Nền nhạt
        static readonly Color BgLightBlue = Color.FromArgb(234, 241, 251);
        static readonly Color BgCellBlue  = Color.FromArgb(214, 235, 251);
        static readonly Color BgZebraBlue = Color.FromArgb(207, 230, 248);
        static readonly Color BgLightGray = Color.FromArgb(243, 243, 243);
        static readonly Color BgLightRed  = Color.FromArgb(251, 227, 228);

        /// <summary>Vẽ icon 32x32 theo mã nút và kiểu (id:kiểu).</summary>
        public static Bitmap Draw(string id)
        {
            string style = "us";
            if (!string.IsNullOrEmpty(id))
            {
                int colon = id.IndexOf(':');
                if (colon > 0) { style = id.Substring(colon + 1); id = id.Substring(0, colon); }
            }
            var bmp = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                g.Clear(Color.Transparent);
                switch (id)
                {
                    // --- Bảng ---
                    case "btnApplyAll":
                        DrawTableBase(g, 4, 6, 24, 20);
                        BadgeMark(g, 25, 24, 5.5f, 'v', Green);
                        break;
                    case "btnApplySel":
                        DrawTableBase(g, 5, 7, 22, 18);
                        using (var p = new Pen(Blue, 1.4f) { DashStyle = DashStyle.Dash })
                        {
                            p.DashPattern = new float[] { 2.5f, 2.0f };
                            using (var path = RoundPath(2, 4, 28, 24, 2f)) g.DrawPath(p, path);
                        }
                        break;
                    case "btnLayout":
                        DrawTableBase(g, 4, 6, 24, 20);
                        using (var p = new Pen(Blue, 2.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawArc(p, 17, 16, 11, 11, 40, 270);
                            using (var b = new SolidBrush(Blue))
                            {
                                PointF[] arrow = { new PointF(27, 18), new PointF(29.5f, 23), new PointF(24, 22) };
                                g.FillPolygon(b, arrow);
                            }
                        }
                        break;
                    case "mnuHdr":
                        DrawTableBase(g, 4, 6, 24, 20);
                        using (var f = new Font("Segoe UI", 9f, FontStyle.Bold))
                        using (var b = new SolidBrush(Color.White))
                        {
                            g.DrawString("A", f, b, 12, 4.5f);
                        }
                        break;

                    // --- Thêm / Xóa Hàng & Cột ---
                    case "mnuAddRow":
                    case "btnRowBelow":
                        DrawTableMini(g, 4, 4, 24, 16, 5);
                        using (var b = new SolidBrush(BgCellBlue))
                        using (var p = new Pen(Blue, 1.2f) { DashStyle = DashStyle.Dash })
                        {
                            p.DashPattern = new float[] { 2.0f, 1.6f };
                            g.FillRectangle(b, 4, 22, 24, 6);
                            g.DrawRectangle(p, 4, 22, 24, 6);
                        }
                        BadgeMark(g, 25, 24, 5f, '+', Green);
                        break;
                    case "btnRowAbove":
                        using (var b = new SolidBrush(BgCellBlue))
                        using (var p = new Pen(Blue, 1.2f) { DashStyle = DashStyle.Dash })
                        {
                            p.DashPattern = new float[] { 2.0f, 1.6f };
                            g.FillRectangle(b, 4, 4, 24, 6);
                            g.DrawRectangle(p, 4, 4, 24, 6);
                        }
                        DrawTableMini(g, 4, 12, 24, 16, 5);
                        BadgeMark(g, 25, 7, 5f, '+', Green);
                        break;
                    case "mnuAddCol":
                    case "btnColRight":
                        DrawTableMini(g, 3, 4, 18, 24, 5);
                        using (var b = new SolidBrush(BgCellBlue))
                        using (var p = new Pen(Blue, 1.2f) { DashStyle = DashStyle.Dash })
                        {
                            p.DashPattern = new float[] { 2.0f, 1.6f };
                            g.FillRectangle(b, 23, 4, 6, 24);
                            g.DrawRectangle(p, 23, 4, 6, 24);
                        }
                        BadgeMark(g, 25, 24, 5f, '+', Green);
                        break;
                    case "btnColLeft":
                        using (var b = new SolidBrush(BgCellBlue))
                        using (var p = new Pen(Blue, 1.2f) { DashStyle = DashStyle.Dash })
                        {
                            p.DashPattern = new float[] { 2.0f, 1.6f };
                            g.FillRectangle(b, 3, 4, 6, 24);
                            g.DrawRectangle(p, 3, 4, 6, 24);
                        }
                        DrawTableMini(g, 11, 4, 18, 24, 5);
                        BadgeMark(g, 7, 24, 5f, '+', Green);
                        break;
                    case "mnuDel":
                    case "btnDelRow":
                        DrawTableBase(g, 4, 5, 24, 20);
                        using (var b = new SolidBrush(BgLightRed))
                        {
                            g.FillRectangle(b, 4.5f, 18.5f, 23, 6);
                        }
                        using (var p = new Pen(Dark, 1.3f))
                        {
                            g.DrawLine(p, 4, 18.5f, 28, 18.5f);
                        }
                        BadgeMark(g, 25, 24, 5.5f, 'x', Red);
                        break;
                    case "btnDelCol":
                        DrawTableBase(g, 4, 5, 24, 20);
                        using (var b = new SolidBrush(BgLightRed))
                        {
                            g.FillRectangle(b, 20.5f, 10.5f, 7, 14);
                        }
                        using (var p = new Pen(Dark, 1.3f))
                        {
                            g.DrawLine(p, 20.5f, 10.5f, 20.5f, 24.5f);
                        }
                        BadgeMark(g, 25, 24, 5.5f, 'x', Red);
                        break;

                    // --- Tính toán ---
                    case "mnuCalc":
                    case "btnSum":
                        DrawCalcBox(g);
                        using (var p = new Pen(Blue, 2.3f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            PointF[] pts = { new PointF(21, 10), new PointF(11, 10), new PointF(17, 16), new PointF(11, 22), new PointF(21, 22) };
                            g.DrawLines(p, pts);
                        }
                        break;
                    case "btnAvg":
                        DrawCalcBox(g);
                        using (var p = new Pen(Blue, 2.3f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 11, 10.5f, 21, 10.5f); // x-bar
                            g.DrawLine(p, 12, 14.5f, 20, 23.5f);
                            g.DrawLine(p, 20, 14.5f, 12, 23.5f);
                        }
                        break;
                    case "btnMin":
                        DrawCalcBox(g);
                        using (var p = new Pen(Blue, 2.3f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            g.DrawLine(p, 16, 8.5f, 16, 19.5f);
                            PointF[] head = { new PointF(11.5f, 15.5f), new PointF(16, 20f), new PointF(20.5f, 15.5f) };
                            g.DrawLines(p, head);
                            g.DrawLine(p, 10, 23.5f, 22, 23.5f);
                        }
                        break;
                    case "btnMax":
                        DrawCalcBox(g);
                        using (var p = new Pen(Blue, 2.3f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            g.DrawLine(p, 16, 23.5f, 16, 12.5f);
                            PointF[] head = { new PointF(11.5f, 16.5f), new PointF(16, 12f), new PointF(20.5f, 16.5f) };
                            g.DrawLines(p, head);
                            g.DrawLine(p, 10, 8.5f, 22, 8.5f);
                        }
                        break;
                    case "btnCnt":
                        DrawCalcBox(g);
                        using (var p = new Pen(Blue, 2.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 13.5f, 9, 11.5f, 23);
                            g.DrawLine(p, 20.5f, 9, 18.5f, 23);
                            g.DrawLine(p, 9, 13.5f, 23, 13.5f);
                            g.DrawLine(p, 8, 18.5f, 22, 18.5f);
                        }
                        break;
                    case "btnFmtCalc":
                        DrawCalcBox(g);
                        using (var p = new Pen(Blue, 2.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            PointF[] pts = { new PointF(18, 9.5f), new PointF(9.5f, 9.5f), new PointF(14.5f, 15.5f), new PointF(9.5f, 21.5f), new PointF(18, 21.5f) };
                            g.DrawLines(p, pts);
                        }
                        using (var f = new Font("Segoe UI", 9.5f, FontStyle.Bold))
                        using (var b = new SolidBrush(Dark))
                        {
                            g.DrawString("B", f, b, 17, 11.5f);
                        }
                        break;
                    case "btnFmtProd":
                        DrawCalcBox(g);
                        using (var p = new Pen(Blue, 2.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 8.5f, 9.5f, 19f, 9.5f);
                            g.DrawLine(p, 11.5f, 9.5f, 11.5f, 22f);
                            g.DrawLine(p, 16.5f, 9.5f, 16.5f, 22f);
                        }
                        using (var f = new Font("Segoe UI", 9.5f, FontStyle.Bold))
                        using (var b = new SolidBrush(Dark))
                        {
                            g.DrawString("B", f, b, 17, 11.5f);
                        }
                        break;

                    // --- Định dạng số ---
                    case "mnuNum":
                    case "btnNumStyle":
                        DrawCalcBox(g);
                        string numTxt = style == "vn" ? "1.2" : (style == "si" ? "1 2" : "1,2");
                        using (var f = new Font("Segoe UI", 12f, FontStyle.Bold, GraphicsUnit.Pixel))
                        using (var b = new SolidBrush(Dark))
                        using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        {
                            g.DrawString(numTxt, f, b, new RectangleF(4, 5, 24, 22), sf);
                        }
                        break;
                    case "btnDecInc":
                        DrawCalcBox(g);
                        using (var f = new Font("Segoe UI", 12f, FontStyle.Bold, GraphicsUnit.Pixel))
                        using (var b = new SolidBrush(Dark))
                        {
                            g.DrawString(".0", f, b, 5, 9);
                        }
                        BadgeMark(g, 24, 9, 4.5f, '+', Green);
                        break;
                    case "btnDecDec":
                        DrawCalcBox(g);
                        using (var f = new Font("Segoe UI", 12f, FontStyle.Bold, GraphicsUnit.Pixel))
                        using (var b = new SolidBrush(Dark))
                        {
                            g.DrawString(".0", f, b, 5, 9);
                        }
                        BadgeMark(g, 24, 9, 4.5f, '-', Red);
                        break;
                    case "btnSync":
                        DrawCalcBox(g);
                        if (style == "sci")
                        {
                            using (var f = new Font("Segoe UI", 8.5f, FontStyle.Bold, GraphicsUnit.Pixel))
                            using (var b = new SolidBrush(Dark))
                            {
                                g.DrawString("a×10ⁿ", f, b, 5, 10);
                            }
                        }
                        else
                        {
                            using (var f = new Font("Segoe UI", 8.5f, FontStyle.Bold, GraphicsUnit.Pixel))
                            using (var b = new SolidBrush(Dark))
                            {
                                string d = (style.Length > 0 && char.IsDigit(style[0])) ? style : "2";
                                g.DrawString(".00", f, b, 5, 7);
                                g.DrawString("." + d.PadRight(2, '0'), f, b, 5, 15);
                            }
                            using (var p = new Pen(Blue, 1.5f))
                            {
                                g.DrawLine(p, 23, 7, 23, 23);
                            }
                        }
                        break;

                    // --- Căn lề bảng ---
                    case "btnTblAlL":
                        DrawPageMiniWithTable(g, 8.5f);
                        break;
                    case "btnTblAlC":
                        DrawPageMiniWithTable(g, 12.5f);
                        break;
                    case "btnTblAlR":
                        DrawPageMiniWithTable(g, 16.5f);
                        break;

                    // --- Căn ô / tiêu đề ---
                    case "mnuHdrAl":
                        using (var p = new Pen(Dark, 2.4f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 5, 7, 27, 7);
                            g.DrawLine(p, 9, 13, 23, 13);
                            g.DrawLine(p, 5, 19, 27, 19);
                            g.DrawLine(p, 9, 25, 23, 25);
                        }
                        break;
                    case "mnuTxtAl":
                        using (var p = new Pen(Dark, 2.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 5, 7, 27, 7);
                            g.DrawLine(p, 5, 13, 21, 13);
                            g.DrawLine(p, 5, 19, 25, 19);
                            g.DrawLine(p, 5, 25, 17, 25);
                        }
                        break;
                    case "mnuNumAl":
                        using (var p = new Pen(Dark, 2.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 5, 7, 27, 7);
                            g.DrawLine(p, 11, 13, 27, 13);
                            g.DrawLine(p, 7, 19, 27, 19);
                            g.DrawLine(p, 15, 25, 27, 25);
                        }
                        break;

                    // --- Viền, Kiểu nét, Tô xen kẽ ---
                    case "mnuBorder":
                        using (var p = new Pen(GrayLine, 1.0f))
                        {
                            g.DrawLine(p, 5, 12, 27, 12);
                            g.DrawLine(p, 5, 19, 27, 19);
                            g.DrawLine(p, 13, 6, 13, 26);
                            g.DrawLine(p, 20, 6, 20, 26);
                        }
                        using (var p = new Pen(Blue, 2.8f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(5, 6, 22, 20, 1.5f))
                        {
                            g.DrawPath(p, path);
                        }
                        break;
                    case "mnuBdSt":
                        using (var p = new Pen(Dark, 2.4f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 4, 8, 28, 8);
                            p.DashStyle = DashStyle.Dash;
                            p.DashPattern = new float[] { 3.0f, 2.0f };
                            g.DrawLine(p, 4, 16, 28, 16);
                            p.DashStyle = DashStyle.Dot;
                            p.DashPattern = new float[] { 1.2f, 1.8f };
                            g.DrawLine(p, 4, 24, 28, 24);
                        }
                        break;
                    case "mnuZebra":
                        using (var b = new SolidBrush(BgZebraBlue))
                        {
                            g.FillRectangle(b, 4, 6, 24, 5);
                            g.FillRectangle(b, 4, 16, 24, 5);
                        }
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(4, 6, 24, 20, 2f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var p = new Pen(Dark, 1.1f))
                        {
                            g.DrawLine(p, 4, 11, 28, 11);
                            g.DrawLine(p, 4, 16, 28, 16);
                            g.DrawLine(p, 4, 21, 28, 21);
                        }
                        break;

                    // --- Màu sắc ---
                    case "btnHdrFg":
                    case "swA":
                        {
                            Color c = HexColor(style);
                            using (var p = new Pen(Dark, 2.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                            {
                                PointF[] ptsA = { new PointF(8, 21), new PointF(16, 6), new PointF(24, 21) };
                                g.DrawLines(p, ptsA);
                                g.DrawLine(p, 11, 16.5f, 21, 16.5f);
                            }
                            using (var b = new SolidBrush(c))
                            using (var path = RoundPath(6, 24, 20, 4, 1f))
                            {
                                g.FillPath(b, path);
                            }
                        }
                        break;
                    case "btnHdrBg":
                        {
                            Color c = HexColor(style);
                            DrawTableMini(g, 4, 7, 24, 18, 6, c);
                            using (var b = new SolidBrush(c))
                            using (var p = new Pen(Dark, 1.0f))
                            {
                                g.FillEllipse(b, 21, 19, 7, 7);
                                g.DrawEllipse(p, 21, 19, 7, 7);
                            }
                        }
                        break;
                    case "btnBdClr":
                        {
                            Color c = HexColor(style);
                            using (var p = new Pen(GrayLine, 1.0f))
                            {
                                g.DrawLine(p, 5, 12, 27, 12);
                                g.DrawLine(p, 5, 19, 27, 19);
                                g.DrawLine(p, 13, 6, 13, 26);
                                g.DrawLine(p, 20, 6, 20, 26);
                            }
                            using (var p = new Pen(c, 2.6f) { LineJoin = LineJoin.Round })
                            using (var path = RoundPath(5, 6, 22, 20, 1.5f))
                            {
                                g.DrawPath(p, path);
                            }
                        }
                        break;
                    case "btnZebraClr":
                        {
                            Color c = HexColor(style);
                            using (var b = new SolidBrush(c))
                            {
                                g.FillRectangle(b, 4, 6, 24, 5);
                                g.FillRectangle(b, 4, 16, 24, 5);
                            }
                            using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                            using (var path = RoundPath(4, 6, 24, 20, 2f))
                            {
                                g.DrawPath(p, path);
                            }
                            using (var p = new Pen(Dark, 1.1f))
                            {
                                g.DrawLine(p, 4, 11, 28, 11);
                                g.DrawLine(p, 4, 16, 28, 16);
                                g.DrawLine(p, 4, 21, 28, 21);
                            }
                        }
                        break;
                    case "sw":
                        {
                            Color c = (style == "us" || string.IsNullOrEmpty(style)) ? Color.FromArgb(170, 170, 170) : HexColor(style);
                            using (var b = new SolidBrush(c))
                            using (var path = RoundPath(5, 5, 22, 22, 2f))
                            {
                                g.FillPath(b, path);
                            }
                            using (var p = new Pen(Color.FromArgb(110, 110, 110), 1.5f))
                            using (var path = RoundPath(5, 5, 22, 22, 2f))
                            {
                                g.DrawPath(p, path);
                            }
                        }
                        break;
                    case "gs":
                        {
                            Color c = (style == "us" || string.IsNullOrEmpty(style)) ? Color.FromArgb(170, 170, 170) : HexColor(style);
                            using (var b = new SolidBrush(c))
                            using (var path = RoundPath(5, 5, 22, 22, 2f))
                            {
                                g.FillPath(b, path);
                            }
                            using (var p = new Pen(Color.FromArgb(140, 140, 140), 1.4f))
                            using (var path = RoundPath(5, 5, 22, 22, 2f))
                            {
                                g.DrawPath(p, path);
                            }
                        }
                        break;

                    // --- Ngôn ngữ & Rà soát ---
                    case "btnTerm":
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(7, 4, 19, 24, 2f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var b = new SolidBrush(Blue))
                        {
                            g.FillRectangle(b, 7, 4, 4, 24);
                        }
                        using (var p = new Pen(Dark, 1.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 14, 9, 23, 9);
                            g.DrawLine(p, 14, 13, 21, 13);
                        }
                        using (var p = new Pen(Green, 2.4f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            PointF[] check = { new PointF(15.5f, 18.5f), new PointF(18.5f, 21.5f), new PointF(24, 14.5f) };
                            g.DrawLines(p, check);
                        }
                        break;
                    case "btnLookup":
                        using (var b = new SolidBrush(BgLightBlue))
                        using (var p = new Pen(Dark, 1.8f))
                        {
                            g.FillEllipse(b, 5.5f, 5.5f, 15, 15);
                            g.DrawEllipse(p, 5.5f, 5.5f, 15, 15);
                        }
                        using (var p = new Pen(Blue, 3.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 19.5f, 19.5f, 27, 27);
                        }
                        break;
                    case "btnSyn":
                        using (var p = new Pen(Blue, 2.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            g.DrawLine(p, 6, 11, 25, 11);
                            PointF[] ar1 = { new PointF(21, 7), new PointF(25, 11), new PointF(21, 15) };
                            g.DrawLines(p, ar1);
                        }
                        using (var p = new Pen(Dark, 2.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            g.DrawLine(p, 26, 21, 7, 21);
                            PointF[] ar2 = { new PointF(11, 17), new PointF(7, 21), new PointF(11, 25) };
                            g.DrawLines(p, ar2);
                        }
                        break;
                    case "btnAiTools":
                        using (var p = new Pen(Dark, 1.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        using (var path = RoundPath(4, 11, 24, 16, 3f))
                        {
                            g.DrawPath(p, path);
                            g.DrawLine(p, 9, 17, 21, 17);
                            g.DrawLine(p, 9, 22, 17, 22);
                        }
                        using (var p = new Pen(Blue, 1.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 22, 3, 22, 13);
                            g.DrawLine(p, 17, 8, 27, 8);
                        }
                        break;
                    case "btnGram":
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(6, 3, 20, 26, 2f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var p = new Pen(Dark, 1.4f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 10, 9, 22, 9);
                            g.DrawLine(p, 10, 14, 18, 14);
                        }
                        using (var p = new Pen(Blue, 1.7f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            PointF[] wave = {
                                new PointF(10, 21), new PointF(11.5f, 19), new PointF(13, 21), new PointF(14.5f, 23),
                                new PointF(16, 21), new PointF(17.5f, 19), new PointF(19, 21), new PointF(20.5f, 23), new PointF(22, 21)
                            };
                            g.DrawCurve(p, wave, 0.5f);
                        }
                        break;
                    case "btnStyle":
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(6, 3, 20, 26, 2f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var p = new Pen(Dark, 1.4f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 10, 8, 22, 8);
                            g.DrawLine(p, 10, 13, 20, 13);
                            g.DrawLine(p, 10, 18, 22, 18);
                        }
                        using (var p = new Pen(Orange, 1.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            PointF[] wave = {
                                new PointF(10, 23), new PointF(13, 21), new PointF(16, 24), new PointF(19, 21), new PointF(22, 23)
                            };
                            g.DrawCurve(p, wave, 0.5f);
                        }
                        break;
                    case "btnReview":
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(5, 3, 17, 24, 2f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var p = new Pen(Dark, 1.3f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 8, 8, 18, 8);
                            g.DrawLine(p, 8, 12, 16, 12);
                            g.DrawLine(p, 8, 16, 18, 16);
                        }
                        using (var b = new SolidBrush(BgLightBlue))
                        using (var p = new Pen(Blue, 1.8f))
                        {
                            g.FillEllipse(b, 15, 14, 11, 11);
                            g.DrawEllipse(p, 15, 14, 11, 11);
                        }
                        using (var p = new Pen(Blue, 2.6f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 24, 23, 28, 27);
                        }
                        break;
                    case "btnRead":
                        // Rev0.9.5.2: "Đọc số" là VIẾT SỐ BẰNG CHỮ (123 -> một trăm hai mươi ba), không phải đọc thành tiếng. Số ở trên, mũi tên xuống, dòng chữ ở dưới.
                        using (var f = new Font("Segoe UI", 8.5f, FontStyle.Bold, GraphicsUnit.Point))
                        using (var b = new SolidBrush(Dark))
                        using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        {
                            g.DrawString("123", f, b, new RectangleF(3, 1, 26, 13), sf);
                        }
                        using (var p = new Pen(Blue, 2.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            g.DrawLine(p, 16, 15, 16, 20);
                            g.DrawLine(p, 12.5f, 17, 16, 20.5f);
                            g.DrawLine(p, 19.5f, 17, 16, 20.5f);
                        }
                        using (var p = new Pen(Dark, 2.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 6, 24.5f, 26, 24.5f);
                            g.DrawLine(p, 6, 28.5f, 17, 28.5f);
                        }
                        break;
                    case "btnQuick":
                        using (var p = new Pen(Dark, 3.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 7, 25, 20, 12);
                        }
                        using (var p = new Pen(Blue, 2.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 24, 4, 24, 11);
                            g.DrawLine(p, 20.5f, 7.5f, 27.5f, 7.5f);
                        }
                        using (var p = new Pen(Blue, 1.5f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 25.5f, 15, 25.5f, 19);
                            g.DrawLine(p, 23.5f, 17, 27.5f, 17);
                        }
                        break;

                    // --- Công thức & Chỉ số ---
                    case "btnCompose":
                        DrawCalcBox(g);
                        using (var p = new Pen(Blue, 2.4f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            g.DrawLine(p, 9, 10.5f, 23, 10.5f);
                            g.DrawLine(p, 13, 10.5f, 12, 22.5f);
                            PointF[] legR = { new PointF(20, 10.5f), new PointF(20, 20f), new PointF(22.5f, 22.5f) };
                            g.DrawLines(p, legR);
                        }
                        break;
                    case "btnIndex":
                        DrawCalcBox(g);
                        using (var p = new Pen(Dark, 2.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 8, 11, 15, 21);
                            g.DrawLine(p, 15, 11, 8, 21);
                        }
                        using (var p = new Pen(Blue, 2.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            PointF[] two = { new PointF(18, 17), new PointF(21, 16), new PointF(23, 18), new PointF(18, 23), new PointF(23.5f, 23) };
                            g.DrawLines(p, two);
                        }
                        break;

                    // --- Dàn trang ---
                    case "tbPageSetup":
                    case "btnPageSetup":
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(6, 3, 20, 26, 2f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var p = new Pen(Blue, 1.3f) { DashStyle = DashStyle.Dash })
                        {
                            p.DashPattern = new float[] { 2.0f, 1.6f };
                            g.DrawRectangle(p, 10, 7, 12, 18);
                        }
                        break;
                    case "tbParaFormat":
                    case "btnParaFormat":
                        using (var p = new Pen(Dark, 1.6f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 6, 8, 26, 8);
                            g.DrawLine(p, 6, 14, 26, 14);
                            g.DrawLine(p, 6, 20, 20, 20);
                            g.DrawLine(p, 6, 26, 18, 26);
                        }
                        using (var p = new Pen(Blue, 1.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 24, 17, 24, 27);
                            PointF[] ar1 = { new PointF(22, 19), new PointF(24, 17), new PointF(26, 19) };
                            PointF[] ar2 = { new PointF(22, 25), new PointF(24, 27), new PointF(26, 25) };
                            g.DrawLines(p, ar1);
                            g.DrawLines(p, ar2);
                        }
                        break;
                    case "tbAddLand":
                    case "btnAddLand":
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(3, 8, 26, 16, 1.5f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var p = new Pen(Dark, 1.3f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 8, 13, 20, 13);
                            g.DrawLine(p, 8, 17, 17, 17);
                        }
                        BadgeMark(g, 25, 23, 5f, '+', Green);
                        break;
                    case "tbAddPort":
                    case "btnAddPort":
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(8, 3, 16, 26, 1.5f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var p = new Pen(Dark, 1.3f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 12, 9, 20, 9);
                            g.DrawLine(p, 12, 14, 20, 14);
                            g.DrawLine(p, 12, 19, 18, 19);
                        }
                        BadgeMark(g, 24, 24, 5f, '+', Green);
                        break;
                    case "btnSpInc":
                        using (var p = new Pen(Dark, 1.5f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 6, 7, 26, 7);
                            g.DrawLine(p, 6, 25, 26, 25);
                        }
                        using (var p = new Pen(Blue, 2.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            g.DrawLine(p, 16, 14, 16, 10);
                            PointF[] arUp = { new PointF(13.5f, 12.5f), new PointF(16, 10), new PointF(18.5f, 12.5f) };
                            g.DrawLines(p, arUp);

                            g.DrawLine(p, 16, 18, 16, 22);
                            PointF[] arDn = { new PointF(13.5f, 19.5f), new PointF(16, 22), new PointF(18.5f, 19.5f) };
                            g.DrawLines(p, arDn);
                        }
                        break;
                    case "btnSpDec":
                        using (var p = new Pen(Dark, 1.5f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 6, 6, 26, 6);
                            g.DrawLine(p, 6, 26, 26, 26);
                        }
                        using (var p = new Pen(Blue, 2.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            g.DrawLine(p, 16, 10, 16, 14);
                            PointF[] arDn = { new PointF(13.5f, 11.5f), new PointF(16, 14), new PointF(18.5f, 11.5f) };
                            g.DrawLines(p, arDn);

                            g.DrawLine(p, 16, 22, 16, 18);
                            PointF[] arUp = { new PointF(13.5f, 20.5f), new PointF(16, 18), new PointF(18.5f, 20.5f) };
                            g.DrawLines(p, arUp);
                        }
                        break;

                    // --- Bảng điều khiển & Cài đặt ---
                    case "btnSettings":
                        using (var p = new Pen(Dark, 3.4f) { StartCap = LineCap.Square, EndCap = LineCap.Square })
                        {
                            g.DrawLine(p, 25, 16, 28.5f, 16);
                            g.DrawLine(p, 22.4f, 22.4f, 24.8f, 24.8f);
                            g.DrawLine(p, 16, 25, 16, 28.5f);
                            g.DrawLine(p, 9.6f, 22.4f, 7.2f, 24.8f);
                            g.DrawLine(p, 7, 16, 3.5f, 16);
                            g.DrawLine(p, 9.6f, 9.6f, 7.2f, 7.2f);
                            g.DrawLine(p, 16, 7, 16, 3.5f);
                            g.DrawLine(p, 22.4f, 9.6f, 24.8f, 7.2f);
                        }
                        using (var b = new SolidBrush(BgLightGray))
                        using (var p = new Pen(Dark, 1.6f))
                        {
                            g.FillEllipse(b, 7, 7, 18, 18);
                            g.DrawEllipse(p, 7, 7, 18, 18);
                        }
                        using (var b = new SolidBrush(Color.White))
                        using (var p = new Pen(Blue, 1.8f))
                        {
                            g.FillEllipse(b, 12.6f, 12.6f, 6.8f, 6.8f);
                            g.DrawEllipse(p, 12.6f, 12.6f, 6.8f, 6.8f);
                        }
                        break;
                    case "btnPane":
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(4, 6, 24, 20, 2f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var b = new SolidBrush(BgZebraBlue))
                        {
                            g.FillRectangle(b, 19, 6.5f, 8.5f, 19);
                        }
                        using (var p = new Pen(Dark, 1.4f))
                        {
                            g.DrawLine(p, 19, 6, 19, 26);
                        }
                        using (var p = new Pen(Blue, 1.4f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 22, 11, 25.5f, 11);
                            g.DrawLine(p, 22, 15, 25.5f, 15);
                            g.DrawLine(p, 22, 19, 25.5f, 19);
                        }
                        break;
                    case "btnCopy":
                        using (var p = new Pen(Dark, 1.5f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(6, 3, 16, 19, 2f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var b = new SolidBrush(BgLightBlue))
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(11, 8, 16, 20, 2f))
                        {
                            g.FillPath(b, path);
                            g.DrawPath(p, path);
                        }
                        using (var p = new Pen(Blue, 1.4f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(p, 15, 14, 23, 14);
                            g.DrawLine(p, 15, 18, 23, 18);
                            g.DrawLine(p, 15, 22, 20, 22);
                        }
                        break;
                    case "btnCsv":
                        using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
                        using (var path = RoundPath(6, 3, 20, 26, 2f))
                        {
                            g.DrawPath(p, path);
                        }
                        using (var pg = new Pen(Blue, 1.3f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            // trang chứa một bảng nhỏ: dễ đọc hơn chữ "CSV" nhồi trong ô 20 px
                            g.DrawRectangle(pg, 10, 9, 12, 14);
                            g.DrawLine(pg, 10, 14, 22, 14);
                            g.DrawLine(pg, 10, 19, 22, 19);
                            g.DrawLine(pg, 16, 9, 16, 23);
                        }
                        break;
                    case "galFont":
                        using (var p = new Pen(Dark, 2.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            PointF[] ptsA = { new PointF(6, 23), new PointF(13, 8), new PointF(20, 23) };
                            g.DrawLines(p, ptsA);
                            g.DrawLine(p, 9, 18, 17, 18);
                        }
                        using (var p = new Pen(Blue, 2.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawArc(p, 20, 16, 6, 7, 0, 360);
                            g.DrawLine(p, 26, 16, 26, 23);
                        }
                        break;
                    case "btnNInc":
                        using (var p = new Pen(Blue, 2.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            PointF[] pts = { new PointF(8, 19), new PointF(16, 11), new PointF(24, 19) };
                            g.DrawLines(p, pts);
                        }
                        break;
                    case "btnNDec":
                        using (var p = new Pen(Blue, 2.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                        {
                            PointF[] pts = { new PointF(8, 13), new PointF(16, 21), new PointF(24, 13) };
                            g.DrawLines(p, pts);
                        }
                        break;

                    // --- Icon dự phòng cho mã lạ ---
                    default:
                        using (var p = new Pen(GrayLine, 1.4f) { DashStyle = DashStyle.Dash })
                        {
                            p.DashPattern = new float[] { 3.0f, 3.0f };
                            using (var path = RoundPath(5, 5, 22, 22, 2f)) g.DrawPath(p, path);
                        }
                        using (var f = new Font("Segoe UI", 13f, FontStyle.Bold))
                        using (var b = new SolidBrush(Dark))
                        using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        {
                            g.DrawString("?", f, b, new RectangleF(5, 5, 22, 22), sf);
                        }
                        break;
                }
            }
            return bmp;
        }

        static void DrawCalcBox(Graphics g)
        {
            using (var b = new SolidBrush(BgLightBlue))
            using (var path = RoundPath(4, 5, 24, 22, 3f))
            {
                g.FillPath(b, path);
            }
            using (var p = new Pen(Dark, 1.4f) { LineJoin = LineJoin.Round })
            using (var path = RoundPath(4, 5, 24, 22, 3f))
            {
                g.DrawPath(p, path);
            }
        }

        static void DrawTableBase(Graphics g, float x, float y, float w, float h)
        {
            using (var b = new SolidBrush(Blue))
            using (var path = TopRoundPath(x, y, w, 6, 2f))
            {
                g.FillPath(b, path);
            }
            using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
            using (var path = RoundPath(x, y, w, h, 2f))
            {
                g.DrawPath(p, path);
            }
            using (var p = new Pen(Dark, 1.3f))
            {
                g.DrawLine(p, x, y + 6, x + w, y + 6);
                g.DrawLine(p, x, y + 13, x + w, y + 13);
                g.DrawLine(p, x + w * 0.38f, y + 6, x + w * 0.38f, y + h);
                g.DrawLine(p, x + w * 0.71f, y + 6, x + w * 0.71f, y + h);
            }
        }

        static void DrawTableMini(Graphics g, float x, float y, float w, float h, float hdrH, Color? hdrClr = null)
        {
            Color c = hdrClr.HasValue ? hdrClr.Value : Blue;
            using (var b = new SolidBrush(c))
            using (var path = TopRoundPath(x, y, w, hdrH, 2f))
            {
                g.FillPath(b, path);
            }
            using (var p = new Pen(Dark, 1.5f) { LineJoin = LineJoin.Round })
            using (var path = RoundPath(x, y, w, h, 2f))
            {
                g.DrawPath(p, path);
            }
            using (var p = new Pen(Dark, 1.1f))
            {
                g.DrawLine(p, x, y + hdrH, x + w, y + hdrH);
                g.DrawLine(p, x + w * 0.5f, y + hdrH, x + w * 0.5f, y + h);
            }
        }

        static void DrawPageMiniWithTable(Graphics g, float tblX)
        {
            using (var p = new Pen(Dark, 1.6f) { LineJoin = LineJoin.Round })
            using (var path = RoundPath(6, 3, 20, 26, 2f))
            {
                g.DrawPath(p, path);
            }
            using (var b = new SolidBrush(Blue))
            using (var path = TopRoundPath(tblX, 10, 7, 4, 1.5f))
            {
                g.FillPath(b, path);
            }
            using (var p = new Pen(Dark, 1.2f) { LineJoin = LineJoin.Round })
            using (var path = RoundPath(tblX, 10, 7, 12, 1.5f))
            {
                g.DrawPath(p, path);
            }
            using (var p = new Pen(Dark, 1.0f))
            {
                g.DrawLine(p, tblX, 14, tblX + 7, 14);
                g.DrawLine(p, tblX, 18, tblX + 7, 18);
            }
        }

        static void BadgeMark(Graphics g, float cx, float cy, float r, char mark, Color bg)
        {
            using (var b = new SolidBrush(bg))
            using (var p = new Pen(Color.White, 1.4f))
            {
                g.FillEllipse(b, cx - r, cy - r, r * 2, r * 2);
                g.DrawEllipse(p, cx - r, cy - r, r * 2, r * 2);
            }
            using (var p = new Pen(Color.White, 1.7f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            {
                if (mark == '+')
                {
                    g.DrawLine(p, cx - r * 0.5f, cy, cx + r * 0.5f, cy);
                    g.DrawLine(p, cx, cy - r * 0.5f, cx, cy + r * 0.5f);
                }
                else if (mark == '-')
                {
                    g.DrawLine(p, cx - r * 0.5f, cy, cx + r * 0.5f, cy);
                }
                else if (mark == 'v' || mark == '✓')
                {
                    PointF[] pts = { new PointF(cx - 2.5f, cy), new PointF(cx - 0.5f, cy + 2.2f), new PointF(cx + 2.8f, cy - 2.2f) };
                    g.DrawLines(p, pts);
                }
                else if (mark == 'x' || mark == '✕')
                {
                    g.DrawLine(p, cx - 2.2f, cy - 2.2f, cx + 2.2f, cy + 2.2f);
                    g.DrawLine(p, cx + 2.2f, cy - 2.2f, cx - 2.2f, cy + 2.2f);
                }
            }
        }

        static GraphicsPath RoundPath(float x, float y, float w, float h, float r)
        {
            var p = new GraphicsPath(); float d = r * 2;
            p.AddArc(x, y, d, d, 180, 90);
            p.AddArc(x + w - d, y, d, d, 270, 90);
            p.AddArc(x + w - d, y + h - d, d, d, 0, 90);
            p.AddArc(x, y + h - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        static GraphicsPath TopRoundPath(float x, float y, float w, float h, float r)
        {
            var p = new GraphicsPath(); float d = r * 2;
            p.AddArc(x, y, d, d, 180, 90);
            p.AddArc(x + w - d, y, d, d, 270, 90);
            p.AddLine(x + w, y + h, x, y + h);
            p.CloseFigure();
            return p;
        }

        static Color HexColor(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex.Length < 6) return Blue;
            try
            {
                return Color.FromArgb(
                    Convert.ToInt32(hex.Substring(0, 2), 16),
                    Convert.ToInt32(hex.Substring(2, 2), 16),
                    Convert.ToInt32(hex.Substring(4, 2), 16));
            }
            catch { return Blue; }
        }

        /// <summary>Vẽ mẫu tên font bằng chính kiểu chữ đó để hiển thị trong danh sách chọn font trên Ribbon.</summary>
        public static Bitmap FontPreview(string fontName)
        {
            var bmp = new Bitmap(170, 26, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                g.Clear(Color.White);
                string fn = string.IsNullOrEmpty(fontName) ? "Times New Roman" : fontName;
                using (var f = new Font(fn, 11.5f, FontStyle.Regular, GraphicsUnit.Point))
                using (var b = new SolidBrush(Color.FromArgb(30, 30, 30)))
                using (var sf = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center })
                {
                    g.DrawString(fn, f, b, new RectangleF(6, 0, 160, 26), sf);
                }
            }
            return bmp;
        }

        public static object FontPicture(string fontName)
        {
            using (Bitmap b = FontPreview(fontName)) return PictureConverter.ToPicture(b);
        }

        public static object Picture(string id)
        {
            using (Bitmap b = Draw(id)) return PictureConverter.ToPicture(b);
        }
    }
}
