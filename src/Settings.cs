// Trợ lý Kỹ thuật - Rev0.4: lưu và nạp cài đặt (tệp văn bản đơn giản trong %AppData%).
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace TroLyKyThuat
{
    public sealed class AppSettings
    {
        // --- Master style bảng ---
        public string HdrBg = "156082";
        public string HdrFg = "FFFFFF";
        public string BorderColor = "000000";
        public double BorderPt = 1.0;
        public int BorderStyle = 0;            // 0 liền, 1 đứt, 2 chấm
        public string FontName = "Times New Roman";
        public double FontSize = 11;
        public int HdrAlign = 1;               // 0 trái, 1 giữa, 2 phải
        public int TextAlign = 0;              // 0 trái, 3 hai bên
        public int NumAlign = 2;               // 0 trái, 1 giữa, 2 phải
        public bool Zebra = false;
        public string ZebraColor = "F2F2F2";
        public bool ClearManual = true;
        public double CellSpace = 2;           // khoảng cách đoạn trước và sau trong ô (pt)
        public bool RepeatHeader = true;
        public bool NoRowBreak = true;
        // --- Quy ước số ---
        public string NumStyle = "us";         // us, vn, si
        public string NumDec = "keep";         // keep hoặc 0..4
        public bool NumThou = true;
        public bool NumAddDecToInt = false;
        // --- Công cụ tính ---
        public string CalcOp = "sum";
        public string CalcDec = "auto";
        // --- Tính theo định dạng ô (Rev0.9.1A, thử nghiệm) ---
        public bool CalcFmtOn = false;          // bật: Tính chỉ gộp các ô khớp CalcFmtKind, bỏ qua ô khác (dù là số)
        public string CalcFmtKind = "bold";     // bold = đậm cả ô, highlight = có tô sáng (highlighter), color = màu chữ không phải Tự động, fill = có tô nền ô
        public bool TopMost = false;
        public int SyncDec = 2;                // nút Đồng bộ số lẻ: số chữ số thập phân (0 đến 5)
        public bool SyncSci = false;           // nút Đồng bộ số lẻ ở chế độ viết kiểu kỹ thuật (a × 10ⁿ)
        public int SciDec = 2;                 // số chữ số thập phân của phần a trong kiểu kỹ thuật (0 đến 5)
        public string UnitNames = "law";       // tên đơn vị khi đọc số: law = theo Nghị định 134/2007/NĐ-CP (kilôpascan), common = tên quen dùng (kilopascal)
        public bool ShowDialog = false;        // hiện hộp thoại báo kết quả sau mỗi lệnh (mặc định tắt, chỉ ghi lên thanh trạng thái của Word)
        public int AddN = 1;                   // số hàng, cột thêm cho mỗi hàng, cột đã chọn
        public string Manual = "";             // danh sách hàng, cột chọn theo số (để trống = dùng vùng chọn trong Word)
        // --- Chuẩn hóa ---
        public string EqMode = "fmt";          // cách viết chỉ số cho công thức cùng dòng với chữ: fmt = định dạng chỉ số trên/dưới, glyph = ký tự chỉ số nhỏ thấp của font Cambria. Công thức nằm riêng một dòng luôn dùng Equation.
        public string RvNumStyle = "auto";     // theo ngôn ngữ quanh số, hoặc chọn riêng vn, us, si
        public bool EqStacked = true;          // Soạn phương trình: phân số đứng (tử trên, mẫu dưới); tắt thì viết phân số một dòng a/b
        public bool EqAlone = false;           // giữ để tương thích file cài đặt cũ (không còn dùng)
        public string RvScope = "all";         // all = toàn văn bản, sel = phần đang chọn
        public bool RvChem = true, RvMath = true, RvSym = true, RvNum = true, RvNumTbl = true, RvIncl = false;
        // --- Thuật ngữ và Văn phong (Rev0.9.5) ---
        public bool TmGloss = true, TmOrth = true, TmAbbr = true;
        public bool StLong = true, StInf = true, StDup = true, StPunc = true, StAi = true, StConn = true, StPron = true;
        public int StLongN = 40;               // câu dài hơn số chữ này thì báo (10 đến 120)
        // --- Ngữ pháp (Rev0.9.2) ---
        public bool GrVi = true, GrEn = true, GrSpell = true, GrAgree = true, GrPunc = true, GrStyle = true, GrTerms = true;
        public int GrLongN = 35;               // câu tiếng Anh dài hơn số chữ này thì báo (15 đến 80)
        // --- Dàn trang (Rev0.9.3) ---
        public string PagePaper = "A4";        // A4, Letter, A3, Legal
        public double PgPortTop = 2.0, PgPortBot = 2.0, PgPortLeft = 3.0, PgPortRight = 2.0;   // cm cho trang dọc (trên-dưới-trái-phải: 2-2-3-2)
        public double PgLandTop = 3.0, PgLandBot = 2.0, PgLandLeft = 2.0, PgLandRight = 2.0;   // cm cho trang ngang (3-2-2-2)
        public double PgHeaderDist = 1.0, PgFooterDist = 1.0;                                  // cm
        public double ParaSizeTnr = 13.0, ParaSizeArial = 11.0;                                // pt cho Times New Roman và Arial
        public double ParaLineSpace = 1.2, ParaSpaceBefore = 6.0, ParaSpaceAfter = 6.0;        // giãn dòng 1.2, trên 6 pt, dưới 6 pt
        public double ParaFirstLineCm = 1.25, ParaHangingCm = 0;                               // thụt dòng đầu (cm) và thụt treo (cm)
        public string ParaAlign = "justify";   // Rev0.9.5.13: canh lề đoạn văn cho nút Chuẩn hóa đoạn văn: left, center, right, justify
        // --- AI trực tuyến (khóa được mã hóa bằng DPAPI theo người dùng Windows hiện tại) ---
        public string ApiProvider = "";
        public string ApiKeyEnc = "";
        public string ApiModel = "";
        public double AiTemp = 0.3;
        public int AiMaxTokens = 1024;
        public string AiTone = "Chuyên nghiệp kỹ thuật";
        public bool RwDoRewrite = true, RwDoTranslate = false;   // Rev0.9.5.6: Công cụ AI, hai ô chọn độc lập
        public string RwTargetLang = "en";                        // mã ngôn ngữ đích (xem AiClient.TargetLangs)
        public int RwInH = 0, RwOutH = 0;                          // chiều cao hai ô văn bản của thẻ Công cụ AI, 0 = mặc định
        public int Ver = 6;                    // phiên bản định dạng tệp cài đặt

        public static string Dir
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TroLyKyThuat"); }
        }
        public static string FilePath { get { return Path.Combine(Dir, "settings.ini"); } }

        public static AppSettings Load()
        {
            var s = new AppSettings();
            try
            {
                if (!File.Exists(FilePath)) return s;
                var map = new Dictionary<string, string>();
                foreach (string line in File.ReadAllLines(FilePath, Encoding.UTF8))
                {
                    int i = line.IndexOf('=');
                    if (i > 0) map[line.Substring(0, i).Trim()] = line.Substring(i + 1).Trim();
                }
                s.HdrBg = Str(map, "HdrBg", s.HdrBg);
                s.HdrFg = Str(map, "HdrFg", s.HdrFg);
                s.BorderColor = Str(map, "BorderColor", s.BorderColor);
                s.BorderPt = Dbl(map, "BorderPt", s.BorderPt);
                s.BorderStyle = Int(map, "BorderStyle", s.BorderStyle);
                s.FontName = Str(map, "FontName", s.FontName);
                s.FontSize = Dbl(map, "FontSize", s.FontSize);
                s.HdrAlign = Int(map, "HdrAlign", s.HdrAlign);
                s.TextAlign = Int(map, "TextAlign", s.TextAlign);
                s.NumAlign = Int(map, "NumAlign", s.NumAlign);
                s.Zebra = Bool(map, "Zebra", s.Zebra);
                s.ZebraColor = Str(map, "ZebraColor", s.ZebraColor);
                s.ClearManual = Bool(map, "ClearManual", s.ClearManual);
                s.CellSpace = Dbl(map, "CellSpace", s.CellSpace);
                s.RepeatHeader = Bool(map, "RepeatHeader", s.RepeatHeader);
                s.NoRowBreak = Bool(map, "NoRowBreak", s.NoRowBreak);
                s.NumStyle = Str(map, "NumStyle", s.NumStyle);
                s.NumDec = Str(map, "NumDec", s.NumDec);
                s.NumThou = Bool(map, "NumThou", s.NumThou);
                s.NumAddDecToInt = Bool(map, "NumAddDecToInt", s.NumAddDecToInt);
                s.CalcOp = Str(map, "CalcOp", s.CalcOp);
                s.CalcDec = Str(map, "CalcDec", s.CalcDec);
                s.CalcFmtOn = Bool(map, "CalcFmtOn", s.CalcFmtOn);
                { string fk = Str(map, "CalcFmtKind", s.CalcFmtKind); if (fk == "bold" || fk == "highlight" || fk == "color" || fk == "fill") s.CalcFmtKind = fk; }
                s.TopMost = Bool(map, "TopMost", s.TopMost);
                s.SyncDec = Math.Max(0, Math.Min(5, Int(map, "SyncDec", s.SyncDec)));
                s.SyncSci = Bool(map, "SyncSci", s.SyncSci);
                s.SciDec = Math.Max(0, Math.Min(5, Int(map, "SciDec", s.SciDec)));
                s.ShowDialog = Bool(map, "ShowDialog", s.ShowDialog);
                { string un; if (map.TryGetValue("UnitNames", out un) && (un == "law" || un == "common")) s.UnitNames = un; }
                s.AddN = Math.Max(0, Math.Min(20, Int(map, "AddN", s.AddN)));
                { string mv; s.Manual = map.TryGetValue("Manual", out mv) ? mv : ""; }
                s.EqMode = Str(map, "EqMode", s.EqMode); if (s.EqMode != "glyph") s.EqMode = "fmt";
                s.RvScope = Str(map, "RvScope", s.RvScope);
                s.RvChem = Bool(map, "RvChem", s.RvChem); s.RvMath = Bool(map, "RvMath", s.RvMath); s.RvSym = Bool(map, "RvSym", s.RvSym);
                s.RvNum = Bool(map, "RvNum", s.RvNum); s.RvNumTbl = Bool(map, "RvNumTbl", s.RvNumTbl); s.RvIncl = Bool(map, "RvIncl", s.RvIncl);
                s.TmGloss = Bool(map, "TmGloss", s.TmGloss); s.TmOrth = Bool(map, "TmOrth", s.TmOrth); s.TmAbbr = Bool(map, "TmAbbr", s.TmAbbr);
                s.StLong = Bool(map, "StLong", s.StLong); s.StInf = Bool(map, "StInf", s.StInf); s.StDup = Bool(map, "StDup", s.StDup); s.StPunc = Bool(map, "StPunc", s.StPunc);
                s.StAi = Bool(map, "StAi", s.StAi); s.StConn = Bool(map, "StConn", s.StConn); s.StPron = Bool(map, "StPron", s.StPron);
                s.StLongN = Math.Max(10, Math.Min(120, Int(map, "StLongN", s.StLongN)));
                s.GrVi = Bool(map, "GrVi", s.GrVi); s.GrEn = Bool(map, "GrEn", s.GrEn);
                s.GrSpell = Bool(map, "GrSpell", s.GrSpell); s.GrAgree = Bool(map, "GrAgree", s.GrAgree);
                s.GrPunc = Bool(map, "GrPunc", s.GrPunc); s.GrStyle = Bool(map, "GrStyle", s.GrStyle); s.GrTerms = Bool(map, "GrTerms", s.GrTerms);
                s.GrLongN = Math.Max(15, Math.Min(80, Int(map, "GrLongN", s.GrLongN)));
                { string pp = Str(map, "PagePaper", s.PagePaper); if (pp == "A4" || pp == "Letter" || pp == "A3" || pp == "Legal") s.PagePaper = pp; }
                s.PgPortTop = Clamp(Dbl(map, "PgPortTop", s.PgPortTop), 0.2, 10.0);
                s.PgPortBot = Clamp(Dbl(map, "PgPortBot", s.PgPortBot), 0.2, 10.0);
                s.PgPortLeft = Clamp(Dbl(map, "PgPortLeft", s.PgPortLeft), 0.2, 10.0);
                s.PgPortRight = Clamp(Dbl(map, "PgPortRight", s.PgPortRight), 0.2, 10.0);
                s.PgLandTop = Clamp(Dbl(map, "PgLandTop", s.PgLandTop), 0.2, 10.0);
                s.PgLandBot = Clamp(Dbl(map, "PgLandBot", s.PgLandBot), 0.2, 10.0);
                s.PgLandLeft = Clamp(Dbl(map, "PgLandLeft", s.PgLandLeft), 0.2, 10.0);
                s.PgLandRight = Clamp(Dbl(map, "PgLandRight", s.PgLandRight), 0.2, 10.0);
                s.PgHeaderDist = Clamp(Dbl(map, "PgHeaderDist", s.PgHeaderDist), 0.2, 5.0);
                s.PgFooterDist = Clamp(Dbl(map, "PgFooterDist", s.PgFooterDist), 0.2, 5.0);
                s.ParaSizeTnr = Clamp(Dbl(map, "ParaSizeTnr", s.ParaSizeTnr), 6.0, 72.0);
                s.ParaSizeArial = Clamp(Dbl(map, "ParaSizeArial", s.ParaSizeArial), 6.0, 72.0);
                s.ParaLineSpace = Clamp(Dbl(map, "ParaLineSpace", s.ParaLineSpace), 0.8, 3.0);
                s.ParaSpaceBefore = Clamp(Dbl(map, "ParaSpaceBefore", s.ParaSpaceBefore), 0.0, 72.0);
                s.ParaSpaceAfter = Clamp(Dbl(map, "ParaSpaceAfter", s.ParaSpaceAfter), 0.0, 72.0);
                s.ParaFirstLineCm = Clamp(Dbl(map, "ParaFirstLineCm", s.ParaFirstLineCm), 0.0, 5.0);
                { string pa = Str(map, "ParaAlign", s.ParaAlign); if (pa == "left" || pa == "center" || pa == "right" || pa == "justify") s.ParaAlign = pa; }
                s.ParaHangingCm = Clamp(Dbl(map, "ParaHangingCm", s.ParaHangingCm), 0.0, 5.0);
                s.ApiProvider = Str(map, "ApiProvider", s.ApiProvider); s.ApiKeyEnc = Str(map, "ApiKeyEnc", s.ApiKeyEnc);
                s.ApiModel = Str(map, "ApiModel", s.ApiModel);
                s.AiTemp = Clamp(Dbl(map, "AiTemp", s.AiTemp), 0.0, 2.0);
                s.AiMaxTokens = Math.Max(64, Math.Min(8192, Int(map, "AiMaxTokens", s.AiMaxTokens)));
                s.AiTone = Str(map, "AiTone", s.AiTone);
                s.RwDoRewrite = Bool(map, "RwDoRewrite", s.RwDoRewrite); s.RwDoTranslate = Bool(map, "RwDoTranslate", s.RwDoTranslate);
                { string tl = Str(map, "RwTargetLang", s.RwTargetLang); if (Array.IndexOf(new[] { "vi", "en", "zh-CN", "ja", "ko", "ru", "fr", "de", "es", "id" }, tl) >= 0) s.RwTargetLang = tl; }
                { int h1 = (int)Dbl(map, "RwInH", s.RwInH), h2 = (int)Dbl(map, "RwOutH", s.RwOutH); s.RwInH = h1 < 60 ? 0 : Math.Min(800, h1); s.RwOutH = h2 < 60 ? 0 : Math.Min(800, h2); }
                s.RvNumStyle = Str(map, "RvNumStyle", s.RvNumStyle); if (s.RvNumStyle != "vn" && s.RvNumStyle != "us" && s.RvNumStyle != "si") s.RvNumStyle = "auto";
                s.EqStacked = Bool(map, "EqStacked", s.EqStacked);
                s.EqAlone = Bool(map, "EqAlone", s.EqAlone);
                int ver = Int(map, "Ver", 0);
                if (ver < 4) { s.RvChem = s.RvMath = s.RvSym = s.RvNum = s.RvNumTbl = true; }      // Rev0.4: mặc định chọn hết các kiểu quét
                if (ver < 5) { s.RepeatHeader = true; s.NoRowBreak = true; }
                if (ver < 6) s.RvNumStyle = "auto";
                s.RepeatHeader = true; s.NoRowBreak = true;                                          // Rev0.5: luôn lặp dòng tiêu đề và không ngắt hàng giữa trang
                s.Ver = 6;
            }
            catch { }
            return s;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Dir);
                var sb = new StringBuilder();
                Add(sb, "HdrBg", HdrBg); Add(sb, "HdrFg", HdrFg); Add(sb, "BorderColor", BorderColor);
                Add(sb, "BorderPt", BorderPt.ToString(CultureInfo.InvariantCulture));
                Add(sb, "BorderStyle", BorderStyle.ToString(CultureInfo.InvariantCulture));
                Add(sb, "FontName", FontName);
                Add(sb, "FontSize", FontSize.ToString(CultureInfo.InvariantCulture));
                Add(sb, "HdrAlign", HdrAlign.ToString(CultureInfo.InvariantCulture));
                Add(sb, "TextAlign", TextAlign.ToString(CultureInfo.InvariantCulture));
                Add(sb, "NumAlign", NumAlign.ToString(CultureInfo.InvariantCulture));
                Add(sb, "Zebra", Zebra ? "1" : "0"); Add(sb, "ZebraColor", ZebraColor);
                Add(sb, "ClearManual", ClearManual ? "1" : "0");
                Add(sb, "CellSpace", CellSpace.ToString(CultureInfo.InvariantCulture));
                Add(sb, "RepeatHeader", RepeatHeader ? "1" : "0");
                Add(sb, "NoRowBreak", NoRowBreak ? "1" : "0");
                Add(sb, "NumStyle", NumStyle); Add(sb, "NumDec", NumDec);
                Add(sb, "NumThou", NumThou ? "1" : "0");
                Add(sb, "NumAddDecToInt", NumAddDecToInt ? "1" : "0");
                Add(sb, "CalcOp", CalcOp); Add(sb, "CalcDec", CalcDec);
                Add(sb, "CalcFmtOn", CalcFmtOn ? "1" : "0"); Add(sb, "CalcFmtKind", CalcFmtKind);
                Add(sb, "TopMost", TopMost ? "1" : "0");
                Add(sb, "SyncDec", SyncDec.ToString(CultureInfo.InvariantCulture));
                Add(sb, "SyncSci", SyncSci ? "1" : "0");
                Add(sb, "SciDec", SciDec.ToString(CultureInfo.InvariantCulture));
                Add(sb, "ShowDialog", ShowDialog ? "1" : "0");
                Add(sb, "UnitNames", UnitNames);
                Add(sb, "AddN", AddN.ToString(CultureInfo.InvariantCulture));
                Add(sb, "Manual", Manual);
                Add(sb, "RvNumStyle", RvNumStyle); Add(sb, "EqStacked", EqStacked ? "1" : "0"); Add(sb, "EqMode", EqMode); Add(sb, "RvScope", RvScope);
                Add(sb, "RvChem", RvChem ? "1" : "0"); Add(sb, "RvMath", RvMath ? "1" : "0"); Add(sb, "RvSym", RvSym ? "1" : "0");
                Add(sb, "RvNum", RvNum ? "1" : "0"); Add(sb, "RvNumTbl", RvNumTbl ? "1" : "0"); Add(sb, "RvIncl", RvIncl ? "1" : "0");
                Add(sb, "TmGloss", TmGloss ? "1" : "0"); Add(sb, "TmOrth", TmOrth ? "1" : "0"); Add(sb, "TmAbbr", TmAbbr ? "1" : "0");
                Add(sb, "StLong", StLong ? "1" : "0"); Add(sb, "StInf", StInf ? "1" : "0"); Add(sb, "StDup", StDup ? "1" : "0"); Add(sb, "StPunc", StPunc ? "1" : "0");
                Add(sb, "StAi", StAi ? "1" : "0"); Add(sb, "StConn", StConn ? "1" : "0"); Add(sb, "StPron", StPron ? "1" : "0");
                Add(sb, "StLongN", StLongN.ToString(CultureInfo.InvariantCulture));
                Add(sb, "GrVi", GrVi ? "1" : "0"); Add(sb, "GrEn", GrEn ? "1" : "0");
                Add(sb, "GrSpell", GrSpell ? "1" : "0"); Add(sb, "GrAgree", GrAgree ? "1" : "0");
                Add(sb, "GrPunc", GrPunc ? "1" : "0"); Add(sb, "GrStyle", GrStyle ? "1" : "0"); Add(sb, "GrTerms", GrTerms ? "1" : "0");
                Add(sb, "GrLongN", GrLongN.ToString(CultureInfo.InvariantCulture));
                Add(sb, "PagePaper", PagePaper);
                Add(sb, "PgPortTop", PgPortTop.ToString(CultureInfo.InvariantCulture));
                Add(sb, "PgPortBot", PgPortBot.ToString(CultureInfo.InvariantCulture));
                Add(sb, "PgPortLeft", PgPortLeft.ToString(CultureInfo.InvariantCulture));
                Add(sb, "PgPortRight", PgPortRight.ToString(CultureInfo.InvariantCulture));
                Add(sb, "PgLandTop", PgLandTop.ToString(CultureInfo.InvariantCulture));
                Add(sb, "PgLandBot", PgLandBot.ToString(CultureInfo.InvariantCulture));
                Add(sb, "PgLandLeft", PgLandLeft.ToString(CultureInfo.InvariantCulture));
                Add(sb, "PgLandRight", PgLandRight.ToString(CultureInfo.InvariantCulture));
                Add(sb, "PgHeaderDist", PgHeaderDist.ToString(CultureInfo.InvariantCulture));
                Add(sb, "PgFooterDist", PgFooterDist.ToString(CultureInfo.InvariantCulture));
                Add(sb, "ParaSizeTnr", ParaSizeTnr.ToString(CultureInfo.InvariantCulture));
                Add(sb, "ParaSizeArial", ParaSizeArial.ToString(CultureInfo.InvariantCulture));
                Add(sb, "ParaLineSpace", ParaLineSpace.ToString(CultureInfo.InvariantCulture));
                Add(sb, "ParaSpaceBefore", ParaSpaceBefore.ToString(CultureInfo.InvariantCulture));
                Add(sb, "ParaSpaceAfter", ParaSpaceAfter.ToString(CultureInfo.InvariantCulture));
                Add(sb, "ParaFirstLineCm", ParaFirstLineCm.ToString(CultureInfo.InvariantCulture));
                Add(sb, "ParaAlign", ParaAlign);
                Add(sb, "ParaHangingCm", ParaHangingCm.ToString(CultureInfo.InvariantCulture));
                Add(sb, "ApiProvider", ApiProvider); Add(sb, "ApiKeyEnc", ApiKeyEnc);
                Add(sb, "ApiModel", ApiModel);
                Add(sb, "AiTemp", AiTemp.ToString(CultureInfo.InvariantCulture));
                Add(sb, "AiMaxTokens", AiMaxTokens.ToString(CultureInfo.InvariantCulture));
                Add(sb, "AiTone", AiTone);
                Add(sb, "RwDoRewrite", RwDoRewrite ? "1" : "0"); Add(sb, "RwDoTranslate", RwDoTranslate ? "1" : "0"); Add(sb, "RwTargetLang", RwTargetLang);
                Add(sb, "RwInH", RwInH.ToString(CultureInfo.InvariantCulture)); Add(sb, "RwOutH", RwOutH.ToString(CultureInfo.InvariantCulture));
                Add(sb, "Ver", Ver.ToString(CultureInfo.InvariantCulture));
                File.WriteAllText(FilePath, sb.ToString(), new UTF8Encoding(true));
            }
            catch { }
        }

        static double Clamp(double v, double lo, double hi) { return v < lo ? lo : (v > hi ? hi : v); }
        static void Add(StringBuilder sb, string k, string v) { sb.Append(k).Append('=').Append(v).Append("\r\n"); }
        static string Str(Dictionary<string, string> m, string k, string d) { string v; return m.TryGetValue(k, out v) && v.Length > 0 ? v : d; }
        static double Dbl(Dictionary<string, string> m, string k, double d)
        {
            string v; double r;
            return m.TryGetValue(k, out v) && double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out r) ? r : d;
        }
        static int Int(Dictionary<string, string> m, string k, int d)
        {
            string v; int r;
            return m.TryGetValue(k, out v) && int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out r) ? r : d;
        }
        static bool Bool(Dictionary<string, string> m, string k, bool d)
        {
            string v; return m.TryGetValue(k, out v) ? v == "1" : d;
        }

        /// <summary>"RRGGBB" sang số màu Word.</summary>
        public static int HexToWord(string hex)
        {
            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = Convert.ToInt32(hex.Substring(4, 2), 16);
            return WordMap.ColorToWord(r, g, b);
        }
    }

    /// <summary>Đặt thông số master style của bảng từ các nút và hộp trên Ribbon (không cần mở bảng điều khiển). Trả về thông báo cho người dùng.</summary>
    public static class TableParams
    {
        public const string Hint = " Áp dụng ngay cho bảng đang chọn. Mọi bảng hoặc Bảng này áp dụng toàn bộ thông số.";
        static readonly string[] AlName = { "trái", "giữa", "phải" };
        static readonly double[] Widths = { 0.25, 0.5, 0.75, 1, 1.5, 2.25, 3 };

        public static readonly string[] BorderPresets = { "0.25", "0.5", "0.75", "1.0", "1.5", "2.0", "3.0" };
        public static readonly string[] FontPresets = { "Times New Roman", "Arial", "Calibri", "Segoe UI" };
        public static readonly string[] SizePresets = { "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72" };
        public static readonly string[] NPresets = { "0", "1", "2", "3" };

        /// <summary>Các mục của menu màu (theo bố cục ô chọn màu của Word) cho một thông số: hdrBg, hdrFg, border, zebra.</summary>
        public static string PresetMenuXml(string target, string moreId) { return WordPalette.MenuXml(target, moreId); }

        /// <summary>Tăng, giảm một bậc: cỡ chữ (0,5 pt), khoảng cách đoạn (1 pt), độ dày viền (bậc kế tiếp của Word).</summary>
        public static string Step(string id, AppSettings s)
        {
            bool up = id.EndsWith("Inc");
            if (id.StartsWith("btnSz"))
            {
                s.FontSize = Math.Max(6, Math.Min(72, s.FontSize + (up ? 0.5 : -0.5)));
                return "Đã đặt cỡ chữ trong bảng: " + s.FontSize.ToString("0.##", CultureInfo.CurrentCulture) + " pt." + Hint;
            }
            if (id.StartsWith("btnSp"))
            {
                s.CellSpace = Math.Max(0, Math.Min(12, s.CellSpace + (up ? 1 : -1)));
                return "Đã đặt khoảng cách trước và sau đoạn trong ô: " + s.CellSpace.ToString("0.##", CultureInfo.CurrentCulture) + " pt." + Hint;
            }
            int k = 0; double best = 1e9;
            for (int i = 0; i < Widths.Length; i++) { double d = Math.Abs(Widths[i] - s.BorderPt); if (d < best) { best = d; k = i; } }
            k = Math.Max(0, Math.Min(Widths.Length - 1, k + (up ? 1 : -1)));
            s.BorderPt = Widths[k];
            return "Đã đặt độ dày viền: " + Widths[k].ToString("0.##", CultureInfo.CurrentCulture) + " pt." + Hint;
        }

        public static string SetBorder(AppSettings s, string text)
        {
            double v; if (!Num(text, out v) || v < 0.25 || v > 6.0) return "Độ dày viền phải từ 0,25 đến 6 pt (bước 0,25 pt).";
            double step = Math.Round(v * 4.0, MidpointRounding.AwayFromZero) / 4.0;
            step = Math.Max(0.25, Math.Min(6.0, step));
            s.BorderPt = step;
            bool rounded = Math.Abs(step - v) > 1e-6;
            return "Đã đặt độ dày viền: " + step.ToString("0.##", CultureInfo.CurrentCulture) + " pt" + (rounded ? " (đã làm tròn theo bước 0,25 pt)." : ".") + Hint;
        }

        /// <summary>Xử lý một mục menu theo mã. Trả về null nếu mã không thuộc nhóm này.</summary>
        public static string Apply(string id, AppSettings s)
        {
            switch (id)
            {
                case "hdrAlL": s.HdrAlign = 0; return "Đã đặt căn chữ dòng tiêu đề: trái." + Hint;
                case "hdrAlC": s.HdrAlign = 1; return "Đã đặt căn chữ dòng tiêu đề: giữa." + Hint;
                case "hdrAlR": s.HdrAlign = 2; return "Đã đặt căn chữ dòng tiêu đề: phải." + Hint;
                case "bdStSolid": s.BorderStyle = 0; return "Đã đặt kiểu viền: liền." + Hint;
                case "bdStDash": s.BorderStyle = 1; return "Đã đặt kiểu viền: nét đứt." + Hint;
                case "bdStDot": s.BorderStyle = 2; return "Đã đặt kiểu viền: nét chấm." + Hint;
                case "zebraToggle": s.Zebra = !s.Zebra; return (s.Zebra ? "Đã bật tô xen kẽ dòng chẵn." : "Đã tắt tô xen kẽ.") + Hint;
                case "txtAlL": s.TextAlign = 0; return "Đã đặt căn ô chữ: trái." + Hint;
                case "txtAlJ": s.TextAlign = 3; return "Đã đặt căn ô chữ: hai bên." + Hint;
                case "numAlL": s.NumAlign = 0; return "Đã đặt căn ô số: trái." + Hint;
                case "numAlC": s.NumAlign = 1; return "Đã đặt căn ô số: giữa." + Hint;
                case "numAlR": s.NumAlign = 2; return "Đã đặt căn ô số: phải." + Hint;
            }
            if (id.StartsWith("bdW"))
            {
                int i; if (int.TryParse(id.Substring(3), out i) && i >= 0 && i < Widths.Length)
                {
                    s.BorderPt = Widths[i];
                    return "Đã đặt độ dày viền: " + Widths[i].ToString("0.##", CultureInfo.CurrentCulture) + " pt." + Hint;
                }
            }
            return null;
        }

        /// <summary>Đặt màu theo mã nút: hdrBg, hdrFg, border, zebra. hex dạng RRGGBB.</summary>
        public static string SetColor(string id, AppSettings s, string hex)
        {
            hex = hex.ToUpperInvariant();
            switch (id)
            {
                case "hdrBg": s.HdrBg = hex; return "Đã đặt màu nền dòng tiêu đề #" + hex + "." + Hint;
                case "hdrFg": s.HdrFg = hex; return "Đã đặt màu chữ dòng tiêu đề #" + hex + "." + Hint;
                case "border": s.BorderColor = hex; return "Đã đặt màu viền #" + hex + "." + Hint;
                case "zebra": s.ZebraColor = hex; return "Đã đặt màu dòng chẵn #" + hex + "." + Hint;
            }
            return null;
        }
        public static string GetColor(string id, AppSettings s)
        {
            switch (id) { case "hdrBg": return s.HdrBg; case "hdrFg": return s.HdrFg; case "border": return s.BorderColor; default: return s.ZebraColor; }
        }

        static bool Num(string t, out double v)
        {
            t = (t ?? "").Trim().Replace(',', '.');
            return double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out v);
        }

        public static string SetFont(AppSettings s, string text)
        {
            string t = (text ?? "").Trim();
            if (t.Length == 0) return "Hãy nhập tên font, ví dụ Times New Roman.";
            s.FontName = t; return "Đã đặt font chữ trong bảng: " + t + "." + Hint;
        }
        public static string SetSize(AppSettings s, string text)
        {
            double v; if (!Num(text, out v) || v < 6 || v > 72) return "Cỡ chữ phải từ 6 đến 72 pt.";
            s.FontSize = v; return "Đã đặt cỡ chữ trong bảng: " + v.ToString("0.##", CultureInfo.CurrentCulture) + " pt." + Hint;
        }
        public static string SetSpace(AppSettings s, string text)
        {
            double v; if (!Num(text, out v) || v < 0 || v > 12) return "Khoảng cách đoạn phải từ 0 đến 12 pt.";
            s.CellSpace = v; return "Đã đặt khoảng cách trước và sau đoạn trong ô: " + v.ToString("0.##", CultureInfo.CurrentCulture) + " pt." + Hint;
        }
    }
}
