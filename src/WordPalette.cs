// Trợ lý Kỹ thuật - Rev0.7: bảng màu giống hệt ô chọn màu của Word: Tự động, Màu giao diện (Theme Colors: 10 cột x 6 hàng), Màu chuẩn (Standard Colors), Màu khác.
// Màu sáng hơn, tối hơn tính bằng công thức của Office (đổi sang HSL, nhân và cộng độ sáng: lumMod, lumOff), mã màu lưu vào tài liệu là RGB nên mở ở máy không cài add-in vẫn đúng màu.
// Bộ màu giao diện lấy từ chính tài liệu đang mở (DocumentTheme) khi đọc được, ngược lại dùng bộ màu mặc định của Office.
using System;
using System.Globalization;

namespace TroLyKyThuat
{
    public static class WordPalette
    {
        // Thứ tự cột của Word: Nền 1 (Light 1), Chữ 1 (Dark 1), Nền 2 (Light 2), Chữ 2 (Dark 2), Nhấn 1 đến 6
        public static readonly string[] DefaultTheme = { "FFFFFF", "000000", "E8E8E8", "0E2841", "156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E" };
        public static readonly string[] ThemeNames = { "Nền 1", "Chữ 1", "Nền 2", "Chữ 2", "Nhấn 1", "Nhấn 2", "Nhấn 3", "Nhấn 4", "Nhấn 5", "Nhấn 6" };
        // Màu chuẩn của Word (hàng Standard Colors)
        public static readonly string[] Standard = { "C00000", "FF0000", "FFC000", "FFFF00", "92D050", "00B050", "00B0F0", "0070C0", "002060", "7030A0" };
        public static readonly string[] StandardNames = { "Đỏ đậm", "Đỏ", "Cam", "Vàng", "Xanh lá nhạt", "Xanh lá", "Xanh da trời nhạt", "Xanh da trời", "Xanh navy", "Tím" };

        static string[] theme = (string[])DefaultTheme.Clone();
        public static string[] Theme { get { return theme; } }
        public static void SetTheme(string[] t) { if (t != null && t.Length == 10) theme = t; }
        public static void ResetTheme() { theme = (string[])DefaultTheme.Clone(); }

        // Bộ số nhân độ sáng của từng hàng biến thể (hàng 1 đến 5): { lumMod, lumOff } và tên
        struct Var { public double Mod, Off; public string Name; public Var(double m, double o, string n) { Mod = m; Off = o; Name = n; } }
        static Var[] Variants(string hex, int col)
        {
            double h, s, l; ToHsl(hex, out h, out s, out l);
            if (col == 0 && l > 0.97)     // trắng: tối hơn 5, 15, 25, 35, 50 %
                return new[] { new Var(0.95, 0, "tối hơn 5%"), new Var(0.85, 0, "tối hơn 15%"), new Var(0.75, 0, "tối hơn 25%"), new Var(0.65, 0, "tối hơn 35%"), new Var(0.50, 0, "tối hơn 50%") };
            if (l < 0.03)                 // đen: sáng hơn 50, 35, 25, 15, 5 %
                return new[] { new Var(0.50, 0.50, "sáng hơn 50%"), new Var(0.65, 0.35, "sáng hơn 35%"), new Var(0.75, 0.25, "sáng hơn 25%"), new Var(0.85, 0.15, "sáng hơn 15%"), new Var(0.95, 0.05, "sáng hơn 5%") };
            if (l > 0.85)                 // xám, be rất sáng: tối hơn 10, 25, 50, 75, 90 %
                return new[] { new Var(0.90, 0, "tối hơn 10%"), new Var(0.75, 0, "tối hơn 25%"), new Var(0.50, 0, "tối hơn 50%"), new Var(0.25, 0, "tối hơn 75%"), new Var(0.10, 0, "tối hơn 90%") };
            if (l < 0.20)                 // màu rất tối: sáng hơn 90, 75, 50, 25 %, rồi tối hơn 25 %
                return new[] { new Var(0.10, 0.90, "sáng hơn 90%"), new Var(0.25, 0.75, "sáng hơn 75%"), new Var(0.50, 0.50, "sáng hơn 50%"), new Var(0.75, 0.25, "sáng hơn 25%"), new Var(0.75, 0, "tối hơn 25%") };
            return new[] { new Var(0.20, 0.80, "sáng hơn 80%"), new Var(0.40, 0.60, "sáng hơn 60%"), new Var(0.60, 0.40, "sáng hơn 40%"), new Var(0.75, 0, "tối hơn 25%"), new Var(0.50, 0, "tối hơn 50%") };
        }

        public const int ThemeCount = 60, StdCount = 10;

        /// <summary>Mã màu (RRGGBB) của ô thứ index trong bảng Màu giao diện, xếp theo hàng: hàng 0 là 10 màu gốc, hàng 1 đến 5 là các biến thể.</summary>
        public static string ThemeHex(int index)
        {
            if (index < 0 || index >= ThemeCount) return null;
            int row = index / 10, col = index % 10;
            string b = theme[col];
            if (row == 0) return b;
            Var v = Variants(b, col)[row - 1];
            return Tint(b, v.Mod, v.Off);
        }

        /// <summary>Tên đọc được của ô: ví dụ "Nhấn 1, sáng hơn 80% (DAE3F3)".</summary>
        public static string ThemeLabel(int index)
        {
            if (index < 0 || index >= ThemeCount) return "";
            int row = index / 10, col = index % 10;
            string hex = ThemeHex(index);
            if (row == 0) return ThemeNames[col] + " (" + hex + ")";
            return ThemeNames[col] + ", " + Variants(theme[col], col)[row - 1].Name + " (" + hex + ")";
        }

        public static string StdHex(int i) { return i >= 0 && i < StdCount ? Standard[i] : null; }
        public static string StdLabel(int i) { return i >= 0 && i < StdCount ? StandardNames[i] + " (" + Standard[i] + ")" : ""; }

        /// <summary>Mã màu của mục trong một bảng chọn: gallery "gT_đích" (Màu giao diện), "gS_đích" (Màu chuẩn). Trả null nếu mã lạ.</summary>
        public static string Resolve(string galleryId, int index, out string target)
        {
            target = null;
            if (galleryId == null) return null;
            string[] parts = galleryId.Split('_');
            if (parts.Length != 2) return null;
            target = parts[1];
            if (parts[0] == "gT") return ThemeHex(index);
            if (parts[0] == "gS") return StdHex(index);
            target = null; return null;
        }

        /// <summary>Màu "Tự động" (chữ, viền) là đen; "Không màu" (nền) là trắng.</summary>
        public static string AutoHex(string target) { return target == "hdrFg" || target == "border" ? "000000" : "FFFFFF"; }
        public static string AutoLabel(string target) { return target == "hdrFg" || target == "border" ? "Tự động" : "Tự động (không màu)"; }

        /// <summary>Thông số bố cục bảng màu thống nhất kiểu Word hiện đại (dùng chung cho mọi ô chọn màu).</summary>
        public const string ThemeSectionLabel = "Màu giao diện";
        public const string StdSectionLabel = "Màu chuẩn";
        public const string MoreColorsLabel = "Thêm màu...";

        /// <summary>XML nút màu duy nhất trên Ribbon (bấm mở trực tiếp bảng màu thống nhất, không qua menu lồng).</summary>
        public static string ButtonXml(string id, string screentip, string supertip)
        {
            return "<button id=\"" + id + "\" showLabel=\"false\" getImage=\"GetImg\" onAction=\"OnBtn\" screentip=\"" + screentip + "\" supertip=\"" + supertip + "\" />";
        }

        /// <summary>XML dự phòng nếu dựng dạng menu.</summary>
        public static string MenuXml(string target, string moreId)
        {
            return "<button id=\"auto_" + target + "\" label=\"" + AutoLabel(target) + "\" getImage=\"GetImg\" onAction=\"OnBtn\" />"
                + "<menuSeparator id=\"sepT_" + target + "\" title=\"" + ThemeSectionLabel + "\" />"
                + "<gallery id=\"gT_" + target + "\" columns=\"10\" rows=\"6\" itemWidth=\"22\" itemHeight=\"22\" showItemLabel=\"false\" getItemCount=\"GalCount\" getItemImage=\"GalImg\" getItemLabel=\"GalLabel\" getItemScreentip=\"GalLabel\" onAction=\"OnGal\" />"
                + "<menuSeparator id=\"sepS_" + target + "\" title=\"" + StdSectionLabel + "\" />"
                + "<gallery id=\"gS_" + target + "\" columns=\"10\" rows=\"1\" itemWidth=\"22\" itemHeight=\"22\" showItemLabel=\"false\" getItemCount=\"GalCount\" getItemImage=\"GalImg\" getItemLabel=\"GalLabel\" getItemScreentip=\"GalLabel\" onAction=\"OnGal\" />"
                + "<menuSeparator id=\"sepM_" + target + "\" />"
                + "<button id=\"" + moreId + "\" label=\"" + MoreColorsLabel + "\" onAction=\"OnBtn\" />";
        }

        // ---------------------------------------------------------------- công thức màu của Office (HSL)
        public static string Tint(string hex, double lumMod, double lumOff)
        {
            double h, s, l; ToHsl(hex, out h, out s, out l);
            l = Math.Max(0, Math.Min(1, l * lumMod + lumOff));
            return FromHsl(h, s, l);
        }

        static void ToHsl(string hex, out double h, out double s, out double l)
        {
            double r = Convert.ToInt32(hex.Substring(0, 2), 16) / 255.0, g = Convert.ToInt32(hex.Substring(2, 2), 16) / 255.0, b = Convert.ToInt32(hex.Substring(4, 2), 16) / 255.0;
            double mx = Math.Max(r, Math.Max(g, b)), mn = Math.Min(r, Math.Min(g, b)), d = mx - mn;
            l = (mx + mn) / 2; h = 0; s = 0;
            if (d < 1e-9) return;
            s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
            if (mx == r) h = (g - b) / d + (g < b ? 6 : 0);
            else if (mx == g) h = (b - r) / d + 2;
            else h = (r - g) / d + 4;
            h /= 6;
        }

        static double Hue(double p, double q, double t)
        {
            if (t < 0) t += 1; if (t > 1) t -= 1;
            if (t < 1.0 / 6) return p + (q - p) * 6 * t;
            if (t < 0.5) return q;
            if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
            return p;
        }

        // Làm tròn nửa xuống như Word: 127,5 thành 127 (nên Nền 1 tối hơn 50% là 7F7F7F, không phải 808080)
        static int Chan(double v) { return (int)Math.Max(0, Math.Min(255, Math.Ceiling(v * 255 - 0.5 - 1e-9))); }

        static string FromHsl(double h, double s, double l)
        {
            double r, g, b;
            if (s < 1e-9) { r = g = b = l; }
            else
            {
                double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
                r = Hue(p, q, h + 1.0 / 3); g = Hue(p, q, h); b = Hue(p, q, h - 1.0 / 3);
            }
            return string.Format(CultureInfo.InvariantCulture, "{0:X2}{1:X2}{2:X2}", Chan(r), Chan(g), Chan(b));
        }

        /// <summary>Đọc bộ màu giao diện của tài liệu đang mở (DocumentTheme.ThemeColorScheme). Không đọc được thì giữ bộ đang dùng. Trả true nếu đọc được.</summary>
        public static bool LoadFrom(object wordApp)
        {
            try
            {
                dynamic app = wordApp;
                if ((int)app.Documents.Count == 0) return false;
                dynamic cs = app.ActiveDocument.DocumentTheme.ThemeColorScheme;
                int[] idx = { 2, 1, 4, 3, 5, 6, 7, 8, 9, 10 };      // Light1, Dark1, Light2, Dark2, Accent1..6
                var t = new string[10];
                for (int i = 0; i < 10; i++)
                {
                    int bgr = (int)cs.Colors(idx[i]).RGB;
                    t[i] = string.Format(CultureInfo.InvariantCulture, "{0:X2}{1:X2}{2:X2}", bgr & 255, (bgr >> 8) & 255, (bgr >> 16) & 255);
                }
                theme = t; return true;
            }
            catch (Exception ex) { Log.W("Doc bo mau giao dien loi: " + ex.Message); return false; }
        }
    }
}
