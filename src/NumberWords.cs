// Trợ lý Kỹ thuật - Rev0.7: đọc số, tức viết số bằng chữ tiếng Việt. Làm việc trên chuỗi, không gọi Word.
// Quy ước đọc (nêu rõ để người dùng biết): hai mươi mốt, hai mươi lăm, mười lăm, một trăm linh năm, không trăm linh một (khi có hàng cao hơn),
// bốn (không dùng "tư"), phần thập phân đọc từng chữ số sau "phẩy" (3,05 = ba phẩy không năm), số âm đọc "âm", số kỹ thuật a × 10ⁿ đọc "a nhân mười mũ n".
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace TroLyKyThuat
{
    public static class NumberWords
    {
        static readonly string[] Digit = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };
        public const int MaxDigits = 30;          // tối đa 30 chữ số phần nguyên

        /// <summary>Đọc một số nguyên không dấu viết bằng chữ số (bỏ số 0 đứng đầu). Trả null nếu quá dài.</summary>
        public static string Integer(string digits)
        {
            if (string.IsNullOrEmpty(digits)) return "không";
            digits = digits.TrimStart('0');
            if (digits.Length == 0) return "không";
            if (digits.Length > MaxDigits) return null;
            return ReadDigits(digits, false);
        }

        // Đọc dãy chữ số (đã bỏ số 0 đầu khi leading). Mỗi 9 chữ số là một "tỷ": phần cao đọc như số thường rồi thêm "tỷ".
        static string ReadDigits(string d, bool hasHigher)
        {
            if (d.Length > 9)
            {
                string hi = d.Substring(0, d.Length - 9), lo = d.Substring(d.Length - 9);
                string h = ReadDigits(hi.TrimStart('0').Length == 0 ? "0" : hi.TrimStart('0'), hasHigher);
                string r = h + " tỷ";
                if (lo.Trim('0').Length > 0) r += " " + ReadDigits(lo, true);
                return r;
            }
            d = d.PadLeft(9, '0');
            string[] unit = { "triệu", "nghìn", "" };
            var parts = new List<string>();
            bool higher = hasHigher;
            for (int g = 0; g < 3; g++)
            {
                int h = d[g * 3] - '0', t = d[g * 3 + 1] - '0', u = d[g * 3 + 2] - '0';
                if (h == 0 && t == 0 && u == 0) continue;
                string w = Group(h, t, u, higher);
                if (unit[g].Length > 0) w += " " + unit[g];
                parts.Add(w);
                higher = true;
            }
            return parts.Count == 0 ? "không" : string.Join(" ", parts.ToArray());
        }

        // Một nhóm ba chữ số. full = đã có hàng cao hơn: đọc "không trăm" và "linh" đầy đủ.
        static string Group(int h, int t, int u, bool full)
        {
            var sb = new StringBuilder();
            if (h > 0) sb.Append(Digit[h]).Append(" trăm");
            else if (full) sb.Append("không trăm");
            if (t == 0)
            {
                if (u > 0) { if (sb.Length > 0) sb.Append(" linh "); sb.Append(Digit[u]); }
            }
            else if (t == 1)
            {
                if (sb.Length > 0) sb.Append(' ');
                sb.Append("mười");
                if (u == 5) sb.Append(" lăm"); else if (u > 0) sb.Append(' ').Append(Digit[u]);
            }
            else
            {
                if (sb.Length > 0) sb.Append(' ');
                sb.Append(Digit[t]).Append(" mươi");
                if (u == 1) sb.Append(" mốt"); else if (u == 5) sb.Append(" lăm"); else if (u > 0) sb.Append(' ').Append(Digit[u]);
            }
            return sb.ToString();
        }

        /// <summary>Đọc chuỗi số đã chuẩn hóa (chữ số, một dấu '.' thập phân, không dấu). neg = số âm. Trả null nếu quá dài.</summary>
        public static string FromNorm(string norm, bool neg)
        {
            if (string.IsNullOrEmpty(norm)) return null;
            string ip = norm, fp = "";
            int dot = norm.IndexOf('.');
            if (dot >= 0) { ip = norm.Substring(0, dot); fp = norm.Substring(dot + 1); }
            if (ip.Length == 0) ip = "0";
            string w = Integer(ip);
            if (w == null) return null;
            if (fp.Length > 0)
            {
                var sb = new StringBuilder(w).Append(" phẩy");
                foreach (char c in fp) { if (c < '0' || c > '9') return null; sb.Append(' ').Append(Digit[c - '0']); }
                w = sb.ToString();
            }
            bool zero = true; foreach (char c in norm) if (c >= '1' && c <= '9') { zero = false; break; }
            return (neg && !zero ? "âm " : "") + w;
        }

        /// <summary>Đọc một mục số trong văn bản kèm đơn vị đứng ngay sau (nếu có): "20 kPa" thành "hai mươi kilôpascan". uend = vị trí hết đơn vị (hết số nếu không có đơn vị).
        /// Số có ký hiệu tiền tệ đứng trước ($100) được đọc kèm tên tiền. Trả null nếu không đọc được số.</summary>
        public static string ReadItem(string text, NumItem it, char hint, bool lawNames, out int uend)
        {
            uend = it.E;
            string w = Read(text.Substring(it.S, it.E - it.S), hint);
            if (w == null) return null;
            UnitSpan u = UnitWords.Scan(text, it.E, lawNames);
            if (u != null) { w += " " + u.Words; uend = u.End; }
            else { string pre = UnitWords.Prefix(text, it.S); if (pre != null) w += " " + pre; }
            return w;
        }

        /// <summary>Đọc số theo chữ trong văn bản (ví dụ "1.234,5", "-12,5", "1 234", "3.14", "2,3 × 10⁶"). hint = dấu thập phân của vùng để hiểu số mơ hồ 1.234.
        /// Trả null nếu không đọc được.</summary>
        public static string Read(string raw, char hint)
        {
            if (raw == null) return null;
            // số kiểu kỹ thuật a × 10ⁿ
            var sci = NumberEdit.FindSci(raw.Trim());
            if (sci.Count == 1 && sci[0].Index == 0 && sci[0].Length == raw.Trim().Length)
            {
                string a = Read(sci[0].Groups["m"].Value, hint);
                if (a == null) return null;
                int e = NumberEdit.ParseExp(sci[0].Groups["e"].Value);
                return a + " nhân mười mũ " + (e < 0 ? "âm " : "") + Integer(Math.Abs(e).ToString(CultureInfo.InvariantCulture));
            }
            ParsedNumber p = NumberUtil.Parse(raw, hint);
            if (p == null || p.Norm == null) return null;
            return FromNorm(p.Norm, p.D < 0 || raw.TrimStart().StartsWith("-") || raw.TrimStart().StartsWith("−"));
        }
    }
}
