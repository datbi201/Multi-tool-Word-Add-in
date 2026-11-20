// Trợ lý Kỹ thuật - Rev0.4: mã hóa khóa API bằng DPAPI của Windows (chỉ giải mã được trên đúng máy và đúng tài khoản Windows này).
using System;
using System.Security.Cryptography;
using System.Text;

namespace TroLyKyThuat
{
    public static class Secret
    {
        static readonly byte[] Entropy = Encoding.UTF8.GetBytes("TroLyKyThuat.v4");

        public static string Protect(string plain)
        {
            if (string.IsNullOrEmpty(plain)) return "";
            byte[] b = ProtectedData.Protect(Encoding.UTF8.GetBytes(plain), Entropy, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(b);
        }

        /// <summary>Trả về chuỗi rỗng nếu không giải mã được (khác máy, khác tài khoản, tệp hỏng).</summary>
        public static string Unprotect(string enc)
        {
            if (string.IsNullOrEmpty(enc)) return "";
            try { return Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(enc), Entropy, DataProtectionScope.CurrentUser)); }
            catch { return ""; }
        }

        /// <summary>Hiện khóa dưới dạng che bớt: ba ký tự đầu và bốn ký tự cuối.</summary>
        public static string Mask(string key)
        {
            if (string.IsNullOrEmpty(key)) return "";
            if (key.Length <= 10) return new string('•', key.Length);
            return key.Substring(0, 3) + new string('•', 6) + key.Substring(key.Length - 4);
        }
    }
}
