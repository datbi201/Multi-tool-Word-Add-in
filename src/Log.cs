// Trợ lý Kỹ thuật - Rev0.2: ghi nhật ký chẩn đoán (nhỏ, không bao giờ làm lỗi chương trình).
using System;

namespace TroLyKyThuat
{
    public static class Log
    {
        static void To(string dir, string m)
        {
            try
            {
                System.IO.Directory.CreateDirectory(dir);
                System.IO.File.AppendAllText(System.IO.Path.Combine(dir, "log.txt"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + m + "\r\n", new System.Text.UTF8Encoding(true));
            }
            catch { }
        }
        public static void W(string m)
        {
            To(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "TroLyKyThuat"), m);   // luôn ghi được
            try { To(AppSettings.Dir, m); } catch { }
        }
    }
}
