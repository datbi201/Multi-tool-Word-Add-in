// Công cụ xem trước icon (Rev0.9.5.1): vẽ mọi icon trong Icons.cs ra một ảnh PNG để người và Claude kiểm tra bằng mắt.
// Mỗi ô gồm: icon 32 px thật, bản phóng 96 px, bản thu nhỏ 16 px (kích thước nút nhỏ trên Ribbon), mã icon bên dưới.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Text.RegularExpressions;
using TroLyKyThuat;

static class IconSheet
{
    static int Main(string[] args)
    {
        string src = args.Length > 0 ? args[0] : @"src\Icons.cs";
        string outPath = args.Length > 1 ? args[1] : @"icons-sheet.png";
        var ids = new List<string>();
        foreach (Match m in Regex.Matches(File.ReadAllText(src), "case \"([A-Za-z0-9_:]+)\":"))
        {
            string id = m.Groups[1].Value;
            if (!ids.Contains(id)) ids.Add(id);
        }
        if (!ids.Contains("btnNumStyle:us")) ids.Add("btnNumStyle:us");
        if (!ids.Contains("gs:FF0000")) ids.Add("gs:FF0000");
        if (!ids.Contains("sw:0070C0")) ids.Add("sw:0070C0");
        ids.Add("khongco_xyz");   // mã lạ: phải có icon dự phòng

        int cols = 6, cellW = 150, cellH = 150, pad = 8;
        int rows = (ids.Count + cols - 1) / cols;
        using (var sheet = new Bitmap(cols * cellW, rows * cellH, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
        using (Graphics g = Graphics.FromImage(sheet))
        using (var font = new Font("Segoe UI", 8f))
        {
            g.Clear(Color.White);
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            for (int i = 0; i < ids.Count; i++)
            {
                int cx = (i % cols) * cellW, cy = (i / cols) * cellH;
                using (var p = new Pen(Color.FromArgb(225, 225, 225))) g.DrawRectangle(p, cx, cy, cellW - 1, cellH - 1);
                Bitmap ic = null;
                try { ic = Icons.Draw(ids[i]); } catch (Exception ex) { g.DrawString("LỖI " + ex.Message, font, Brushes.Red, cx + pad, cy + pad); }
                if (ic != null)
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.DrawImage(ic, new Rectangle(cx + pad, cy + pad, 96, 96));        // phóng 3 lần
                    g.DrawImageUnscaled(ic, cx + pad + 104, cy + pad);                  // 32 px thật
                    g.DrawImage(ic, new Rectangle(cx + pad + 108, cy + pad + 44, 16, 16)); // 16 px
                    ic.Dispose();
                }
                g.DrawString(ids[i], font, Brushes.Black, cx + pad, cy + pad + 108);
            }
            sheet.Save(outPath, System.Drawing.Imaging.ImageFormat.Png);
        }
        Console.WriteLine("Da ve " + ids.Count + " icon vao " + outPath);
        return 0;
    }
}
