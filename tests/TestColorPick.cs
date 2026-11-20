// Rev0.9.5: kiểm thử hộp chọn màu. Tái hiện lỗi Rev0.9.3 và 0.9.4: chọn một ô màu xong kết quả bị đổi thành Hủy
// (đóng hộp làm phát sinh sự kiện mất tiêu điểm, sự kiện đó ghi đè DialogResult OK thành Cancel).
// Kiểm thử mở hộp ở chế độ MODAL như trong Word, bấm ô màu thật bằng bộ đếm giờ, và phát sinh sự kiện mất tiêu điểm đúng lúc
// hộp đang đóng (FormClosing). Xác định, không phụ thuộc cửa sổ có đang được kích hoạt hay không.
using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using TroLyKyThuat;

static class TCP
{
    static int fails = 0, total = 0;
    static void Eq(object a, object b, string name) { total++; bool ok = Equals(a, b); if (!ok) fails++; Console.WriteLine((ok ? "PASS " : "FAIL ") + name + (ok ? "" : "  (nhan " + a + ", mong " + b + ")")); }

    const BindingFlags NP = BindingFlags.NonPublic | BindingFlags.Instance;

    static Rectangle[] Rects(ColorPickerForm f, string field)
    {
        return (Rectangle[])typeof(ColorPickerForm).GetField(field, NP).GetValue(f);
    }

    static void Deact(ColorPickerForm f)
    {
        typeof(Form).GetMethod("OnDeactivate", NP).Invoke(f, new object[] { EventArgs.Empty });
    }

    static void Click(ColorPickerForm f, Point p)
    {
        typeof(Control).GetMethod("OnMouseDown", NP).Invoke(f, new object[] { new MouseEventArgs(MouseButtons.Left, 1, p.X, p.Y, 0) });
    }

    static Point Center(Rectangle r) { return new Point(r.X + r.Width / 2, r.Y + r.Height / 2); }

    /// <summary>Mở hộp modal. what: "theme", "std", "auto", "outside" (bấm chỗ trống) hoặc "none" (không bấm gì).
    /// deactOnClosing: phát sinh mất tiêu điểm đúng lúc hộp đang đóng. deactNow: phát sinh mất tiêu điểm khi hộp còn mở, chưa chọn gì.</summary>
    static ColorPickerForm RunModal(string what, int index, bool deactOnClosing, bool deactNow)
    {
        var f = new ColorPickerForm("hdrBg", "FFFFFF", new int[16]);
        bool fired = false;   // chỉ phát sinh một lần (mã cũ sẽ đóng lại và gây lặp vô hạn nếu phát sinh mỗi lần đóng)
        if (deactOnClosing) f.FormClosing += delegate { if (!fired) { fired = true; Deact(f); } };
        var t = new Timer { Interval = 150 };
        t.Tick += delegate
        {
            t.Stop();
            if (what == "theme") Click(f, Center(Rects(f, "themeRects")[index]));
            else if (what == "std") Click(f, Center(Rects(f, "stdRects")[index]));
            else if (what == "auto") Click(f, Center((Rectangle)typeof(ColorPickerForm).GetField("autoRect", NP).GetValue(f)));
            else if (what == "outside") Click(f, new Point(1, 1));
            if (deactNow) Deact(f);
        };
        var guard = new Timer { Interval = 3000 };   // an toàn: không để kiểm thử treo
        guard.Tick += delegate { guard.Stop(); if (!f.IsDisposed) f.DialogResult = DialogResult.Abort; };
        t.Start(); guard.Start();
        f.ShowDialog();
        t.Dispose(); guard.Dispose();
        return f;
    }

    [STAThread]
    static int Main()
    {
        // 1) Chọn một ô Màu giao diện; hộp mất tiêu điểm đúng lúc đóng (như trong Word): màu giữ nguyên, kết quả vẫn OK
        var f = RunModal("theme", 13, true, false);
        Eq(f.SelectedHex, WordPalette.ThemeHex(13), "chọn ô Màu giao diện: có mã màu");
        Eq(f.DialogResult, DialogResult.OK, "chọn ô Màu giao diện: kết quả OK, không bị đổi thành Hủy khi mất tiêu điểm lúc đóng");
        f.Dispose();

        // 2) Màu chuẩn
        f = RunModal("std", 2, true, false);
        Eq(f.SelectedHex, WordPalette.StdHex(2), "chọn ô Màu chuẩn: có mã màu");
        Eq(f.DialogResult, DialogResult.OK, "chọn ô Màu chuẩn: kết quả OK");
        f.Dispose();

        // 3) Tự động
        f = RunModal("auto", 0, true, false);
        Eq(f.SelectedHex, WordPalette.AutoHex("hdrBg"), "chọn Tự động: có mã màu");
        Eq(f.DialogResult, DialogResult.OK, "chọn Tự động: kết quả OK");
        f.Dispose();

        // 4) Không có sự kiện mất tiêu điểm thì chọn màu vẫn bình thường
        f = RunModal("theme", 5, false, false);
        Eq(f.SelectedHex, WordPalette.ThemeHex(5), "chọn màu khi không mất tiêu điểm: có mã màu");
        Eq(f.DialogResult, DialogResult.OK, "chọn màu khi không mất tiêu điểm: kết quả OK");
        f.Dispose();

        // 5) Bấm ra ngoài vùng ô màu: không chọn gì (đóng bằng bộ đếm an toàn)
        f = RunModal("outside", 0, false, false);
        Eq(f.SelectedHex == null, true, "bấm vào chỗ trống: không chọn màu");
        f.Dispose();

        // 6) Mất tiêu điểm khi CHƯA chọn thì hủy (không có màu, kết quả Cancel)
        f = RunModal("none", 0, false, true);
        Eq(f.SelectedHex == null, true, "mất tiêu điểm khi chưa chọn: không có màu");
        Eq(f.DialogResult, DialogResult.Cancel, "mất tiêu điểm khi chưa chọn: kết quả Cancel");
        f.Dispose();

        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}
