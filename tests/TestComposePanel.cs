using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using TroLyKyThuat;

static class TestComposePanel
{
    static int total, fails;

    static void Check(bool ok, string message)
    {
        total++;
        if (!ok) fails++;
        Console.WriteLine((ok ? "PASS " : "FAIL ") + message);
    }

    static T Field<T>(PaneForm form, string name)
    {
        var field = typeof(PaneForm).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance);
        return field == null ? default(T) : (T)field.GetValue(form);
    }

    static bool HasInsertButton(Control parent)
    {
        foreach (Control child in parent.Controls)
            if ((child is Button && (child.Text == "Chèn" || child.Text == "Thay vùng chọn")) || HasInsertButton(child))
                return true;
        return false;
    }

    [STAThread]
    static int Main(string[] args)
    {
        Application.EnableVisualStyles();
        var doc = new RDoc("Chữ trước | chữ sau\r");
        var app = new RApp(doc);
        var form = new PaneForm(app);
        var page = Field<TabPage>(form, "pCt");
        var input = Field<TextBox>(form, "tbCt");
        var inline = Field<Control>(form, "inlinePreview");
        var display = Field<Control>(form, "displayPreview");
        var insertInline = Field<CheckBox>(form, "insertInline");
        var insertDisplay = Field<CheckBox>(form, "insertDisplay");
        Check(input != null && inline != null && display != null && inline.Parent != display.Parent,
            "Raw input and two separate previews");
        Check(!HasInsertButton(page), "Insertion needs no extra button");
        Check(insertInline != null && insertDisplay != null, "Each preview has its own checkbox");
        if (insertInline != null && insertDisplay != null)
        {
            Check(!insertInline.Enabled && !insertDisplay.Enabled, "Empty input cannot be inserted");
            input.Text = "h = 1/2 * sqrt((e^2 - h)/(h*a - u*e))";
            Check(insertInline.Enabled && insertDisplay.Enabled, "Both actions enabled for valid input");
            Check(doc.EqList.Count == 0, "Typing a formula only updates previews");
            Check(insertInline.Parent == inline.Parent && insertDisplay.Parent == display.Parent,
                "Each checkbox is next to its own preview");
            app.Sel = doc.Mk(0, 9);
            insertInline.Checked = true;
            Check(doc.EqList.Count == 1 && !RRange.LastXml.Contains("<m:oMathPara>") &&
                !RRange.LastXml.Contains("m:val=\"bar\""), "Inline checkbox immediately inserts inline equation");
            Check(!insertInline.Checked && doc.T.ToString().StartsWith("Chữ trước"), "Checkbox resets and selected text is preserved");
            app.Sel = doc.Mk(doc.T.Length - 1, doc.T.Length - 1);
            insertDisplay.Checked = true;
            Check(doc.EqList.Count == 2 && RRange.LastXml.Contains("<m:oMathPara>") &&
                RRange.LastXml.Split(new[] { "m:val=\"bar\"" }, StringSplitOptions.None).Length == 3,
                "Display checkbox immediately inserts both stacked fractions");
            Check(!insertDisplay.Checked, "Display checkbox resets after insertion");
            insertDisplay.Checked = true;
            Check(doc.EqList.Count == 3, "Same checkbox can insert again");

            Field<TabControl>(form, "main").SelectedTab = page;
            var handle = form.Handle;
            foreach (int zoom in new[] { 50, 100, 130 })
            {
                DpiUtil.SetZoom(form, zoom);
                inline.Parent.PerformLayout();
                display.Parent.PerformLayout();
                Check(insertInline.Bottom <= inline.Top && insertDisplay.Bottom <= display.Top &&
                    inline.Height > 0 && display.Height > 0, "Both previews fit at zoom " + zoom);
            }
            if (args.Length > 0)
            {
                DpiUtil.SetZoom(form, 100);
                form.Size = new Size(600, 720);
                form.Show();
                Application.DoEvents();
                using (var bitmap = new Bitmap(form.Width, form.Height))
                {
                    form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                    bitmap.Save(args[0]);
                }
            }
        }
        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}