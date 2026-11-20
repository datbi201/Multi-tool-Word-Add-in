using System;
using System.Threading;
using TroLyKyThuat;

static class RealWordAlign
{
    [STAThread]
    static int Main()
    {
        Type wt = Type.GetTypeFromProgID("Word.Application");
        dynamic word = Activator.CreateInstance(wt);
        word.Visible = false; word.DisplayAlerts = 0;
        Thread.Sleep(1500);
        try
        {
            dynamic doc = word.Documents.Add();
            doc.Content.Text = "Doan thuong mot.\rTieu de muc.\rDoan thuong hai.\rDoan thuong ba.";
            doc.Paragraphs.Item(1).Range.Font.Name = "Times New Roman";
            doc.Paragraphs.Item(1).Range.Font.Size = 12;
            doc.Paragraphs.Item(2).Style = doc.Styles.Item(-2);     // Heading 1
            doc.Paragraphs.Item(3).Range.Font.Name = "Arial"; doc.Paragraphs.Item(3).Range.Font.Size = 14;
            doc.Paragraphs.Item(4).Range.Font.Name = "Times New Roman"; doc.Paragraphs.Item(4).Range.Font.Size = 12;
            dynamic normal = doc.Styles.Item(-1);
            normal.AutomaticallyUpdate = true;
            Console.WriteLine("Normal.AutomaticallyUpdate truoc: " + (bool)normal.AutomaticallyUpdate);

            foreach (string al in new[] { "justify", "left", "center", "right" })
            {
                doc.Range(0, doc.Content.End - 1).Select();        // bôi đen cả 4 đoạn
                var st = new AppSettings(); st.ParaAlign = al;
                string msg = new WordLayout((object)word).StandardizeParagraphs(st);
                var line = al + ": ";
                for (int i = 1; i <= 4; i++)
                {
                    dynamic p = doc.Paragraphs.Item(i);
                    line += "[" + i + " canh=" + (int)p.Format.Alignment + " size=" + (double)p.Range.Font.Size + " dongDau=" + Math.Round((double)p.Format.FirstLineIndent, 2) + "] ";
                }
                Console.WriteLine(line);
                Console.WriteLine("   " + msg);
            }
            Console.WriteLine("Normal.AutomaticallyUpdate sau: " + (bool)normal.AutomaticallyUpdate);
            // một đoạn tại con trỏ
            doc.Paragraphs.Item(4).Range.ParagraphFormat.Alignment = 0;
            doc.Paragraphs.Item(3).Range.ParagraphFormat.Alignment = 0;
            doc.Paragraphs.Item(4).Range.Select(); word.Selection.Collapse(1);
            var st2 = new AppSettings(); st2.ParaAlign = "center";
            string m2 = new WordLayout((object)word).StandardizeParagraphs(st2);
            Console.WriteLine("Con tro: p3 canh=" + (int)doc.Paragraphs.Item(3).Format.Alignment + " p4 canh=" + (int)doc.Paragraphs.Item(4).Format.Alignment + " | " + m2);
            doc.Close(0);
        }
        finally { try { word.Quit(0); } catch { } }
        return 0;
    }
}
