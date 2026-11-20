// Rev0.9.5.14: kiểm tra trên Word thật rằng mọi idMso trong Ribbon XML của add-in đều là lệnh có thật.
// Một idMso sai khiến Word bỏ cả Ribbon của add-in, nên phải kiểm trước khi phát hành.
using System;
using System.Collections.Generic;
using System.Threading;
using System.Xml;
using TroLyKyThuat;

static class RealWordMso
{
    [STAThread]
    static int Main()
    {
        var doc = new XmlDocument();
        doc.LoadXml(Connect.BuildRibbonXml());
        var names = new List<string>();
        foreach (XmlNode n in doc.SelectNodes("//*[@idMso]"))
        {
            string m = n.Attributes["idMso"].Value;
            if (!names.Contains(m)) names.Add(m);
        }
        Console.WriteLine("Số lệnh idMso trong Ribbon XML: " + names.Count);

        Type wt = Type.GetTypeFromProgID("Word.Application");
        dynamic word = Activator.CreateInstance(wt);
        word.Visible = false; word.DisplayAlerts = 0;
        Thread.Sleep(1500);
        int bad = 0;
        try
        {
            dynamic d = word.Documents.Add();
            foreach (string m in names)
            {
                try { string label = (string)word.CommandBars.GetLabelMso(m); Console.WriteLine("OK   " + m + " = " + label); }
                catch (Exception) { bad++; Console.WriteLine("SAI  " + m + "  (Word không biết lệnh này)"); }
            }
            d.Close(0);
        }
        finally { try { word.Quit(0); } catch { } }
        Console.WriteLine(bad == 0 ? "TẤT CẢ idMso HỢP LỆ (" + names.Count + ")" : "CÓ " + bad + " idMso SAI");
        return bad;
    }
}
