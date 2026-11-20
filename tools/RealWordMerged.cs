using System;
using System.Collections.Generic;
using System.Reflection;
using TroLyKyThuat;

static class RealWordMerged
{
    static string Row(dynamic tbl, int r)
    {
        var sb = new System.Text.StringBuilder();
        dynamic cells = tbl.Range.Cells;
        int n = (int)cells.Count;
        for (int i = 1; i <= n; i++)
        {
            dynamic c = cells.Item(i);
            if ((int)c.RowIndex != r) continue;
            string t = (string)c.Range.Text; t = t.Replace("\r\a", "").Replace("\r", "");
            sb.Append("[" + (int)c.ColumnIndex + ":" + t + " w=" + Math.Round((double)c.Width) + "]");
        }
        return sb.ToString();
    }

    [STAThread]
    static int Main()
    {
        Type wt = Type.GetTypeFromProgID("Word.Application");
        dynamic word = Activator.CreateInstance(wt);
        word.Visible = false;
        word.DisplayAlerts = 0;
        try
        {
            dynamic doc = word.Documents.Add();
            dynamic tbl = doc.Tables.Add(doc.Range(0, 0), 6, 5);
            string[][] data = {
                new[] { "", "A", "B", "C", "D" },
                new[] { "1", "", "1", "2", "2" },
                new[] { "2", "", "4", "5", "20" },
                new[] { "3", "", "6", "8", "48" },
                new[] { "4", "", "", "", "70" } };
            for (int r = 0; r < data.Length; r++) for (int c = 0; c < 5; c++) if (data[r][c].Length > 0) tbl.Cell(r + 1, c + 1).Range.Text = data[r][c];
            // gộp A6:C6 (hàng 6: ô 2 đến ô 4)
            tbl.Cell(6, 2).Merge(tbl.Cell(6, 4));
            Console.WriteLine("Hang 6 sau gop: " + Row(tbl, 6));
            bool colsThrow = false; try { int n = (int)tbl.Columns.Count; } catch { colsThrow = true; }
            Console.WriteLine("Columns.Count bao loi: " + colsThrow);
            bool c65 = false; try { var x = tbl.Cell(6, 5); } catch { c65 = true; }
            Console.WriteLine("Cell(6,5) bao loi: " + c65);

            var st = new AppSettings(); st.CalcDec = "auto";
            foreach (string op in new[] { "sum", "avg", "min", "max", "cnt" })
            {
                // chọn D2:D5 (hàng 2 đến 5, ô thứ 5)
                tbl.Cell(6, 3).Range.Text = "";
                tbl.Cell(2, 5).Select();
                dynamic sel = word.Selection;
                sel.MoveDown(5, 3, 1);        // wdLine, 3 dòng, mở rộng vùng chọn
                st.CalcOp = op;
                string msg = new WordTables((object)word).Calc(st);
                string d6 = (string)tbl.Cell(6, 3).Range.Text; d6 = d6.Replace("\r\a", "");
                Console.WriteLine(op + " -> D6 = [" + d6 + "] | " + msg);
            }

            // Tích theo định dạng nhiều hàng: B, C đậm
            for (int r = 2; r <= 4; r++) { tbl.Cell(r, 3).Range.Font.Bold = 1; tbl.Cell(r, 4).Range.Font.Bold = 1; }
            tbl.Cell(2, 5).Select(); word.Selection.MoveDown(5, 2, 1);
            var st2 = new AppSettings(); st2.CalcFmtKind = "bold";
            string mp = new WordTables((object)word).CalcProdFmt(st2);
            Console.WriteLine("Tich: D2..D4 = " + ((string)tbl.Cell(2, 5).Range.Text).Replace("\r\a", "") + "|" + ((string)tbl.Cell(3, 5).Range.Text).Replace("\r\a", "") + "|" + ((string)tbl.Cell(4, 5).Range.Text).Replace("\r\a", "") + " | " + mp);
            doc.Close(0);
        }
        finally { try { word.Quit(0); } catch { } }
        return 0;
    }
}
