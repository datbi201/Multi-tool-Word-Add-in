// Rev0.9.5: kiểm thử chuỗi Ribbon XML (hai thẻ My table và Chuẩn hóa). Biên dịch cùng toàn bộ mã nguồn vì cần lớp Connect.
using System;
using System.Collections.Generic;
using System.Linq;
using TroLyKyThuat;

static class TRB
{
    static int fails = 0, total = 0;
    static void Eq(object a, object b, string name) { total++; bool ok = Equals(a, b); if (!ok) fails++; Console.WriteLine((ok ? "PASS " : "FAIL ") + name + (ok ? "" : "  (nhan " + a + ", mong " + b + ")")); }

    // 20. Kiểm tra chuỗi Ribbon XML: hai thẻ, nhóm đúng, không còn id cũ
    static void TestRibbonXml()
    {
        string xml = Connect.RibbonXmlForTest();
        // Phân tích bằng System.Xml
        var doc = new System.Xml.XmlDocument();
        doc.LoadXml(xml);
        Eq(true, true, "Ribbon XML hợp lệ (XmlDocument.LoadXml không lỗi)");

        var ns = new System.Xml.XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("r", "http://schemas.microsoft.com/office/2006/01/customui");

        // tabTable tồn tại
        var tabTable = doc.SelectSingleNode("//r:tab[@id='tabTable']", ns);
        Eq(tabTable != null, true, "Ribbon: thẻ tabTable tồn tại");

        // tabNorm tồn tại
        var tabNorm = doc.SelectSingleNode("//r:tab[@id='tabNorm']", ns);
        Eq(tabNorm != null, true, "Ribbon: thẻ tabNorm tồn tại");

        // tabTable chứa nhóm grpStyle, grpChars, grpRowCol, grpCalc
        Eq(tabTable.SelectSingleNode("r:group[@id='grpStyle']", ns) != null, true, "Ribbon: tabTable có grpStyle");
        Eq(tabTable.SelectSingleNode("r:group[@id='grpChars']", ns) != null, true, "Ribbon: tabTable có grpChars");
        Eq(tabTable.SelectSingleNode("r:group[@id='grpRowCol']", ns) != null, true, "Ribbon: tabTable có grpRowCol");
        var grpCalc = tabTable.SelectSingleNode("r:group[@id='grpCalc']", ns);
        Eq(grpCalc != null, true, "Ribbon: tabTable có grpCalc");

        // grpCalc nhãn mới
        if (grpCalc != null)
        {
            string calcLabel = grpCalc.Attributes["label"] != null ? grpCalc.Attributes["label"].Value : "";
            Eq(calcLabel, "Tính trong bảng", "Ribbon: grpCalc nhãn = Tính trong bảng");
        }

        // Không còn btnCsv và grpExport
        Eq(doc.SelectSingleNode("//*[@id='btnCsv']", ns) == null && doc.SelectSingleNode("//*[@id='btnCsv']") == null, true, "Ribbon: không còn id btnCsv");
        Eq(doc.SelectSingleNode("//*[@id='grpExport']", ns) == null && doc.SelectSingleNode("//*[@id='grpExport']") == null, true, "Ribbon: không còn id grpExport");

        // tabNorm chứa nhóm grpLayoutDoc, grpNum, grpRv, grpLang, grpPane
        Eq(tabNorm.SelectSingleNode("r:group[@id='grpLayoutDoc']", ns) != null, true, "Ribbon: tabNorm có grpLayoutDoc");
        Eq(tabNorm.SelectSingleNode("r:group[@id='grpNum']", ns) != null, true, "Ribbon: tabNorm có grpNum");
        Eq(tabNorm.SelectSingleNode("r:group[@id='grpRv']", ns) != null, true, "Ribbon: tabNorm có grpRv");
        Eq(tabNorm.SelectSingleNode("r:group[@id='grpLang']", ns) != null, true, "Ribbon: tabNorm có grpLang");
        Eq(tabNorm.SelectSingleNode("r:group[@id='grpPane']", ns) != null, true, "Ribbon: tabNorm có grpPane");

        var lang = tabNorm.SelectSingleNode("r:group[@id='grpLang']", ns);
        Eq(lang.Attributes["label"].Value, "Công cụ AI", "Ribbon đổi tên nhóm Công cụ AI");
        var aiButtons = lang.SelectNodes("r:button|r:box/r:button", ns);
        Eq(aiButtons.Count, 5, "Ribbon có đúng năm nút AI");
        foreach (System.Xml.XmlNode button in aiButtons)
            Eq(button.Attributes["size"] == null || button.Attributes["size"].Value == "normal", true, "Ribbon nút AI nhỏ");
        foreach (string id in new[] { "btnAiTools", "btnStyle", "btnGram", "btnSyn", "btnTerm" })
            Eq(lang.SelectSingleNode("r:button[@id='" + id + "']", ns) != null, true,
                "Ribbon: lối mở công cụ ngôn ngữ " + id);

        // Mọi id duy nhất
        var allIds = new System.Collections.Generic.Dictionary<string, bool>();
        bool dup = false;
        foreach (System.Xml.XmlNode node in doc.SelectNodes("//*[@id]"))
        {
            string nid = node.Attributes["id"].Value;
            if (allIds.ContainsKey(nid)) { dup = true; Console.WriteLine("  ID trùng: " + nid); }
            allIds[nid] = true;
        }
        Eq(dup, false, "Ribbon: mọi id duy nhất trong toàn bộ XML");

        // tabTable có getLabel
        if (tabTable != null)
        {
            string gl = tabTable.Attributes["getLabel"] != null ? tabTable.Attributes["getLabel"].Value : "";
            Eq(gl, "GetTabTableLabel", "Ribbon: tabTable có getLabel=GetTabTableLabel");
        }

        // Không còn tabTroLy
        Eq(doc.SelectSingleNode("//r:tab[@id='tabTroLy']", ns) == null, true, "Ribbon: không còn id tabTroLy");

        // Không còn mnuCalc
        Eq(doc.SelectSingleNode("//*[@id='mnuCalc']", ns) == null && doc.SelectSingleNode("//*[@id='mnuCalc']") == null, true, "Ribbon: không còn menu mnuCalc");

        // btnSum..btnFmtProd vẫn tồn tại
        string[] calcBtns = { "btnSum", "btnAvg", "btnMin", "btnMax", "btnCnt", "btnFmtCalc", "btnFmtProd" };
        foreach (string bid in calcBtns)
            Eq((doc.SelectSingleNode("//*[@id='" + bid + "']", ns) ?? doc.SelectSingleNode("//*[@id='" + bid + "']")) != null, true, "Ribbon: nút " + bid + " vẫn tồn tại");

        // Callback nhãn thẻ tabTable: mặc định trả về "My table"
        var conn = new Connect();
        Eq(conn.GetTabTableLabel(null), "My table", "GetTabTableLabel ngoài bảng trả về 'My table'");

        TestToolsTab(doc, ns);
    }

    // Rev0.9.5.14: thẻ "Công cụ" dựng bằng lệnh có sẵn của Word (idMso). Tên idMso sai làm hỏng cả Ribbon, nên mọi tên phải nằm trong danh sách đã kiểm trên Word 16 thật (tools\RealWordMso.cs).
    static readonly string[] VerifiedMso = {
        "FormatPainter", "Font", "FontSize", "Bold", "Italic", "Underline", "Strikethrough", "Subscript", "Superscript",
        "FontColorPicker", "TextHighlightColorPicker", "ChangeCase", "ClearFormatting",
        "Bullets", "Numbering", "MultilevelListGallery", "IndentDecrease", "IndentIncrease", "ParagraphMarks",
        "AlignLeft", "AlignCenter", "AlignRight", "AlignJustify", "LineSpacingMenu", "ParagraphDialog",
        "PageSetupDialog", "PageSizeGallery", "PageOrientationGallery", "ColumnsDialog", "BreaksGallery", "BordersShadingDialog", "PageNumberFormat",
        "SymbolInsertGallery", "ShapesInsertGallery", "SmartArtInsert", "CaptionInsert", "CrossReferenceInsert", "TextBoxInsertGallery", "EquationInsertGallery",
        "PicturesCompress", "PictureCrop", "ZoomOnePage", "ZoomTwoPages", "ZoomPageWidth",
        "TableInsertGallery", "ShadingColorPicker", "TableBordersMenu", "TextDirectionGallery", "TableRowsDistribute", "TableColumnsDistribute", "TableRepeatHeaderRows",
        "SplitCells", "TableSplitTable", "MergeCells", "ConvertTextToTable", "TableAutoFitMenu", "TableFormulaDialog",
        "TableCellAlignTopLeft", "TableCellAlignTopCenter", "TableCellAlignTopRight",
        "TableCellAlignMiddleLeft", "TableCellAlignMiddleCenter", "TableCellAlignMiddleRight",
        "TableCellAlignBottomLeft", "TableCellAlignBottomCenter", "TableCellAlignBottomRight" };

    static void TestToolsTab(System.Xml.XmlDocument doc, System.Xml.XmlNamespaceManager ns)
    {
        var tools = doc.SelectSingleNode("//r:tab[@id='tabTools']", ns);
        Eq(tools != null, true, "Ribbon: thẻ Công cụ (tabTools) tồn tại");
        if (tools == null) return;
        Eq(tools.Attributes["label"].Value, "Công cụ", "Ribbon: nhãn thẻ là Công cụ (khác tên thẻ tự tạo My tools của Word)");
        foreach (string g in new[] { "grpTlFont", "grpTlPara", "grpTlPage", "grpTlIns", "grpTlPic", "grpTlTbl" })
            Eq(tools.SelectSingleNode("r:group[@id='" + g + "']", ns) != null, true, "Ribbon: thẻ Công cụ có nhóm " + g);

        // thẻ Công cụ đứng trước My table để người dùng thấy đầu tiên
        var tabs = doc.SelectNodes("//r:tabs/r:tab", ns);
        Eq(tabs[0].Attributes["id"].Value, "tabTools", "Ribbon: thẻ Công cụ là thẻ đầu tiên của add-in");

        // mọi idMso đều đã kiểm trên Word thật; không có idMso lặp
        var seen = new HashSet<string>();
        bool allVerified = true, noDup = true;
        foreach (System.Xml.XmlNode n in tools.SelectNodes(".//*[@idMso]"))
        {
            string m = n.Attributes["idMso"].Value;
            if (Array.IndexOf(VerifiedMso, m) < 0) { allVerified = false; Console.WriteLine("  idMso chưa kiểm: " + m); }
            if (!seen.Add(m)) { noDup = false; Console.WriteLine("  idMso lặp: " + m); }
        }
        Eq(allVerified, true, "Ribbon: mọi idMso trong thẻ Công cụ nằm trong danh sách đã kiểm trên Word thật");
        Eq(noDup, true, "Ribbon: không lặp idMso trong thẻ Công cụ");
        Eq(seen.Count >= 55, true, "Ribbon: thẻ Công cụ có ít nhất 55 lệnh Word có sẵn (" + seen.Count + ")");

        // các lệnh của add-in trong thẻ Công cụ dùng chung hàm xử lý với nút gốc
        foreach (string b in new[] { "tbParaFormat", "tbPageSetup", "tbAddPort", "tbAddLand" })
        {
            var node = tools.SelectSingleNode(".//r:button[@id='" + b + "']", ns);
            Eq(node != null && node.Attributes["onAction"].Value == "OnBtn" && node.Attributes["getImage"].Value == "GetImg", true, "Ribbon: nút " + b + " dùng OnBtn và GetImg");
        }
        string src = System.IO.File.ReadAllText(@"src\Connect.cs");
        foreach (string pair in new[] { "btnParaFormat\": case \"tbParaFormat", "btnPageSetup\": case \"tbPageSetup", "btnAddPort\": case \"tbAddPort", "btnAddLand\": case \"tbAddLand" })
            Eq(src.Contains("case \"" + pair + "\":"), true, "OnBtn: " + pair.Replace("\": case \"", " và ") + " chạy cùng một lệnh");

        // tên nhóm không thừa dài: Word cần đủ chỗ để hiện nhóm trên màn hình nhỏ
        foreach (System.Xml.XmlNode g in tools.SelectNodes("r:group", ns))
            Eq(g.Attributes["label"].Value.Length <= 14, true, "Ribbon: nhãn nhóm ngắn: " + g.Attributes["label"].Value);
    }

    static int Main()
    {
        TestRibbonXml();
        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}
