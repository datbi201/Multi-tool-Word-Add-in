// Kiểm thử bộ nhận diện của mục Chuẩn hóa (đối chiếu với bản demo JavaScript).
using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using TroLyKyThuat;

static class TR
{
    static int fails = 0, total = 0;
    static void Eq(object a, object b, string name)
    {
        total++; bool ok = Equals(a, b); if (!ok) fails++;
        Console.WriteLine((ok ? "PASS " : "FAIL ") + name + (ok ? "" : "  (nhan [" + a + "], mong [" + b + "])"));
    }
    static string Sp(string raw) { Species s = Chem.NormSpecies(raw); return s.Unknown ? "UNK" : s.Out; }
    static List<Hit> D(string t, ReviewOpts o, bool inTable, bool heading, bool numCell)
    {
        return Detector.Detect(t, o, delegate { return new NumCtx { InTable = inTable, Heading = heading, NumericCell = numCell }; });
    }
    static List<Hit> D(string t) { return D(t, new ReviewOpts(), false, false, false); }
    static string Found(List<Hit> h, string t)
    {
        return string.Join(" ; ", h.Select(x => t.Substring(x.S, x.E - x.S) + "|" + x.Issue.Cat + "|" + x.Issue.Sev + "|" + (x.Issue.Fixes.Count > 0 ? x.Issue.Fixes[0].Show : "")).ToArray());
    }
    static int Main()
    {
        // --- hóa học: từng chất
        Eq(Sp("fecl2"), "FeCl2", "fecl2 -> FeCl2"); Eq(Sp("NaCl2"), "NaCl", "NaCl2 -> NaCl"); Eq(Sp("OH2"), "H2O", "OH2 -> H2O");
        Eq(Sp("H2S"), "H2S", "H2S giữ nguyên"); Eq(Sp("Fe2+"), "Fe^2+", "Fe2+ -> Fe^2+"); Eq(Sp("SO4^2-"), "SO4^2-", "SO4^2-");
        Eq(Sp("h2so4"), "H2SO4", "h2so4"); Eq(Sp("naoh"), "NaOH", "naoh"); Eq(Sp("FECL2"), "FeCl2", "FECL2");
        Eq(Sp("OH-"), "OH^-", "OH-"); Eq(Sp("Fe(OH)3"), "Fe(OH)3", "Fe(OH)3"); Eq(Sp("fe3+"), "Fe^3+", "fe3+");
        Eq(Sp("na2so4"), "Na2SO4", "na2so4"); Eq(Sp("nh3"), "NH3", "nh3");
        Eq(Chem.NormSpecies("Co2").Flags.Any(f => f.Type == "val"), true, "Co2 báo hóa trị");
        // --- phương trình
        Rxn r = Chem.AnalyzeRxn("Fe + HCl", "->", "fecl2 + H2");
        Eq(r.Std, "Fe + HCl -> FeCl2 + H2", "chuẩn hóa vế phương trình"); Eq(r.Balanced, false, "Fe + HCl chưa cân bằng"); Eq(r.BalStr, "Fe + 2HCl -> FeCl2 + H2", "gợi ý cân bằng");
        r = Chem.AnalyzeRxn("Fe + H2S", "->", "FeS + H2"); Eq(r.Balanced, true, "Fe + H2S -> FeS + H2 cân bằng");
        r = Chem.AnalyzeRxn("h2so4 + 2naoh", "=", "na2so4 + 2h2o"); Eq(r.Std, "H2SO4 + 2NaOH -> Na2SO4 + 2H2O", "h2so4 + 2naoh"); Eq(r.Balanced, true, "cân bằng");
        r = Chem.AnalyzeRxn("Fe3+ + 3OH-", "-->", "Fe(OH)3"); Eq(r.Std, "Fe^3+ + 3OH^- -> Fe(OH)3", "phản ứng ion"); Eq(r.Balanced, true, "ion cân bằng (cả điện tích)");
        // --- RenderRuns
        List<Run> runs = Chem.RenderRuns("Fe^2+ + SO4^2-");
        Eq(string.Join("|", runs.Select(x => x.Kind + ":" + x.T).ToArray()), "0:Fe|2:2+|0: + SO|1:4|2:2−", "runs: chỉ số trên và dưới");
        Eq(Chem.RunsPlain(runs), "Fe2+ + SO42−", "văn bản thuần từ runs");
        Eq(Chem.RunsToLinear(Chem.RenderRuns("H2O")), "\"H\"_2\"O\"", "tuyến tính H2O");
        // --- nhận diện trong đoạn văn (khớp bản demo)
        string t1 = "The produced gas contains H2S and CO2 at a partial pressure of 0.2 bar. Sour corrosion follows Fe + H2S -> FeS + H2, and the dissolved iron appears as Fe2+. Water (H2O) is present, and the stress ratio is x^2 + y^2 = r^2.";
        string f1 = Found(D(t1), t1);
        Eq(f1.Contains("Fe + H2S -> FeS + H2|chem|fmt"), true, "tìm thấy phương trình");
        Eq(f1.Contains("Fe2+|chem|fmt"), true, "tìm thấy ion Fe2+");
        Eq(f1.Contains("H2S|chem|fmt") && f1.Contains("CO2|chem|fmt"), true, "tìm thấy H2S và CO2"); Eq(f1.Contains("H2O|chem|fmt"), true, "tìm thấy (H2O) trong ngoặc");
        Eq(f1.Contains("x^2 + y^2 = r^2|math|fmt"), true, "tìm thấy biểu thức toán");
        Eq(f1.Contains("0.2"), false, "không đụng tới số 0.2 bar (đã đúng quy ước)");
        string t2 = "The brine contains NaCl2, the condensate is mostly OH2, and acid attack follows Fe + HCl -> fecl2 + H2.";
        string f2 = Found(D(t2), t2);
        Eq(f2.Contains("Fe + HCl -> fecl2 + H2|chem|err"), true, "phương trình sai chữ hoa là mức Sai");
        Eq(f2.Contains("NaCl2|chem|sus|NaCl"), true, "NaCl2 nghi ngờ, đề xuất NaCl"); Eq(f2.Contains("OH2|chem|sus|H₂O"), true, "OH2 nghi ngờ, đề xuất H2O");
        var h2 = D(t2); var rx = h2.First(h => h.Issue.Before.StartsWith("Fe + HCl"));
        Eq(rx.Issue.Fixes.Count, 2, "phương trình chưa cân bằng có 2 cách sửa"); Eq(rx.Issue.Fixes[1].Label, "Sửa và cân bằng", "tên cách sửa thứ hai");
        string t3 = "The minimum required thickness, treq,min, is calculated with the design pressure Pd and the allowable stress S.";
        string f3 = Found(D(t3), t3);
        Eq(f3.Contains("treq,min|sym|err"), true, "treq,min là ký hiệu theo ngữ cảnh (Sai)"); Eq(f3.Contains("Pd|sym|sus"), true, "Pd nghi ngờ"); Eq(f3.Contains("|sym|sus|S"), true, "S nghi ngờ");
        string t4 = "Grade SS316 and ASME B31.3, API 570, Section S1, CUI found. Table 1 shows S1.";
        Eq(D(t4).Count(h => h.Issue.Cat == "chem"), 0, "SS316, S1, CUI không bị coi là hóa học");
        // đường gạch dưới
        string t5 = "Use t_req,min and P_{H2S} here.";
        var h5 = D(t5); Eq(h5.Count(h => h.Issue.Cat == "sym"), 2, "ký hiệu dạng gạch dưới");
        // --- bật tắt nhóm
        var oOff = new ReviewOpts(); oOff.Chem = false; oOff.Math = false; oOff.Sym = false; oOff.Num = false;
        Eq(D(t1, oOff, false, false, false).Count, 0, "tắt hết các nhóm: không tìm gì");
        // --- số trong văn bản
        var oNum = new ReviewOpts(); oNum.Chem = false; oNum.Math = false; oNum.Sym = false; oNum.NumStyle = "us"; oNum.NumDec = "keep"; oNum.NumThou = true;
        string n1 = "The wall loss is 12500 mm at 1,234.5 kPa in 2023 on page 1234, see Table 1,234 and 15 March 2024.";
        string fn = Found(D(n1, oNum, false, false, false), n1);
        Eq(fn.Contains("12500|num|fmt|12,500"), true, "12500 mm -> 12,500 (có đơn vị)"); Eq(fn.Contains("1,234.5"), false, "1,234.5 kPa đã đúng");
        Eq(fn.Contains("2023"), false, "năm 2023 không đụng"); Eq(fn.Contains("2024"), false, "năm 2024 trong ngày tháng không đụng");
        Eq(fn.Split(';').Length, 1, "chỉ một số cần sửa trong câu");
        var oVn = new ReviewOpts(); oVn.Chem = false; oVn.Math = false; oVn.Sym = false; oVn.NumStyle = "vn"; oVn.NumDec = "keep"; oVn.NumThou = true;
        string n2 = "Pressure is 12.7 bar and flow 1250 m3/h.";
        string fv = Found(D(n2, oVn, false, false, false), n2);
        Eq(fv.Contains("12.7|num|fmt|12,7"), true, "12.7 bar -> 12,7 kiểu Việt"); Eq(fv.Contains("1250|num|fmt|1.250"), true, "1250 m3/h -> 1.250 kiểu Việt");
        var oSi = new ReviewOpts(); oSi.Chem = false; oSi.Math = false; oSi.Sym = false; oSi.NumStyle = "si"; oSi.NumDec = "keep"; oSi.NumThou = true;
        string fs = Found(D("Flow 5478.80 m3/h", oSi, false, false, false), "Flow 5478.80 m3/h");
        Eq(fs.StartsWith("5478.80|num|fmt|5\u00A0478,80") || fs.StartsWith("5478.80|num|fmt|5\u202F478,80") || fs.StartsWith("5478.80|num|fmt|5 478,80"), true, "SI: 5478.80 -> 5 478,80 (có nhóm nghìn từ 4 chữ số)");
        Eq(fs.Contains("m3/h|unit|fmt|m³/h"), true, "m3/h đề xuất m³/h (số mũ là chỉ số trên)");
        // số mơ hồ không có đơn vị
        string fa = Found(D("The code 1,234 is listed.", oNum, false, false, false), "The code 1,234 is listed.");
        Eq(fa.Contains("1,234|num|sus"), true, "1,234 không đơn vị là nghi ngờ");
        // tiêu đề và bảng
        Eq(D("Section 12500 overview", oNum, false, true, false).Count, 0, "tiêu đề: số được coi là văn bản");
        Eq(D("12500", oNum, true, false, true).Count, 1, "ô số trong bảng được rà soát");
        var oNoTbl = new ReviewOpts(); oNoTbl.Chem = false; oNoTbl.Math = false; oNoTbl.Sym = false; oNoTbl.NumTable = false;
        Eq(D("12500", oNoTbl, true, false, true).Count, 0, "tắt số trong bảng");
        // đã chuẩn: không báo
        Eq(D("The value 12,500 mm is fine.", oNum, false, false, false).Count, 0, "12,500 mm đã đúng kiểu Mỹ: không báo");
        // === Rev0.3: hóa học gõ kiểu gạch dưới (từ ví dụ của người dùng)
        {
            string c1 = "Water H_2O reacts with O_2 and the ion O^{2-} or SO_4^{2-} forms; also CO_2 and Fe^{3+}.";
            var hc = D(c1); string fc = Found(hc, c1);
            Eq(fc.Contains("H_2O|chem|fmt|H₂O"), true, "H_2O là công thức hóa học, chỉ số 2 là chỉ số dưới");
            Eq(fc.Contains("O_2|chem|fmt|O₂"), true, "O_2 là phân tử oxy");
            Eq(fc.Contains("O^{2-}|chem|fmt|O²⁻"), true, "O^{2-} là ion");
            Eq(fc.Contains("SO_4^{2-}|chem|fmt|SO₄²⁻"), true, "SO_4^{2-} là ion sunfat");
            Eq(fc.Contains("CO_2|chem|fmt|CO₂"), true, "CO_2"); Eq(fc.Contains("Fe^{3+}|chem|fmt|Fe³⁺"), true, "Fe^{3+}");
            Eq(fc.Contains("|sym|"), false, "không còn bị xem là ký hiệu biến");
            Hit h0 = hc.First(x => x.Issue.Before == "H_2O"); Fix f0 = h0.Issue.Fixes[0];
            Eq(f0.IsChem, true, "sửa hóa học đánh dấu chữ thẳng"); Eq(string.Join("|", f0.Runs.Select(x => x.Kind + ":" + x.T).ToArray()), "0:H|1:2|0:O", "runs H_2O: chỉ số 2 dưới, H và O thẳng");
            string c2 = "At point P_1 the pressure is P_1 and the temperature is T_2 while t_req,min is the thickness.";
            string fc2 = Found(D(c2), c2);
            Eq(fc2.Contains("P_1|chem"), false, "P_1 là biến, không phải hóa học"); Eq(fc2.Contains("T_2|chem"), false, "T_2 là biến");
            Eq(fc2.Contains("P_1|sym|fmt|P_1"), true, "P_1 là ký hiệu biến"); Eq(fc2.Contains("t_req,min|sym"), true, "t_req,min là ký hiệu biến");
        }
        // === Rev0.3: số thông minh hơn
        {
            var oN = new ReviewOpts(); oN.Chem = false; oN.Math = false; oN.Sym = false;
            string[] plain = {
                "The Eiffel tower was finished in 1889 and the code was revised in 1839.",
                "From 1990 to 2000 the rate fell; in the 1990s it was low (1995).",
                "On 12.05.2020 and 12/05/2020 and 15 March 2024 we met.",
                "See Section 3.2.1 and Table 4 on page 1234 [12] and [1,2].",
                "Ngày 12 tháng 5 năm 2024, theo điều 1234 của chương 5.",
                "Contact 0912345678 or ID 007.",
                "3.2 Corrosion rate" };
            foreach (string pt in plain) Eq(D(pt, oN, false, false, false).Count, 0, "không nghi ngờ: " + pt);
            string tq3 = "Flow of 12500 is high; approximately 3000 is typical; the value 1.234 appears; 5478,80 bar.";
            string f3x = Found(D(tq3, oN, false, false, false), tq3);
            Eq(f3x.Contains("12500|num|sus"), true, "12500 không đơn vị: nghi ngờ"); Eq(f3x.Contains("3000|num|fmt|3,000"), true, "approximately 3000 là số kỹ thuật");
            Eq(f3x.Contains("1.234|num|sus"), true, "1.234 không manh mối: nghi ngờ");
            string tq4 = "The year 2024 and the value 1500 mm.";
            string f4 = Found(D(tq4, oN, false, false, false), tq4);
            Eq(f4.Contains("2024"), false, "năm 2024 không đụng"); Eq(f4.Contains("1500|num|fmt|1,500"), true, "1500 mm có đơn vị: sửa 1,500");
        }
        // === Rev0.3: in nghiêng biến, thẳng cho hóa học
        {
            List<Run> mr = Detector.MathRuns("x^2 + y^2 = r^2");
            Eq(string.Join("|", mr.Select(x => x.Kind + ":" + x.T).ToArray()), "3:x|2:2|0: + |3:y|2:2|0: = |3:r|2:2", "biến x, y, r nghiêng; số mũ chỉ số trên");
            var oS = new ReviewOpts(); oS.Chem = false; oS.Math = false; oS.Num = false;
            string st1 = "The minimum thickness, treq,min, is calculated.";
            var hs = D(st1, oS, false, false, false);
            Eq(hs.Count, 1, "treq,min nhận diện"); if (hs.Count > 0) Eq(string.Join("|", hs[0].Issue.Fixes[0].Runs.Select(x => x.Kind + ":" + x.T).ToArray()), "3:t|1:req,min", "t nghiêng, req,min chỉ số dưới thẳng");
        }
        // === Rev0.3: soạn công thức
        {
            ComposeResult c;
            c = Compose.Build("2H2 + O2 -> 2H2O", "chem", false);
            Eq(c.Ok, true, "chem phản ứng"); Eq(c.Show, "2H₂ + O₂ → 2H₂O", "xem trước 2H2 + O2 -> 2H2O"); Eq(c.Fix.IsChem, true, "đánh dấu hóa học");
            c = Compose.Build("H_2O", "chem", false); Eq(c.Show, "H₂O", "H_2O thành H₂O");
            c = Compose.Build("SO_4^{2-}", "chem", false); Eq(c.Show, "SO₄²⁻", "SO_4^{2-}");
            c = Compose.Build("Fe + HCl -> FeCl2 + H2", "chem", false); Eq(c.Show, "Fe + HCl → FeCl₂ + H₂", "không cân bằng thì giữ nguyên"); Eq(c.Message.Contains("chưa cân bằng"), true, "báo chưa cân bằng");
            c = Compose.Build("Fe + HCl -> FeCl2 + H2", "chem", true); Eq(c.Show, "Fe + 2HCl → FeCl₂ + H₂", "cân bằng khi được yêu cầu");
            c = Compose.Build("h2so4 + 2naoh = na2so4 + 2h2o", "chem", false); Eq(c.Show, "H₂SO₄ + 2NaOH → Na₂SO₄ + 2H₂O", "sửa chữ hoa và dấu bằng");
            c = Compose.Build("Fe^{3+} + 3OH^-", "chem", false); Eq(c.Show, "Fe³⁺ + 3OH⁻", "ion không có mũi tên");
            c = Compose.Build("x^2 + y^2 = r^2", "math", false); Eq(c.Complex, false, "toán đơn giản không cần Equation");
            Eq(string.Join("|", c.Fix.Runs.Select(x => x.Kind + ":" + x.T).ToArray()), "3:x|2:2|0: + |3:y|2:2|0: = |3:r|2:2", "x, y, r nghiêng");
            c = Compose.Build("(a+b)/c", "math", false); Eq(c.Complex, true, "phân số cần Equation"); Eq(c.Fix.Linear, "(a+b)/c", "UnicodeMath giữ phân số");
            c = Compose.Build("sqrt(x^2+y^2)", "math", false); Eq(c.Complex, true, "căn cần Equation"); Eq(c.Fix.Linear, "√(x^2+y^2)", "UnicodeMath căn");
            c = Compose.Build("t_{req} = P D/(2 S)", "math", false); Eq(c.Fix.Linear, "t_(req) = P D/(2 S)", "ngoặc nhọn thành ngoặc tròn UnicodeMath");
            c = Compose.Build("alpha*beta", "math", false); Eq(c.Show, "α · β", "chữ Hy Lạp và dấu nhân");
            c = Compose.Build("   ", "math", false); Eq(c.Ok, false, "chuỗi rỗng không hợp lệ");
        }
        // === Rev0.4: đơn vị kỹ thuật
        {
            var oU = new ReviewOpts(); oU.Chem = false; oU.Math = false; oU.Sym = false; oU.NumStyle = "vn"; oU.NumThou = true;
            Func<string, string> U = delegate(string t) { return Found(D(t, oU, false, false, false), t); };
            Eq(U("Speed 1.000 rpm."), "", "1.000 rpm: rpm là đơn vị nên 1.000 là số kỹ thuật kiểu Việt, đã đúng, không báo");
            Eq(U("Speed 1500 RPM.").Contains("RPM|unit|fmt|rpm"), true, "RPM đề xuất rpm");
            Eq(U("At 25 kpa gauge").Contains("kpa|unit|err|kPa"), true, "kpa sai chữ hoa: Sai, đề xuất kPa");
            Eq(U("At 25 KPA").Contains("KPA|unit|err|kPa"), true, "KPA đề xuất kPa");
            Eq(U("At 25 Mpa").Contains("Mpa|unit|err|MPa"), true, "Mpa đề xuất MPa (chắc chắn)");
            string um = U("Rated 5 mpa here");
            Eq(um.Contains("mpa|unit|sus|MPa"), true, "mpa chữ thường: nghi ngờ, đề xuất MPa");
            var hm = D("Rated 5 mpa here", oU, false, false, false);
            Eq(hm.Count > 0 && hm[0].Issue.Fixes.Count == 2 && hm[0].Issue.Fixes[1].Plain == "mPa", true, "mpa có cách hiểu thứ hai mPa");
            Eq(U("Rated 25 MPa and 5 kPa and 3.5 MW and 12 kg, 5 s, 300 mm, 80 °C, 15 psi, 2 bar.") == "" || U("Rated 25 MPa and 5 kPa and 3.5 MW and 12 kg, 5 s, 300 mm, 80 °C, 15 psi, 2 bar.").Contains("|unit|") == false, true, "đơn vị đã đúng: không báo đơn vị");
            Eq(U("Area 5 m2 and flow 40 m3/h and 850 kg/m3").Contains("m2|unit|fmt|m²"), true, "m2 thành m²"); Eq(U("flow 40 m3/h and 850 kg/m3").Contains("kg/m3|unit|fmt|kg/m³"), true, "kg/m3 thành kg/m³");
            Eq(U("Rate 0.2 mm/yr").Contains("|unit|"), false, "mm/yr đúng");
            Eq(U("Temp 80℃").Contains("℃|unit|fmt|(cách) °C"), true, "℃ đề xuất °C kèm dấu cách");
            Eq(U("Speed 1500rpm").Contains("1500|num|fmt|1.500"), true, "1500rpm viết liền vẫn nhận ra số và đơn vị");
            Eq(U("The 5G network, 2nd and 3D model").Length, 0, "5G, 2nd, 3D không phải số kèm đơn vị");
            Eq(U("Molarity 5 M solution").Contains("|unit|"), false, "5 M (mol/L) không bị ép thành đơn vị khác");
            UnitHit uh = Units.Scan(" m3/h");
            Eq(uh != null && uh.Canon == "m³/h" && uh.CanonRuns.Count == 3 && uh.CanonRuns[1].Kind == 2, true, "runs m³/h: số mũ là chỉ số trên");
            Eq(Units.Scan(" foo") == null, true, "từ lạ không phải đơn vị");
            Eq(Units.Scan(" years") != null && Units.Scan(" days") != null && Units.Scan(" hours") != null, true, "years, days, hours là đơn vị thời gian");
            Eq(U("Service life 20 years and 5 days").Contains("|unit|"), false, "20 years, 5 days không bị báo đơn vị");
            Eq(U("After 12500 hours of service").Contains("12500|num|fmt|12.500") || U("After 12500 hours of service").Contains("12500|num|fmt|12,500"), true, "12500 hours: số kỹ thuật có đơn vị thời gian");
            Eq(U("Is 20 Years old").Contains("|unit|"), false, "Years viết hoa là từ thường, không bị bắt lỗi đơn vị");
        }
        // === Rev0.5: số thứ tự, chú thích bảng hình, thiếu dấu cách trước đơn vị
        {
            var oN = new ReviewOpts(); oN.NumStyle = "vn"; oN.Chem = false; oN.Math = false; oN.Sym = false;
            Func<string, string> N = delegate(string t) { return Found(D(t, oN, false, false, false), t); };
            Eq(N("1.1. Bối cảnh dự án và cơ sở pháp lý"), "", "1.1. đầu đoạn là số mục, không phải số kỹ thuật");
            Eq(N("2,2. Quan hệ truyền khối"), "", "2,2. đầu đoạn cũng là số mục");
            Eq(N("1.2) Nội dung ghi chú"), "", "1.2) và 3.4: là số mục");
            Eq(N("Bảng 2.1. Kết quả thử nghiệm; Hình 3-2; Hình 3.4–3.6; (Hình 2.1)"), "", "chú thích Bảng, Hình, khoảng 3.4–3.6");
            Eq(N("Phụ lục A.1 và Lô 15-1, Giếng 3.5 bị ngập; Version 2.1"), "", "phụ lục, lô, giếng, phiên bản");
            Eq(N("Theo QCVN 01:2020/BCT và TCVN 5574:2018, Thông tư 12.2020"), "", "số hiệu văn bản");
            Eq(N("Speed 1500 rpm").Contains("1500|num|fmt|1.500"), true, "từ kết thúc bằng ed không bị coi là chú thích");
            Eq(N("Hiệu suất đạt 94,86% và 25ppm, 80°C").Contains("%|unit|fmt|(cách) %"), true, "94,86% thiếu dấu cách");
            Eq(N("Hiệu suất đạt 94,86% và 25ppm, 80°C").Contains("ppm|unit|fmt|(cách) ppm") && N("Hiệu suất đạt 94,86% và 25ppm, 80°C").Contains("°C|unit|fmt|(cách) °C"), true, "25ppm, 80°C thiếu dấu cách");
            Eq(N("Góc 45° và 94,86 % và 25 ppm").Contains("unit"), false, "góc 45° và đã có dấu cách: không báo");
            Eq(N("In the 1990s it was low").Contains("unit"), false, "1990s không phải số kèm đơn vị");
            var hs = D("rate 25ppm now", oN, false, false, false);
            Eq(hs.Count == 1 && hs[0].Issue.Fixes[0].Plain == "\u00A0ppm", true, "sửa là chèn dấu cách không ngắt");
        }
        // === Rev0.4: mã LaTeX lộ ra và chỉ số
        {
            var oY = new ReviewOpts(); oY.Chem = true; oY.Math = true; oY.Sym = true; oY.Num = false;
            Func<string, string> Y = delegate(string t) { return Found(D(t, oY, false, false, false), t); };
            Func<string, Fix> F0 = delegate(string t) { var hs0 = D(t, oY, false, false, false); return hs0.Count > 0 ? hs0[0].Issue.Fixes[0] : null; };
            Func<Fix, string> RK = delegate(Fix f) { return f == null ? "NULL" : string.Join("|", f.Runs.Select(x => x.Kind + ":" + x.T).ToArray()); };
            Eq(RK(F0("\\Delta G_{ads}°")), "0:Δ|3:G|1:ads|0:°", "\\Delta G_{ads}° thành ΔG với ads chỉ số dưới và °");
            Eq(Y("\\Delta G_{ads}°").Contains("|sym|err|ΔG_ads°"), true, "mã LaTeX lộ ra: mức Sai");
            Eq(RK(F0("the stress (\\tau_w = 5 Pa)")), "3:τ|1:w|0: = 5", "\\tau_w = 5 là một công thức: τ nghiêng, w chỉ số dưới");
            Eq(RK(F0("pressure P_CO2 here")), "3:P|1:CO2", "P_CO2: P nghiêng, cả cụm CO2 là chỉ số dưới");
            Eq(RK(F0("pressure P_{CO_2} here")), "3:P|1:CO2", "P_{CO_2}: cả cụm CO2 là chỉ số dưới, không còn dấu ngoặc thừa");
            Eq(Y("pressure P_{CO_2} here").Contains("P_{CO_2}|sym|fmt|P_CO₂"), true, "hiển thị P_CO₂");
            Eq(F0("pressure P_{CO_2} here").Linear, "P_(\"CO\"_2)", "UnicodeMath của P_{CO_2}");
            Eq(RK(F0("t_{req} is")), "3:t|1:req", "t_{req}");
            Eq(RK(F0("value $\\tau_w$ here")), "3:τ|1:w", "có dấu $ bao quanh");
            Eq(D("value $\\tau_w$ here", oY, false, false, false)[0].S, 6, "vùng thay gồm cả dấu $");
            Eq(Y("a \\Delta b").Contains("\\Delta b|sym|err|Δb") , true, "\\Delta b thành Δb");
            Eq(Y("shift \\pm 5 and 3 \\times 4").Contains("\\pm|sym|err|±") && Y("shift \\pm 5 and 3 \\times 4").Contains("\\times|sym|err|×"), true, "\\pm và \\times thành ± và ×");
            Eq(Y("Gas H_2O and O_2").Contains("|sym|"), false, "H_2O và O_2 vẫn là hóa học, không bị coi là ký hiệu");
            Eq(RK(F0("index x_i here")), "3:x|1:i", "x_i");
            Eq(Y("file_name and snake_case"), "", "tên kiểu snake_case không bị đổi");
            Eq(Symbols.ReplaceCommands("\\alpha + \\Delta T \\le 5"), "α + Δ T ≤ 5", "thay mã LaTeX trong chuỗi");
        }
        // === Rev0.4: gợi ý sửa công thức hóa học
        {
            Func<string, string> Sg = delegate(string raw) { Species sp = Chem.NormSpecies(raw); return sp.Unknown ? "UNK" : sp.Out + "|" + string.Join(",", sp.Alts.ToArray()); };
            Eq(Sg("AlSO4"), "Al2(SO4)3|Al2(SO4)3", "AlSO4 -> Al2(SO4)3");
            Eq(Sg("H3SO4"), "H2SO4|H2SO4", "H3SO4 -> H2SO4");
            Eq(Sg("also4"), "Al2(SO4)3|Al2(SO4)3", "also4 -> Al2(SO4)3 (cả chữ hoa)");
            Eq(Sg("FeCl"), "FeCl2|FeCl2,FeCl3", "FeCl -> FeCl2 hoặc FeCl3");
            Eq(Sg("NaSO4"), "Na2SO4|Na2SO4", "NaSO4 -> Na2SO4");
            Eq(Sg("CaOH"), "Ca(OH)2|Ca(OH)2", "CaOH -> Ca(OH)2");
            Eq(Sg("Mg(OH)"), "Mg(OH)2|Mg(OH)2", "Mg(OH) -> Mg(OH)2");
            Eq(Sg("Ca3(PO4)"), "Ca3(PO4)2|Ca3(PO4)2", "Ca3(PO4) -> Ca3(PO4)2");
            Eq(Sg("NH4SO4"), "(NH4)2SO4|(NH4)2SO4", "NH4SO4 -> (NH4)2SO4");
            Eq(Sg("Mg3N"), "Mg3N2|Mg3N2", "Mg3N -> Mg3N2 (hai nguyên tố bất kỳ)");
            Eq(Sg("Fe3O4"), "Fe3O4|", "Fe3O4 có thật (hóa trị hỗn hợp), không sửa");
            Eq(Sg("FeS2"), "FeS2|", "FeS2 (pyrit) không sửa");
            Eq(Sg("C3H8"), "C3H8|", "C3H8 là hữu cơ, không bị bắt hóa trị");
            Eq(Sg("C2H5OH"), "C2H5OH|", "C2H5OH hữu cơ");
            Eq(Sg("Al2(SO4)3"), "Al2(SO4)3|", "Al2(SO4)3 đúng, không đụng");
            Rxn rx2 = Chem.AnalyzeRxn("al + h3so4", "=", "also4 + h2");
            Eq(rx2.Std, "Al + H2SO4 -> Al2(SO4)3 + H2", "ví dụ của người dùng: Std đã sửa công thức");
            Eq(rx2.BalStr, "2Al + 3H2SO4 -> Al2(SO4)3 + 3H2", "ví dụ của người dùng: bản cân bằng");
            Rxn rx3 = Chem.AnalyzeRxn("Fe + HCl", "->", "FeCl + H2");
            Eq(rx3.Std, "Fe + HCl -> FeCl2 + H2", "FeCl chọn FeCl2"); Eq(rx3.BalStr, "Fe + 2HCl -> FeCl2 + H2", "cân bằng FeCl2");
            Eq(rx3.Alts.Count > 0 && rx3.Alts[0].Std == "Fe + HCl -> FeCl3 + H2" && rx3.Alts[0].BalStr == "2Fe + 6HCl -> 2FeCl3 + 3H2", true, "có cách đọc khác FeCl3");
            string tw = "Aluminium dissolves: al + h3so4 = also4 + h2 in acid.";
            var hw = D(tw); var hr = hw.First(x => x.Issue.Cat == "chem");
            Eq(hr.Issue.Fixes.Count, 2, "phương trình: sửa và sửa cân bằng"); Eq(hr.Issue.Fixes[1].Plain, "2Al + 3H2SO4→ Al2(SO4)3 + 3H2".Replace("→ ", "→ ").Replace("H2SO4→", "H2SO4 →") , "bản cân bằng dạng chữ");
            Eq(hr.Issue.Reason.Contains("H3SO4 không hợp hóa trị"), true, "lý do nêu H3SO4 không hợp hóa trị");
        }
        // === Rev0.4: Định dạng số (NumberEdit)
        {
            var o = new NumEditOpts(); o.FallbackStyle = "us"; o.Thousands = true;
            Func<string, string, string> Cy = delegate(string t, string style) { var oo = new NumEditOpts(); oo.Style = style; oo.FallbackStyle = "us"; return NumberEdit.Plan(t, oo, NumberEdit.IsPure(t)).Text; };
            Eq(NumberEdit.NextStyle("vn"), "us", "vòng kiểu: Việt sang Mỹ"); Eq(NumberEdit.NextStyle("us"), "si", "Mỹ sang SI"); Eq(NumberEdit.NextStyle("si"), "vn", "SI quay về Việt");
            Eq(Cy("1,234.56", "vn"), "1.234,56", "Mỹ sang Việt"); Eq(Cy("1.234,56", "us"), "1,234.56", "Việt sang Mỹ");
            Eq(Cy("1,234.56", "si"), "1\u00A0234,56", "Mỹ sang SI");
            Eq(Cy("1\u00A0234,56", "us"), "1,234.56", "SI sang Mỹ");
            Eq(Cy("12500", "vn"), "12.500", "số nguyên 12500 sang Việt");
            Eq(Cy("0.5", "vn"), "0,5", "0.5 sang 0,5");
            Eq(Cy("12.5; 1,234.5; 3,5", "us"), "12.5; 1,234.5; 3.5", "nhiều số cùng lúc, số rõ ràng quyết định dấu thập phân (3,5 -> 3.5)");
            Eq(Cy("-5.25", "vn"), "-5,25", "số âm"); Eq(Cy("\u22125,25", "us"), "\u22125.25", "dấu trừ − được giữ");
            string tx = "See Table 3 on page 12: the load is 1234.5 kN in 2024.";
            Eq(Cy(tx, "vn"), "See Table 3 on page 12: the load is 1.234,5 kN in 2024.", "trong câu văn chỉ đổi số kỹ thuật, không đụng số bảng, trang, năm");
            Func<string, int, string> De = delegate(string t, int dd) { var oo = new NumEditOpts(); oo.DecDelta = dd; oo.FallbackStyle = "us"; return NumberEdit.Plan(t, oo, NumberEdit.IsPure(t)).Text; };
            Eq(De("12.5", 1), "12.50", "tăng thập phân"); Eq(De("12", 1), "12.0", "số nguyên tăng thành 12.0"); Eq(De("12.345", -1), "12.35", "giảm thập phân làm tròn lên");
            Eq(De("12", -1), "12", "số nguyên không giảm được"); Eq(De("12.349", -1), "12.35", "12.349 giảm một chữ số");
            Eq(De("0.123456", 1), "0.123456", "đã quá 5 chữ số: không tăng, không làm tròn xuống"); Eq(De("0.12345", 1), "0.12345", "đang 5 chữ số: tối đa");
            Eq(De("1.234,5", 1), "1.234,50", "giữ kiểu Việt và nhóm nghìn khi tăng thập phân"); Eq(De("1\u00A0234,5", 1), "1\u00A0234,50", "giữ kiểu SI khi tăng");
            Eq(De("1,234", 1), "1,234.0", "1,234 (hàng nghìn kiểu Mỹ) thêm một chữ số thập phân"); Eq(De("1234.5", 1), "1234.50", "không tự nhóm nghìn khi chỉ đổi thập phân");
            Eq(De("12.5\t8.25\r5", 1), "12.50\t8.250\r5.0", "nhiều số cùng lúc, mỗi số tăng một chữ số");
            Eq(De("the 3 valves and 12.5 mm", 1), "the 3 valves and 12.50 mm", "trong câu, số nguyên thường ngày (3 valves) không bị thêm thập phân");
            Eq(NumberEdit.IsPure("12,5 kPa"), true, "12,5 kPa là dãy số thuần"); Eq(NumberEdit.IsPure("see 12 now"), false, "có chữ thường là không thuần");
            var rr = NumberEdit.Plan("a 1.5 b 2.5", o, false, 3, 6); Eq(rr.Total, 1, "chỉ xét số chứa con trỏ");
        }
        // === Rev0.4: Soạn phương trình (tự nhận loại, đề xuất sửa)
        {
            Eq(Compose.DetectKind("al+h3so4=also4+h2"), "chem", "nhận là hóa học"); Eq(Compose.DetectKind("x^2 + y^2 = r^2"), "math", "x^2... là toán");
            Eq(Compose.DetectKind("P = F"), "math", "P = F là toán"); Eq(Compose.DetectKind("SO4^2-"), "chem", "ion"); Eq(Compose.DetectKind("H_2O"), "chem", "H_2O");
            Eq(Compose.DetectKind("t_{req} = (P D)/(2 S E)"), "math", "công thức Barlow"); Eq(Compose.DetectKind("\\Delta G = -R T ln K"), "math", "có mã LaTeX là toán");
            Eq(Compose.DetectKind("Fe + 2HCl -> FeCl2 + H2"), "chem", "phản ứng"); Eq(Compose.DetectKind("CO2 + H2O = H2CO3"), "chem", "dấu = giữa các chất");
            Eq(Compose.DetectKind("I = V/R"), "math", "I = V/R là toán"); Eq(Compose.DetectKind("Fe"), "chem", "Fe là nguyên tố"); Eq(Compose.DetectKind("n + 1"), "math", "n + 1 là toán");
            ComposeResult c = Compose.Build("al+h3so4=also4+h2");
            Eq(c.Ok && c.Kind == "chem", true, "ví dụ người dùng dựng được");
            Eq(c.Options.Count >= 3, true, "có các phương án: sửa, cân bằng, giữ nguyên gõ");
            Eq(c.Options[0].Show, "Al + H₂SO₄ → Al₂(SO₄)₃ + H₂", "phương án 1: đã sửa công thức");
            Eq(c.Show, "2Al + 3H₂SO₄ → Al₂(SO₄)₃ + 3H₂", "mặc định chọn bản cân bằng khi chưa gõ hệ số");
            Eq(c.Message.Contains("H3SO4 không hợp hóa trị") && c.Message.Contains("AlSO4 không hợp hóa trị"), true, "cảnh báo từng chất sai");
            Eq(c.Options.Any(x => x.Label == "Giữ công thức như đã gõ" && x.Show == "Al + H₃SO₄ → AlSO₄ + H₂"), true, "có phương án giữ nguyên như đã gõ");
            c = Compose.Build("2Al + 3H3SO4 = AlSO4 + 3H2"); Eq(c.Selected, 0, "đã gõ hệ số thì không tự đổi hệ số");
            c = Compose.Build("Fe + HCl -> FeCl + H2"); Eq(c.Options.Any(x => x.Show.Contains("FeCl₃")), true, "FeCl có phương án đọc FeCl3"); Eq(c.Message.Contains("FeCl2 hoặc FeCl3"), true, "cảnh báo nêu cả hai khả năng");
            c = Compose.Build("2H2 + O2 -> 2H2O"); Eq(c.Options.Count, 1, "phương trình đúng: một phương án"); Eq(c.Show, "2H₂ + O₂ → 2H₂O", "xem trước đúng"); Eq(c.Warnings.Count, 0, "không cảnh báo");
            c = Compose.Build("\\Delta G_{ads} = -R T ln K");
            Eq(c.Kind, "math", "toán"); Eq(c.Show, "Δ G_ads = -R T ln K", "Δ, G, chỉ số ads");
            Eq(string.Join("|", c.Fix.Runs.Take(4).Select(x => x.Kind + ":" + x.T).ToArray()), "0:Δ|0: |3:G|1:ads", "Δ thẳng, G nghiêng, ads chỉ số dưới");
            c = Compose.Build("P_{CO_2} = x_{CO_2} P"); Eq(c.Fix.Runs.Any(x => x.Kind == 1 && x.T == "CO2"), true, "P_{CO_2}: cả CO2 là chỉ số dưới");
            Eq(c.Fix.Linear.StartsWith("P_(CO2)") || c.Fix.Linear.StartsWith("P_(CO_2)"), true, "UnicodeMath của chỉ số");
            c = Compose.Build("\\frac{a+b}{c}"); Eq(c.Complex, true, "\\frac là phân số"); Eq(c.Fix.Linear, "(a+b)/(c)", "UnicodeMath của \\frac");
            c = Compose.Build("\\sqrt{x}"); Eq(c.Fix.Linear, "√(x)", "\\sqrt{x}");
            c = Compose.Build("x = (a + b"); Eq(c.Warnings.Any(w => w.Contains("ngoặc tròn")), true, "cảnh báo ngoặc chưa khớp");
            c = Compose.Build("sum_(i=1)^n x_i"); Eq(c.Fix.Linear.StartsWith("∑_(i=1)^n"), true, "sum thành ∑"); Eq(c.Complex, true, "∑ cần Equation");
            c = Compose.Build("   "); Eq(c.Ok, false, "rỗng không hợp lệ");
        }
        // === Rev0.5: thư viện LaTeX lộ ra, công thức nhiều dòng, OMML
        {
            Func<string, string[]> B = delegate(string t) { return D(t).Select(h => h.Issue.Before).ToArray(); };
            Eq(string.Join("|", B("Số Reynolds quay (Re_{cyl}):")), "Re_{cyl}", "Re_{cyl}: cơ sở hai chữ cái");
            Eq(string.Join("|", B("Re_{cyl} = \\frac{ρ · ω · r_cyl · d_cyl}{μ} = \\frac{ω d_cyl²}{2ν}")), "Re_{cyl} = \\frac{ρ · ω · r_cyl · d_cyl}{μ} = \\frac{ω d_cyl²}{2ν}", "cả công thức hai phân số là một mục");
            Eq(string.Join("|", B("liên hệ qua công thức ω = \\frac{2π N}{60}.")), "ω = \\frac{2π N}{60}", "công thức sau chữ thường: dừng ở từ tiếng Việt");
            Eq(string.Join("|", B("ω = \\frac{2π · 1.000}{60} = 104,72 rad/s")), "ω = \\frac{2π · 1.000}{60} = 104,72", "dừng trước đơn vị rad/s");
            Eq(string.Join("|", B("Re_{cyl} = \\frac{104,72 · (0,012)²}{2 · (0,976 × 10⁻⁶)} ≈ 7.725")), "Re_{cyl} = \\frac{104,72 · (0,012)²}{2 · (0,976 × 10⁻⁶)} ≈ 7.725", "phân số có số mũ Unicode và ngoặc");
            Eq(string.Join("|", B("τw, RCE = 0,0791 · Re_{cyl}^{-0,3} · ρ · (ω · r_cyl)²")), "τw, RCE = 0,0791 · Re_{cyl}^{-0,3} · ρ · (ω · r_cyl)²", "τw, RCE và số mũ ^{-0,3}");
            Eq(string.Join("|", B("f = 0,0791 · (7.725)^{-0,3} ≈ 0,00536.")), "f = 0,0791 · (7.725)^{-0,3} ≈ 0,00536", "(7.725)^{-0,3} trong công thức, bỏ dấu chấm cuối câu");
            Eq(string.Join("|", B("giá trị (7.725)^{-0,3} trong")), "(7.725)^{-0,3}", "lũy thừa của nhóm ngoặc đứng riêng");
            Eq(string.Join("|", B("t_req = P D / (2 S E) và R_cyl")), "t_req = P D / (2 S E)|R_cyl", "công thức gõ phẳng và ký hiệu riêng");
            Eq(string.Join("|", B("Công thức rất tốt và 1990s rpm = 5")), "", "văn xuôi không bị nhận nhầm là công thức");
            Hit gh = D("τw, RCE = 5").First(); Eq(gh.Issue.Cat == "sym" || gh.Issue.Cat == "math", true, "τw, RCE = ... được nhận");
            // nội dung sửa: Re thẳng, cyl thẳng chỉ số dưới
            Fix rf = D("Số (Re_{cyl}) tại").First().Issue.Fixes[0];
            Eq(string.Join("|", rf.Runs.Select(x => x.Kind + ":" + x.T).ToArray()), "0:Re|1:cyl", "Re đứng thẳng, cyl chỉ số dưới");
            // cây công thức và OMML
            MNode mn = MathParser.Parse("Re_{cyl} = \\frac{ρ · ω · r_cyl · d_cyl}{μ}");
            string om = Omml.Build(mn, true);
            Eq(om.Contains("<m:f>") && om.Contains("m:val=\"bar\""), true, "phân số xếp trên dưới (bar)");
            Eq(om.Contains("<m:sSub>"), true, "chỉ số dưới dùng m:sSub");
            Eq(!om.Contains("&quot;") && !Regex.IsMatch(om, "<m:t[^>]*>[^<]*[\"“”][^<]*</m:t>"), true, "không có dấu nháy trong chữ công thức");
            Eq(Regex.Matches(om, "<m:t xml:space=\"preserve\">cyl</m:t>").Count >= 3, true, "cyl là chữ riêng trong chỉ số");
            Eq(Regex.IsMatch(om, "<m:rPr><m:sty m:val=\"p\"/></m:rPr><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr><m:t xml:space=\"preserve\">cyl</m:t>"), true, "cyl đứng thẳng (m:sty p)");
            Eq(Regex.IsMatch(om, "<m:r><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr><m:t xml:space=\"preserve\">r</m:t></m:r>"), true, "r là biến, nghiêng (không có m:sty)");
            string lin = Omml.Build(mn, false);
            Eq(lin.Contains("m:val=\"lin\"") && !lin.Contains("m:val=\"bar\""), true, "kiểu tuyến tính: m:f lin");
            string pk = Omml.Package(om, true);
            System.Xml.XmlDocument xd = new System.Xml.XmlDocument(); xd.LoadXml(pk);
            Eq(xd.GetElementsByTagName("m:oMathPara").Count, 1, "gói Flat OPC đọc được và có oMathPara");
            // OMML từ các đoạn chữ: hóa học thẳng, không có dấu nháy
            Fix cf = Fix.FromChem("Sửa", "2Al + 6HCl -> 2AlCl3 + 3H2");
            string cx = Omml.Build(cf.GetNode(), true);
            Eq(!cx.Contains("&quot;") && !cx.Contains("\"Al\""), true, "hóa học: không có dấu nháy quanh nguyên tố");
            Eq(!Regex.IsMatch(cx, "<m:r><w:rPr>[^<]*<[^>]*>[^<]*</w:rPr><m:t[^>]*>[A-Za-z]"), true, "hóa học: mọi chữ cái đều có m:sty p");
            Eq(Regex.Matches(cx, "<m:sSub>").Count, 2, "AlCl3 và H2: hai chỉ số dưới");
            Eq(cx.Contains("→"), true, "mũi tên phản ứng");
            xd = new System.Xml.XmlDocument(); xd.LoadXml(Omml.Package(cx, false));
            Eq(xd.GetElementsByTagName("m:oMath").Count, 1, "gói hóa học hợp lệ XML");
            // chữ Hy Lạp, chỉ số trên + dưới, ký hiệu
            string o2 = Omml.Build(MathParser.Parse("x_i^2 + \\Delta P"), true);
            Eq(o2.Contains("<m:sSubSup>"), true, "x_i^2 dùng m:sSubSup");
            string o3 = Omml.Build(MathParser.Parse("t_req = P D / (2 S E)"), true);
            Eq(o3.Contains("<m:f>") && !o3.Contains("<m:t xml:space=\"preserve\">/</m:t>"), true, "dấu / thành phân số khi xếp chồng");
            string o4 = Omml.Build(MathParser.Parse("t_req = P D / (2 S E)"), false);
            Eq(o4.Contains("<m:t>/</m:t>") && !o4.Contains("<m:f>"), true, "dấu / giữ nguyên khi kiểu tuyến tính");
            Eq(Omml.Build(MathParser.Parse("\\sqrt{x^2+y^2}"), true).Contains("<m:rad>"), true, "căn");
            Eq(Omml.Build(MathParser.Parse("(7.725)^{-0,3}"), true).Contains("<m:sSup>"), true, "(7.725)^{-0,3} là m:sSup");
            Eq(Omml.PlainOf(MathParser.Parse("(7.725)^{-0,3}")), "(7.725)−0,3", "dạng chữ thuần của lũy thừa");
            // ký hiệu cũ không còn linear có dấu nháy trong Equation: Fix.GetNode
            Fix sf = D("E_corr không").First().Issue.Fixes[0];
            string se = Omml.Build(sf.GetNode(), true);
            Eq(se.Contains("<m:sSub>") && se.Contains(">corr<") && !se.Contains("&quot;"), true, "E_corr: m:sSub, corr thẳng, không nháy");
            // Compose: toán có phân số có Node
            ComposeResult cm = Compose.Build("t_{req} = (P D)/(2 S E)");
            Eq(cm.Fix.Node != null && Omml.IsComplex(cm.Fix.Node, true), true, "Compose: có cây công thức, xếp chồng thì phức tạp");
            Eq(Omml.IsComplex(cm.Fix.Node, false), false, "Compose: kiểu tuyến tính không bắt buộc Equation");
        }
        // === Rev0.6: đọc ngược OMML (Equation → cây công thức), bước đầu của Kiểu chỉ số khi đổi Equation về font thường
        {
            foreach (string ex in new[] { "R_cyl", "P_CO2 = 5", "t_req = (P D)/(2 S E)", "x^2 + y^2 = r^2", "\\Delta G_{ads} = -R T ln K", "H_2SO_4", "Re_{cyl}^{-0,3}" })
            {
                MNode om1 = MathParser.Parse(ex);
                foreach (bool stk in new[] { true, false })
                {
                    string pk = Omml.Package(Omml.Build(om1, stk), false);
                    MNode om2 = Omml.FromOmml(pk);
                    Eq(om2 != null, true, "đọc lại OMML: " + ex);
                    if (om2 == null) continue;
                    // Equation của Word không giữ khoảng trắng, nên so sánh sau khi bỏ khoảng trắng và gộp các đoạn cùng kiểu
                    Func<MNode, string> sig = delegate(MNode nd)
                    {
                        var parts = new List<string>(); int kd = -1; var cur = new System.Text.StringBuilder();
                        foreach (Run q in Omml.ToRuns(nd))
                        {
                            string tx = q.T.Replace(" ", ""); if (tx.Length == 0) continue;
                            if (q.Kind != kd && cur.Length > 0) { parts.Add(kd + ":" + cur); cur.Length = 0; }
                            kd = q.Kind; cur.Append(tx);
                        }
                        if (cur.Length > 0) parts.Add(kd + ":" + cur);
                        return string.Join(",", parts.ToArray());
                    };
                    Eq(sig(om2), sig(om1), "OMML khứ hồi giữ chữ, nghiêng, chỉ số: " + ex + (stk ? " (đứng)" : " (dòng)"));
                }
            }
            MNode nr = Omml.FromOmml(Omml.Package(Omml.Build(MathParser.Parse("R_cyl"), true), false));
            List<Run> rr = Omml.ToRuns(nr);
            Eq(rr.Count == 2 && rr[0].T == "R" && rr[0].Kind == 3 && rr[1].T == "cyl" && rr[1].Kind == 1, true, "R_cyl: R nghiêng, cyl chỉ số dưới");
            Eq(Omml.FromOmml("không phải xml"), null, "XML hỏng trả về null"); Eq(Omml.FromOmml("<a/>"), null, "XML không có m:oMath trả về null");
            // gói kiểu Range.WordOpenXML: nhiều phần, công thức nằm trong w:body
            string wx = "<?xml version=\"1.0\"?><pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/word/styles.xml\"><pkg:xmlData><w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/></pkg:xmlData></pkg:part>"
                + "<pkg:part pkg:name=\"/word/document.xml\"><pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><w:body><w:p><m:oMathPara><m:oMath><m:sSub><m:e><m:r><m:t>T</m:t></m:r></m:e><m:sub><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>w</m:t></m:r></m:sub></m:sSub><m:r><m:t>=1</m:t></m:r></m:oMath></m:oMathPara></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
            List<Run> wr = Omml.ToRuns(Omml.FromOmml(wx));
            Eq(string.Join(",", wr.Select(q => q.Kind + ":" + q.T).ToArray()), "3:T,1:w,0:=1", "gói WordOpenXML: T nghiêng, w chỉ số dưới thẳng, =1 thẳng");
            // số đánh số mục
            var nopt = new ReviewOpts(); nopt.NumStyle = "vn";
            Eq(D("Xem mục 1.1. và 2.3 trong báo cáo, Ω\\cdotcm² cho lớp màng.").Any(h => h.Issue.Cat == "sym" && h.Issue.Fixes.Count > 0 && h.Issue.Fixes[0].Plain == "·"), true, "Ω\\cdotcm²: \\cdot dính đơn vị vẫn nhận ra, đổi thành ·");
        }
        var auto = new ReviewOpts { NumStyle = "auto" };
        const string mixed = "The pressure is 12,5 MPa; áp suất đo là 12.5 MPa. The test limit is 12.5 MPa. Áp suất giới hạn là 12,5 MPa.";
        var numberHits = Detector.Detect(mixed, auto, null).Where(h => h.Issue.Cat == "num").ToList();
        Eq(string.Join("|", numberHits.Select(h => h.Issue.Before + "→" + h.Issue.Fixes[0].Plain)),
            "12,5→12.5|12.5→12,5", "chọn dấu theo ngôn ngữ gần từng số trong đoạn trộn Việt–Anh");
        Eq(numberHits.All(h => h.Issue.Sev == "err"), true, "sai dấu thập phân theo ngôn ngữ được báo là lỗi");
        Eq(Detector.Detect("The pressure is 1,234 MPa.", auto, null).Any(h => h.Issue.Cat == "num"), false,
            "tự động không đổi giá trị của số Anh có nhóm nghìn");
        Eq(Detector.Detect("Áp suất là 1.234 MPa.", auto, null).Any(h => h.Issue.Cat == "num"), false,
            "tự động không đổi giá trị của số Việt có nhóm nghìn");
        Eq(Detector.Detect("12,5 MPa", auto, null).Any(h => h.Issue.Cat == "num"), false,
            "thiếu ngữ cảnh thì không ép đổi dấu");
        var tagged = new NumCtx { Language = (start, end) => "en" };
        Eq(Detector.Detect("12,5 MPa", auto, () => tagged).Single(h => h.Issue.Cat == "num").Issue.Fixes[0].Plain,
            "12.5", "dùng ngôn ngữ Word gán khi chỉ có số và đơn vị");
        Eq(Detector.Detect("Áp suất 12,5 MPa", auto, () => tagged).Any(h => h.Issue.Cat == "num"), false,
            "ưu tiên chữ quanh số khi Word gán tiếng Anh cho cả tài liệu");
        Eq(Detector.Detect("The load is 1\u202F234.50 N.", auto, null).Any(h => h.Issue.Cat == "num"), false,
            "giữ số đúng theo nhóm khoảng trắng SI");
        Eq(Detector.Detect("The load is 1\u202F234,50 N.", auto, null).Single(h => h.Issue.Cat == "num").Issue.Fixes[0].Plain,
            "1\u202F234.50", "chỉ đổi dấu thập phân của số đã nhóm kiểu SI");
        // chuẩn hóa chuỗi toán
        Eq(Detector.NormMath("x^2*y<=alpha"), "x^2 · y≤α", "NormMath");
        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}
