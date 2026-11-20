// Trợ lý Kỹ thuật - Rev0.9.5: cửa sổ Cài đặt riêng (mở từ nút Cài đặt trên Ribbon): quy ước số, áp dụng master style, khóa API, trợ giúp.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace TroLyKyThuat
{
    public sealed class SettingsForm : Form
    {
        AppSettings st = AppSettings.Load();
        readonly ToolTip tip = new ToolTip();
        FlowLayoutPanel f;
        int W = 480;
        ComboBox cbParaAlign, cbNumStyle, cbNumDec, cbCalcDec, cbCalcFmt, cbSciDec, cbUnitNames, cbPagePaper, cbAiModel;
        CheckBox chThou, chAddDec, chTop, chClear, chDlg;
        TextBox tbKey, tbAiTemp, tbAiMaxTok;
        TextBox tbPtTop, tbPtBot, tbPtLeft, tbPtRight, tbLnTop, tbLnBot, tbLnLeft, tbLnRight, tbHdrDist, tbFtrDist;
        TextBox tbSzTnr, tbSzArial, tbLineSp, tbSpBef, tbSpAft, tbFirstLine, tbHanging;
        Label lblKey, lblAiProv;
        Button btnCheckAi, btnTestSaveAi, btnClearAi;
        bool aiBusy;

        /// <summary>Rev0.9.5.2: co giãn theo DPI của màn hình đang hiện (xem DpiUtil).</summary>
        protected override void WndProc(ref Message m)
        {
            if (DpiUtil.HandleDpiMessage(this, ref m)) return;
            base.WndProc(ref m);
        }

        public SettingsForm()
        {
            Text = "Trợ lý Kỹ thuật - Cài đặt";
            DpiUtil.ConfigureForm(this, 11f);
            Shown += delegate { DpiUtil.SyncDpi(this); Log.W("Cài đặt hiện: " + DpiUtil.Describe(this)); };
            Width = 600; Height = 820; MinimumSize = new Size(480, 520);
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = true;
            tip.AutoPopDelay = 20000; tip.InitialDelay = 400; tip.ShowAlways = true;

            var bottom = new TableLayoutPanel(); bottom.Dock = DockStyle.Bottom; bottom.Height = 56; bottom.ColumnCount = 2; bottom.RowCount = 1; bottom.Padding = new Padding(8);
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f)); bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            var ok = new Button(); ok.Text = "Lưu và đóng"; ok.Dock = DockStyle.Fill; ok.AccessibleName = "Lưu và đóng"; ok.Click += delegate { Save(); DialogResult = DialogResult.OK; Close(); };
            var cancel = new Button(); cancel.Text = "Hủy"; cancel.Dock = DockStyle.Fill; cancel.AccessibleName = "Hủy, không lưu thay đổi"; cancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
            bottom.Controls.Add(ok, 0, 0); bottom.Controls.Add(cancel, 1, 0);
            Controls.Add(bottom);
            AcceptButton = ok; CancelButton = cancel;

            var scroll = new Panel(); scroll.Dock = DockStyle.Fill; scroll.AutoScroll = true;
            f = new FlowLayoutPanel(); f.FlowDirection = FlowDirection.TopDown; f.WrapContents = false; f.AutoSize = true; f.AutoSizeMode = AutoSizeMode.GrowAndShrink; f.Padding = new Padding(8);
            scroll.Controls.Add(f); Controls.Add(scroll); scroll.BringToFront();
            UiZoom.Attach(this, scroll, "settings", delegate { Reflow(scroll); });
            scroll.Resize += delegate { Reflow(scroll); };

            Head("Số");
            cbNumStyle = Combo("Kiểu dấu của nút Dấu trên Ribbon", new Item[] { new Item("Việt: 1.234,56", "vn"), new Item("Mỹ: 1,234.56", "us"), new Item("SI: 1 234,56", "si") }, "Kiểu số mà nút Dấu trên Ribbon đang dùng và đổi theo vòng. Kiểu dấu khi rà soát chọn riêng ở bảng điều khiển, thẻ Chuẩn hóa kỹ thuật.");
            chThou = Chk("Dấu hàng nghìn", "Nhóm hàng nghìn từ 4 chữ số trở lên, ở cả ba kiểu.");
            cbNumDec = Combo("Số lẻ khi rà soát", new Item[] { new Item("Giữ nguyên", "keep"), new Item("0", "0"), new Item("1", "1"), new Item("2", "2"), new Item("3", "3"), new Item("4", "4") }, "Số chữ số thập phân khi mục Chuẩn hóa kỹ thuật đề xuất sửa số.");
            chAddDec = Chk("Thêm số lẻ cho số nguyên", "Khi rà soát, số nguyên cũng được thêm chữ số thập phân.");
            cbSciDec = Combo("Số lẻ của kiểu kỹ thuật (2,3 × 10⁶)", new Item[] { new Item("0", "0"), new Item("1", "1"), new Item("2", "2"), new Item("3", "3"), new Item("4", "4"), new Item("5", "5") }, "Số chữ số thập phân của phần đầu khi nút Đồng bộ số lẻ ở chế độ viết kiểu kỹ thuật. Ví dụ 2 chữ số: 3,67 × 10³.");
            cbUnitNames = Combo("Tên đơn vị khi Đọc số", new Item[] { new Item("Theo luật VN: kilôpascan, niutơn", "law"), new Item("Thông dụng: kilopascal, newton", "common") }, "Nút Đọc số đọc cả đơn vị đứng sau số. Theo luật: tên trong Nghị định 134/2007/NĐ-CP (kilôgam, niutơn, pascan, oát, vôn, độ Celsius...). Thông dụng: tên quen dùng như kilopascal, newton, pascal, watt, volt.");
            Head("Tính");
            cbCalcDec = Combo("Số lẻ của kết quả", new Item[] { new Item("Tự động", "auto"), new Item("0", "0"), new Item("1", "1"), new Item("2", "2"), new Item("3", "3"), new Item("4", "4") }, "Số chữ số thập phân của kết quả khi bấm Tính.");
            cbCalcFmt = Combo("Định dạng để tính tổng và tích", new Item[] { new Item("Chữ đậm", "bold"), new Item("Tô sáng (highlight)", "highlight"), new Item("Màu chữ", "color"), new Item("Màu nền ô", "fill") }, "Thử nghiệm (Rev0.9.1A): nút Tổng theo định dạng và Tích theo định dạng chỉ tính các ô có kiểu này. Chữ đậm do kiểu bảng (ví dụ hàng tiêu đề) cũng được tính là đậm.");
            Head("Bảng");
            chClear = Chk("Xóa định dạng chỉnh tay khi áp dụng master style", "Xóa màu chữ, đậm, tô nền ô, thụt lề trước khi áp dụng master style cho bảng. Giữ nguyên chữ nghiêng, chỉ số trên và chỉ số dưới.");
            Head("Dàn trang");
            cbPagePaper = Combo("Khổ giấy", new Item[] { new Item("A4 (21 × 29,7 cm)", "A4"), new Item("Letter (8.5 × 11 in)", "Letter"), new Item("A3 (29,7 × 42 cm)", "A3"), new Item("Legal (8.5 × 14 in)", "Legal") }, "Khổ giấy áp dụng khi bấm Chuẩn hóa dàn trang.");
            Lbl("Lề trang dọc (cm)");
            var rPort = NumGrid(new string[] { "Trên", "Dưới", "Trái", "Phải" }, out tbPtTop, out tbPtBot, out tbPtLeft, out tbPtRight);
            Lbl("Lề trang ngang (cm)");
            var rLand = NumGrid(new string[] { "Trên", "Dưới", "Trái", "Phải" }, out tbLnTop, out tbLnBot, out tbLnLeft, out tbLnRight);
            Lbl("Khoảng cách đầu và chân trang (cm)");
            NumPair("Header", out tbHdrDist, "Footer", out tbFtrDist);
            Lbl("Đoạn văn: cỡ chữ (pt)");
            NumPair("Times New Roman", out tbSzTnr, "Arial", out tbSzArial);
            Lbl("Đoạn văn: giãn dòng (lần) và khoảng cách đoạn (pt)");
            NumTriple("Giãn dòng", out tbLineSp, "Trên", out tbSpBef, "Dưới", out tbSpAft);
            Lbl("Đoạn văn: thụt lề (cm)");
            NumPair("Dòng đầu (First line)", out tbFirstLine, "Thụt treo (Hanging)", out tbHanging);
            cbParaAlign = Combo("Canh lề đoạn văn (nút Chuẩn hóa đoạn văn)", new Item[] { new Item("Hai bên (justify)", "justify"), new Item("Trái", "left"), new Item("Giữa", "center"), new Item("Phải", "right") },
                "Cách canh lề áp cho đoạn văn khi bấm Chuẩn hóa đoạn văn. Bôi đen nhiều đoạn thì áp cho tất cả các đoạn đó cùng lúc.");
            tip.SetToolTip(tbFirstLine, "Thụt dòng đầu (First Line Indent) cho đoạn thường, đơn vị cm (0 đến 5). Mặc định 1,25 cm. Không áp cho tiêu đề, đoạn trong bảng và đoạn danh sách.");
            tip.SetToolTip(tbHanging, "Thụt treo (Hanging Indent): nếu lớn hơn 0, đặt LeftIndent bằng giá trị này và FirstLineIndent âm (0 đến 5 cm). Mặc định 0.");
            Head("Thông báo");
            chDlg = Chk("Hiện hộp thoại báo kết quả sau mỗi lệnh", "Mặc định tắt: kết quả chỉ hiện trên thanh trạng thái của Word (góc dưới bên trái) và trong bảng điều khiển nếu đang mở. Bật nếu muốn có hộp thoại.");
            Head("Cửa sổ");
            chTop = Chk("Bảng điều khiển luôn ở trên cửa sổ khác", null);
            Head("AI trực tuyến");
            Lbl("Khóa API (tự nhận dạng Anthropic Claude sk-ant-..., Google Gemini AIza..., OpenAI sk-...)");
            tbKey = new TextBox(); tbKey.UseSystemPasswordChar = true; tbKey.Width = W - 20; tbKey.AccessibleName = "Khóa API";
            tip.SetToolTip(tbKey, "Dán khóa API rồi bấm Kiểm tra để tải danh sách model khả dụng. Khóa được mã hóa DPAPI (chỉ tài khoản Windows này đọc được), không lưu trong mã nguồn.");
            f.Controls.Add(tbKey);
            lblAiProv = Lbl("Nhà cung cấp: Chưa nhận dạng");
            Lbl("Model khả dụng (bấm Kiểm tra để nạp danh sách)");
            cbAiModel = new ComboBox(); cbAiModel.DropDownStyle = ComboBoxStyle.DropDownList; cbAiModel.Width = W - 20; cbAiModel.AccessibleName = "Chọn model AI";
            tip.SetToolTip(cbAiModel, "Bấm Kiểm tra trước để tải các model hỗ trợ từ nhà cung cấp.");
            f.Controls.Add(cbAiModel);
            Lbl("Tham số sinh văn bản: Temperature (0.0 – 2.0) / Số token tối đa (64 – 8192)");
            NumPair("Temperature", out tbAiTemp, "Max tokens", out tbAiMaxTok);
            tip.SetToolTip(tbAiTemp, "Độ sáng tạo từ 0.0 (ổn định kỹ thuật) đến 2.0. Mặc định 0.3.");
            tip.SetToolTip(tbAiMaxTok, "Số token đầu ra tối đa mỗi lần gọi (64 đến 8192). Mặc định 1024.");
            lblKey = Lbl("");
            Button[] aiBtns = BtnRowRet(new string[] { "Kiểm tra", "Kiểm tra & Lưu", "Xóa khóa" },
                new Action[] { delegate { CheckKeyModels(); }, delegate { TestAndSaveKey(); }, delegate { ClearKey(); } });
            btnCheckAi = aiBtns[0]; btnTestSaveAi = aiBtns[1]; btnClearAi = aiBtns[2];
            Head("Trợ giúp");
            var help = new TextBox(); help.Multiline = true; help.ReadOnly = true; help.ScrollBars = ScrollBars.Vertical; help.Height = 420; help.Width = W - 10;
            help.Text = HelpText; help.AccessibleName = "Hướng dẫn sử dụng"; help.BackColor = Color.White; help.TabStop = true; help.Select(0, 0);
            f.Controls.Add(help);
            BtnRow(new string[] { "Mở nhật ký" }, new Action[] { delegate { OpenLog(); } });
            LoadToUi();
            Reflow(scroll);
        }

        void Reflow(Panel scroll)
        {
            int avail = Math.Max(300, scroll.ClientSize.Width - f.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth - 10);
            W = avail;
            foreach (Control c in f.Controls)
            {
                if (c is Label || c is CheckBox) c.MaximumSize = new Size(avail, 0);
                else if (c is TableLayoutPanel) c.Width = avail - 10;
                else c.Width = avail - 10;
            }
        }

        Label Lbl(string t)
        {
            var l = new Label(); l.Text = t; l.AutoSize = true; l.MaximumSize = new Size(W, 0); l.Margin = new Padding(3, 10, 3, 2); f.Controls.Add(l); return l;
        }
        Label Head(string t) { Label l = Lbl(t); l.Font = new Font(Font, FontStyle.Bold); l.Margin = new Padding(3, 18, 3, 2); return l; }
        ComboBox Combo(string name, Item[] items, string tipText)
        {
            Lbl(name);
            var c = new ComboBox(); c.DropDownStyle = ComboBoxStyle.DropDownList; c.Width = W - 20; c.Items.AddRange(items); c.SelectedIndex = 0; c.AccessibleName = name;
            if (tipText != null) { tip.SetToolTip(c, tipText); c.AccessibleDescription = tipText; }
            f.Controls.Add(c); return c;
        }
        CheckBox Chk(string t, string tipText)
        {
            var c = new CheckBox(); c.Text = t; c.AutoSize = true; c.MaximumSize = new Size(W, 0); c.Margin = new Padding(3, 8, 3, 2); c.AccessibleName = t;
            if (tipText != null) { tip.SetToolTip(c, tipText); c.AccessibleDescription = tipText; }
            f.Controls.Add(c); return c;
        }
        void BtnRow(string[] texts, Action[] acts) { BtnRowRet(texts, acts); }

        Button[] BtnRowRet(string[] texts, Action[] acts)
        {
            var t = new TableLayoutPanel(); t.ColumnCount = texts.Length; t.RowCount = 1; t.Height = 48; t.Width = W - 10; t.Margin = new Padding(3, 6, 3, 2);
            for (int i = 0; i < texts.Length; i++) t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / texts.Length));
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            var list = new Button[texts.Length];
            for (int i = 0; i < texts.Length; i++)
            {
                var b = new Button(); b.Text = texts[i]; b.Dock = DockStyle.Fill; b.AccessibleName = texts[i]; Action a = acts[i]; b.Click += delegate { a(); };
                t.Controls.Add(b, i, 0);
                list[i] = b;
            }
            f.Controls.Add(t);
            return list;
        }

        // Rev0.9.5.6: các hàng "nhãn + ô nhập" dùng cột nhãn AutoSize và hàng AutoSize, nên không chồng chữ ở mọi mức thu phóng và DPI.
        TableLayoutPanel NumRow(params string[] labels)
        {
            var g = new TableLayoutPanel();
            g.ColumnCount = labels.Length * 2; g.RowCount = 1;
            g.AutoSize = true; g.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            g.Width = W - 10; g.Margin = new Padding(3, 2, 3, 6);
            for (int i = 0; i < labels.Length; i++)
            {
                g.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                g.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / labels.Length));
            }
            g.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            f.Controls.Add(g);
            return g;
        }
        void NumPair(string l1, out TextBox t1, string l2, out TextBox t2)
        {
            var g = NumRow(l1, l2);
            t1 = CellNum(g, l1, 0); t2 = CellNum(g, l2, 2);
        }
        void NumTriple(string l1, out TextBox t1, string l2, out TextBox t2, string l3, out TextBox t3)
        {
            var g = NumRow(l1, l2, l3);
            t1 = CellNum(g, l1, 0); t2 = CellNum(g, l2, 2); t3 = CellNum(g, l3, 4);
        }
        TableLayoutPanel NumGrid(string[] ls, out TextBox t1, out TextBox t2, out TextBox t3, out TextBox t4)
        {
            var g = NumRow(ls);
            t1 = CellNum(g, ls[0], 0); t2 = CellNum(g, ls[1], 2); t3 = CellNum(g, ls[2], 4); t4 = CellNum(g, ls[3], 6);
            return g;
        }
        static TextBox CellNum(TableLayoutPanel g, string lbl, int col)
        {
            var l = new Label { Text = lbl, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 6, 6), UseMnemonic = false };
            int minW = TextRenderer.MeasureText("00000", g.Font).Width + 16;
            var tb = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Margin = new Padding(0, 3, 10, 3), MinimumSize = new Size(minW, 0), AccessibleName = lbl };
            g.Controls.Add(l, col, 0); g.Controls.Add(tb, col + 1, 0);
            return tb;
        }
        static double ParseD(TextBox tb, double def, double lo, double hi)
        {
            double v;
            string s = (tb.Text ?? "").Trim().Replace(',', '.');
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return def;
            return v < lo ? lo : (v > hi ? hi : v);
        }
        static string FmtD(double v) { return v.ToString("0.##", CultureInfo.InvariantCulture); }

        void LoadToUi()
        {
            UiUtil.Pick(cbNumStyle, st.NumStyle); UiUtil.Pick(cbNumDec, st.NumDec); UiUtil.Pick(cbCalcDec, st.CalcDec); UiUtil.Pick(cbCalcFmt, st.CalcFmtKind); UiUtil.Pick(cbSciDec, st.SciDec.ToString()); UiUtil.Pick(cbUnitNames, st.UnitNames);
            UiUtil.Pick(cbPagePaper, st.PagePaper);
            tbPtTop.Text = FmtD(st.PgPortTop); tbPtBot.Text = FmtD(st.PgPortBot); tbPtLeft.Text = FmtD(st.PgPortLeft); tbPtRight.Text = FmtD(st.PgPortRight);
            tbLnTop.Text = FmtD(st.PgLandTop); tbLnBot.Text = FmtD(st.PgLandBot); tbLnLeft.Text = FmtD(st.PgLandLeft); tbLnRight.Text = FmtD(st.PgLandRight);
            tbHdrDist.Text = FmtD(st.PgHeaderDist); tbFtrDist.Text = FmtD(st.PgFooterDist);
            tbSzTnr.Text = FmtD(st.ParaSizeTnr); tbSzArial.Text = FmtD(st.ParaSizeArial);
            tbLineSp.Text = FmtD(st.ParaLineSpace); tbSpBef.Text = FmtD(st.ParaSpaceBefore); tbSpAft.Text = FmtD(st.ParaSpaceAfter);
            tbFirstLine.Text = FmtD(st.ParaFirstLineCm); tbHanging.Text = FmtD(st.ParaHangingCm);
            UiUtil.Pick(cbParaAlign, st.ParaAlign);
            chDlg.Checked = st.ShowDialog;
            chThou.Checked = st.NumThou; chAddDec.Checked = st.NumAddDecToInt; chTop.Checked = st.TopMost; chClear.Checked = st.ClearManual;
            tbAiTemp.Text = FmtD(st.AiTemp);
            tbAiMaxTok.Text = st.AiMaxTokens.ToString(CultureInfo.InvariantCulture);
            PopulateModelList(st.ApiModel.Length > 0 ? new List<string> { st.ApiModel } : new List<string>(), st.ApiModel);
            ShowKeyState();
        }

        void PopulateModelList(List<string> models, string prefer)
        {
            cbAiModel.Items.Clear();
            foreach (string m in models)
                if (!string.IsNullOrEmpty(m)) cbAiModel.Items.Add(new Item(m, m));
            if (cbAiModel.Items.Count > 0)
            {
                if (!string.IsNullOrEmpty(prefer)) UiUtil.Pick(cbAiModel, prefer);
                if (cbAiModel.SelectedIndex < 0) cbAiModel.SelectedIndex = 0;
            }
        }

        string ResolveEffectiveKey()
        {
            string typed = (tbKey.Text ?? "").Trim();
            if (typed.Length > 0) return typed;
            return Secret.Unprotect(st.ApiKeyEnc);
        }

        bool TryReadAiParams(out double temp, out int maxTok, out string err)
        {
            err = "";
            temp = st.AiTemp; maxTok = st.AiMaxTokens;
            string ts = (tbAiTemp.Text ?? "").Trim().Replace(',', '.');
            string ms = (tbAiMaxTok.Text ?? "").Trim();
            if (!double.TryParse(ts, NumberStyles.Float, CultureInfo.InvariantCulture, out temp))
            {
                err = "Temperature không phải là số hợp lệ (cần từ 0.0 đến 2.0).";
                return false;
            }
            if (!int.TryParse(ms, NumberStyles.Integer, CultureInfo.InvariantCulture, out maxTok))
            {
                err = "Số token tối đa phải là số nguyên (từ 64 đến 8192).";
                return false;
            }
            return AiClient.ValidateParams(temp, maxTok, out err);
        }

        /// <summary>Ghi các thông số của cửa sổ này lên tệp mới nhất (không đụng thông số của Ribbon và bảng điều khiển).</summary>
        void Save()
        {
            AppSettings s = AppSettings.Load();
            s.NumStyle = UiUtil.Val(cbNumStyle); s.NumDec = UiUtil.Val(cbNumDec); s.CalcDec = UiUtil.Val(cbCalcDec); s.CalcFmtKind = UiUtil.Val(cbCalcFmt);
            { int sd; if (int.TryParse(UiUtil.Val(cbSciDec), out sd)) s.SciDec = Math.Max(0, Math.Min(5, sd)); }
            s.UnitNames = UiUtil.Val(cbUnitNames) == "common" ? "common" : "law";
            s.PagePaper = UiUtil.Val(cbPagePaper);
            s.PgPortTop = ParseD(tbPtTop, s.PgPortTop, 0.2, 10.0); s.PgPortBot = ParseD(tbPtBot, s.PgPortBot, 0.2, 10.0);
            s.PgPortLeft = ParseD(tbPtLeft, s.PgPortLeft, 0.2, 10.0); s.PgPortRight = ParseD(tbPtRight, s.PgPortRight, 0.2, 10.0);
            s.PgLandTop = ParseD(tbLnTop, s.PgLandTop, 0.2, 10.0); s.PgLandBot = ParseD(tbLnBot, s.PgLandBot, 0.2, 10.0);
            s.PgLandLeft = ParseD(tbLnLeft, s.PgLandLeft, 0.2, 10.0); s.PgLandRight = ParseD(tbLnRight, s.PgLandRight, 0.2, 10.0);
            s.PgHeaderDist = ParseD(tbHdrDist, s.PgHeaderDist, 0.2, 5.0); s.PgFooterDist = ParseD(tbFtrDist, s.PgFooterDist, 0.2, 5.0);
            s.ParaSizeTnr = ParseD(tbSzTnr, s.ParaSizeTnr, 6.0, 72.0); s.ParaSizeArial = ParseD(tbSzArial, s.ParaSizeArial, 6.0, 72.0);
            s.ParaLineSpace = ParseD(tbLineSp, s.ParaLineSpace, 0.8, 3.0);
            s.ParaSpaceBefore = ParseD(tbSpBef, s.ParaSpaceBefore, 0.0, 72.0); s.ParaSpaceAfter = ParseD(tbSpAft, s.ParaSpaceAfter, 0.0, 72.0);
            s.ParaFirstLineCm = ParseD(tbFirstLine, s.ParaFirstLineCm, 0.0, 5.0); s.ParaHangingCm = ParseD(tbHanging, s.ParaHangingCm, 0.0, 5.0);
            s.ParaAlign = UiUtil.Val(cbParaAlign);
            s.AiTemp = ParseD(tbAiTemp, s.AiTemp, 0.0, 2.0);
            s.AiMaxTokens = (int)Math.Round(ParseD(tbAiMaxTok, s.AiMaxTokens, 64, 8192));
            if (cbAiModel.SelectedItem is Item) s.ApiModel = ((Item)cbAiModel.SelectedItem).Value;
            s.ShowDialog = chDlg.Checked;
            s.NumThou = chThou.Checked; s.NumAddDecToInt = chAddDec.Checked; s.TopMost = chTop.Checked; s.ClearManual = chClear.Checked;
            s.Save(); st = s;
        }

        void SetAiBusy(bool busy)
        {
            aiBusy = busy;
            btnCheckAi.Enabled = !busy;
            btnTestSaveAi.Enabled = !busy;
            btnClearAi.Enabled = !busy;
        }

        void ShowKeyState()
        {
            string k = Secret.Unprotect(st.ApiKeyEnc);
            string prov = !string.IsNullOrEmpty(st.ApiProvider) ? st.ApiProvider : AiClient.DetectProvider(k);
            lblAiProv.Text = "Nhà cung cấp: " + AiClient.ProviderDisplayName(prov);
            if (st.ApiKeyEnc.Length == 0)
                lblKey.Text = "Chưa lưu khóa API.";
            else if (k.Length > 0)
                lblKey.Text = "Đã lưu khóa " + Secret.Mask(k) + (st.ApiModel.Length > 0 ? (" | Model: " + st.ApiModel) : " (chưa chọn model)");
            else
                lblKey.Text = "Khóa đã lưu không giải mã được trên máy này, hãy nhập lại.";
        }

        void CheckKeyModels()
        {
            if (aiBusy) return;
            string k = ResolveEffectiveKey();
            if (string.IsNullOrWhiteSpace(k)) { lblKey.Text = "Hãy dán khóa API trước khi bấm Kiểm tra."; return; }
            string detected = AiClient.DetectProvider(k);
            lblAiProv.Text = "Nhà cung cấp: " + AiClient.ProviderDisplayName(detected);
            if (string.IsNullOrEmpty(detected))
            {
                lblKey.Text = "Không nhận dạng được nhà cung cấp từ khóa (cần sk-ant-..., AIza..., hoặc sk-...).";
                return;
            }
            SetAiBusy(true);
            lblKey.Text = "Đang kết nối " + AiClient.ProviderDisplayName(detected) + " để tải danh sách model...";
            ThreadPool.QueueUserWorkItem(delegate
            {
                AiModelListResult res = AiClient.ListModels(k);
                try
                {
                    if (!IsDisposed && IsHandleCreated)
                    {
                        BeginInvoke(new MethodInvoker(delegate
                        {
                            SetAiBusy(false);
                            if (res.Ok)
                            {
                                PopulateModelList(res.Models, st.ApiModel);
                                lblAiProv.Text = "Nhà cung cấp: " + AiClient.ProviderDisplayName(res.Provider);
                                lblKey.Text = "Tìm thấy " + res.Models.Count + " model khả dụng. Hãy chọn model rồi bấm Kiểm tra & Lưu.";
                            }
                            else
                            {
                                lblKey.Text = "Kiểm tra thất bại: " + res.Error;
                            }
                        }));
                    }
                }
                catch { }
            });
        }

        void TestAndSaveKey()
        {
            if (aiBusy) return;
            string k = ResolveEffectiveKey();
            if (string.IsNullOrWhiteSpace(k)) { lblKey.Text = "Chưa có khóa API để kiểm tra và lưu."; return; }
            string prov = AiClient.DetectProvider(k);
            if (string.IsNullOrEmpty(prov)) { lblKey.Text = "Định dạng khóa không hợp lệ (hỗ trợ sk-ant-..., AIza..., sk-...)."; return; }
            Item selModel = cbAiModel.SelectedItem as Item;
            string model = selModel != null ? selModel.Value : "";
            if (string.IsNullOrWhiteSpace(model)) { lblKey.Text = "Chưa có model nào được chọn. Hãy bấm Kiểm tra để tải danh sách model trước."; return; }

            double temp; int maxTok; string pErr;
            if (!TryReadAiParams(out temp, out maxTok, out pErr)) { lblKey.Text = pErr; return; }

            SetAiBusy(true);
            lblKey.Text = "Đang thử gọi model \"" + model + "\"...";
            ThreadPool.QueueUserWorkItem(delegate
            {
                AiChatResult res = AiClient.TestConnection(k, model, temp, maxTok);
                try
                {
                    if (!IsDisposed && IsHandleCreated)
                    {
                        BeginInvoke(new MethodInvoker(delegate
                        {
                            SetAiBusy(false);
                            if (res.Ok)
                            {
                                try
                                {
                                    AppSettings s = AppSettings.Load();
                                    s.ApiProvider = res.Provider;
                                    s.ApiModel = model;
                                    s.ApiKeyEnc = Secret.Protect(k);
                                    s.AiTemp = temp;
                                    s.AiMaxTokens = maxTok;
                                    s.Save(); st = s;
                                    tbKey.Text = "";
                                    ShowKeyState();
                                    lblKey.Text = "Kết nối thành công và đã lưu an toàn (" + AiClient.ProviderDisplayName(res.Provider) + " / " + model + ").";
                                }
                                catch (Exception ex) { lblKey.Text = "Gọi model OK nhưng không lưu được khóa: " + ex.Message; }
                            }
                            else
                            {
                                lblKey.Text = "Thử nghiệm model thất bại: " + res.Error;
                            }
                        }));
                    }
                }
                catch { }
            });
        }

        void ClearKey()
        {
            AppSettings s = AppSettings.Load();
            s.ApiKeyEnc = ""; s.ApiProvider = ""; s.ApiModel = "";
            s.Save(); st = s;
            tbKey.Text = ""; cbAiModel.Items.Clear();
            ShowKeyState();
        }

        void OpenLog()
        {
            try { string d = Path.Combine(Path.GetTempPath(), "TroLyKyThuat"); Directory.CreateDirectory(d); Process.Start("explorer.exe", d); }
            catch (Exception ex) { lblKey.Text = "Không mở được thư mục nhật ký: " + ex.Message; }
        }

        const string HelpText =
"HƯỚNG DẪN NHANH\r\n" +
"Mọi lệnh nằm trên Ribbon, thẻ My tab. Bảng điều khiển (nút Bảng điều khiển) dùng để rà soát và soạn phương trình.\r\n\r\n" +
"BẢNG\r\n" +
"Mọi bảng, Bảng này: áp dụng master style. Chỉ ghi thuộc tính đang khác nên bảng đã đúng thì chạy gần như tức thì. Một lần Ctrl+Z hoàn tác cả lần áp dụng. Bảng này chỉ đổi bảng đang chọn (không đổi các bảng khác, kể cả màu nền dòng tiêu đề); Mọi bảng đổi tất cả.\r\n" +
"Lặp tiêu đề: bật hoặc tắt lặp dòng tiêu đề ở mỗi trang (kèm không ngắt hàng giữa hai trang) cho mọi bảng nằm trong phần đang chọn: con trỏ trong một bảng thì bảng đó, bôi đen qua nhiều bảng thì tất cả. Mũi tên bên dưới: bật hoặc tắt cho mọi bảng.\r\n" +
"Tiêu đề (màu nền, màu chữ, căn chữ), Viền (độ dày, màu, kiểu nét), Tô xen kẽ (bật tắt, màu): đặt thông số, sau đó bấm Mọi bảng hoặc Bảng này để áp dụng. Mỗi nút màu mở trực tiếp bảng màu thống nhất kiểu Word hiện đại gồm Tự động, Màu giao diện (10 × 6), Màu chuẩn (10) và Thêm màu... Mã màu ghi vào tài liệu là RGB nên máy không cài add-in vẫn hiển thị đúng màu.\r\n" +
"DÀN TRANG (Rev0.9.3)\r\n" +
"Nhóm Dàn trang nằm ngoài cùng bên trái thẻ Trợ lý Kỹ thuật: Chuẩn hóa dàn trang (khổ giấy, căn lề trang dọc/ngang, đầu/chân trang theo thông số trong Cài đặt), Chuẩn hóa đoạn văn (chỉ đoạn tại con trỏ, căn đều hai bên, Times New Roman 13 pt, Arial 11 pt, font khác giữ nguyên cỡ; giãn dòng 1,2; trên 6 dưới 6; thụt dòng đầu 1,25 cm hoặc thụt treo cho đoạn thường), Thêm trang ngang / Thêm trang dọc chèn 1 trang mới định hướng độc lập ngay sau trang chứa con trỏ.\r\n\r\n" +
"Ký tự trong bảng: font (danh sách là các font đã cài trong máy, hoặc gõ tên font), cỡ chữ, khoảng cách đoạn trong ô, căn ô chữ, căn ô số.\r\n" +
"Xóa định dạng chỉnh tay khi áp dụng: bật tắt ở cửa sổ Cài đặt này.\r\n\r\n" +
"HÀNG VÀ CỘT\r\n" +
"Thêm hàng, Thêm cột, Xóa: mỗi hàng hoặc cột đang chọn được xử lý riêng. Giữ Ctrl để chọn nhiều hàng, cột rời nhau. Ô Thêm cùng hai mũi tên là số hàng, cột thêm cho mỗi hàng, cột.\r\n" +
"Tính: chọn dãy ô số rồi chọn phép tính. Dãy ngang thì điền kết quả vào cột bên phải, dãy dọc thì điền vào hàng bên dưới.\r\n" +
"Xuất: CSV ra tệp, Excel sao chép vào Clipboard để dán vào Excel.\r\n\r\n" +
"ĐỊNH DẠNG SỐ\r\n" +
"Bôi đen một dãy ô, một cột, một hàng, hoặc các số rải rác trong văn bản rồi bấm.\r\n" +
"Nút Dấu: đổi kiểu theo vòng Việt (1.234,56), Mỹ (1,234.56), SI (1 234,56). Thêm số lẻ, Bớt số lẻ: tăng hoặc giảm một chữ số thập phân, tối đa 5.\r\n" +
"Đồng bộ số lẻ: mọi số đang bôi đen có đúng 0 đến 5 chữ số thập phân (số ít hơn được thêm số 0, số nhiều hơn được làm tròn). Số ở góc biểu tượng là mức đang dùng; bấm mũi tên bên dưới để chọn 0 đến 5, hoặc chọn Viết kiểu kỹ thuật: số được viết dạng 2,3 × 10⁶ hay 3,67 × 10³ (số mũ là ký tự chỉ số trên, nên không bị nhận lại thành số). Số chữ số thập phân của kiểu kỹ thuật chỉnh ở mục Số của cửa sổ này. Số đã viết kiểu kỹ thuật được nhận lại và viết lại theo mức mới.\r\n" +
"Đọc số: đặt con trỏ trong hoặc sát một số rồi bấm, chữ được viết ngay sau số trong ngoặc: 1.234 (một nghìn hai trăm ba mươi bốn). Bôi đen nhiều số thì đọc từng số. Quy ước: hai mươi mốt, hai mươi lăm, mười lăm, một trăm linh năm, bốn (không dùng tư), số thập phân đọc từng chữ số sau chữ phẩy (3,05 là ba phẩy không năm), số âm đọc âm, 2,3 × 10⁶ đọc hai phẩy ba nhân mười mũ sáu. Đơn vị đứng ngay sau số cũng được đọc (20 kPa thành hai mươi kilôpascan, 1.000.000 VNĐ thành một triệu Việt Nam đồng), tên theo Nghị định 134/2007/NĐ-CP, có thể đổi sang tên thông dụng (kilopascal, newton) trong mục Số. Bấm lần nữa không viết thêm lần hai.\r\n" +
"Cả bốn nút giữ nguyên vùng bôi đen nên có thể bấm liên tiếp nhiều định dạng trong một lần bôi đen (vùng chọn rời nhau bằng Ctrl thì Word không cho dựng lại).\r\n\r\n" +
"CHUẨN HÓA KỸ THUẬT\r\n" +
"Nút Chuẩn hóa: đặt con trỏ vào hoặc bôi đen một từ, công thức bị lỗi mã (R_cyl, R_{cyl}, t_req, \\Delta P, \\frac{a}{b}, x_i^2) rồi bấm. Sửa ngay về dạng chuẩn, không cần quét cả văn bản: chữ biến in nghiêng, chỉ số dưới đứng thẳng, bỏ các dấu _ { } thừa. Công thức riêng một dòng hoặc có phân số được dựng thành Equation của Word.\r\n" +
"Rà soát: quét văn bản tìm công thức hóa học, toán, ký hiệu, đơn vị, số kỹ thuật chưa chuẩn. Mục đang xem được tô vàng trong văn bản. Bấm ✔ cuối dòng để sửa, ✖ để bỏ qua. Sửa tương tự, Bỏ qua tương tự: áp dụng cho mọi chỗ giống mục đang chọn; với số thì cùng dạng là đủ (1.1 và 2.3; 5.678 và 12.345). Sửa hết không đụng tới mục Nghi ngờ.\r\n" +
"Dấu phân cách số khi rà soát chọn riêng ở thẻ Chuẩn hóa kỹ thuật, không ảnh hưởng nút Dấu trên Ribbon.\r\n" +
"Rà soát tự bỏ qua số thứ tự đầu dòng (1.1., 2.1.), số của bảng và hình (Bảng 2.1, Hình 3), năm, số trang, số hiệu, ngày. Đồng thời báo chỗ thiếu dấu cách giữa số và đơn vị (100kPa).\r\n" +
"Kiểu chỉ số: bôi đen công thức hoặc cả đoạn rồi bấm: đổi luân phiên giữa Equation và font thường có chỉ số trên, dưới. Bấm lần nữa để quay lại.\r\n" +
"Phương trình: gõ phương trình, tự nhận hóa học hay toán. Chọn Phân số đứng để tử trên, mẫu dưới. Sau khi chèn, con trỏ quay lại Word.\r\n" +
"Gõ nhanh: H2O, SO4^2-, Fe^3+, x^2, t_{req}, P_CO2, \\Delta G, \\alpha, sqrt(x), (a+b)/c, \\frac{a}{b}.\r\n\r\n" +
"NGÔN NGỮ, ĐỒNG NGHĨA & AI TRỰC TUYẾN (Rev0.9.3)\r\n" +

"Gợi ý từ: hiển thị danh sách từ đồng nghĩa (hơn 100 từ Việt + 100 từ Anh, tham chiếu chéo theo nhóm); nhấp đúp để thay thế giữ nguyên định dạng ký tự.\r\n" +
"AI trực tuyến: dán khóa API trong Cài đặt -> bấm Kiểm tra để tự nhận dạng nhà cung cấp và tải model -> chọn model, chỉnh Temperature / Max tokens -> bấm Kiểm tra & Lưu (mã hóa DPAPI theo người dùng Windows). Trong thẻ Công cụ AI, chọn Viết lại / Dịch cho đoạn bôi đen tối đa 500 từ hoặc Văn phong / Ngữ pháp để quét phần bôi đen hay toàn bộ nội dung chính. Gợi ý từ đồng nghĩa và Tra thuật ngữ cũng dùng model đã chọn. Tổng token phiên hiển thị ở chân bảng điều khiển và tự reset khi khởi động lại Word.\r\n\r\n" +
"KHI CÓ SỰ CỐ\r\n" +
"Nhật ký nằm ở %TEMP%\\TroLyKyThuat\\log.txt. Chụp màn hình thông báo lỗi và gửi lại.";
    }
}
