// Trợ lý Kỹ thuật - Rev0.5: bảng điều khiển (WinForms). Cỡ chữ lớn, đặt tên truy cập cho mọi ô để trình đọc màn hình đọc được.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace TroLyKyThuat
{
    sealed class Item
    {
        public string Text; public string Value;
        public bool Raw;   // Rev0.9.6: true thì hiện đúng chữ gốc, không dịch (tên ngôn ngữ giao diện)
        public Item(string t, string v) { Text = t; Value = v; }
        public override string ToString() { return Raw ? Text : Lang.Msg(Text); }
    }

    static class UiUtil
    {
        public static void Pick(ComboBox c, string value)
        {
            for (int i = 0; i < c.Items.Count; i++) if (((Item)c.Items[i]).Value == value) { c.SelectedIndex = i; return; }
        }
        public static string Val(ComboBox c) { return ((Item)c.SelectedItem).Value; }
    }

    public sealed partial class PaneForm : Form
    {
        readonly WordReview rv;
        readonly WordCompose wc;
        AppSettings st = AppSettings.Load();
        DateTime savedAt = DateTime.MinValue;
        public bool AllowClose;
        bool loading;

        TextBox status;
        Label lblTokenFooter;
        TabControl main;
        TabPage pRv, pCt, pRw;
        readonly List<Panel> scrolls = new List<Panel>();
        readonly List<FlowLayoutPanel> flows = new List<FlowLayoutPanel>();
        readonly ToolTip tip = new ToolTip();
        System.Windows.Forms.Timer langTimer;
        int W = 440;

        // thẻ Viết lại (AI)
        ComboBox cbRwTone;
        TextBox tbRwOrig, tbRwOut;
        Label lblRwCount;
        Button btnRwRead, btnRwRun, btnRwApply, btnRwUndo;
        CheckBox chkRwRewrite, chkRwTranslate;
        ComboBox cbRwLang;
        int rwInBase, rwOutBase;
        RewriteSnapshot rwSnap, rwResultSnap;
        string rwResultTarget;
        bool rwBusy;

        // thẻ Chuẩn hóa
        ComboBox cbScope, cbFix;
        CheckBox chRvChem, chRvMath, chRvSym, chRvNum, chRvNumTbl, chGlyph;
        ComboBox cbRvNumStyle;
        ListView lv;
        TextBox tbReason;
        Label lblRv;
        Button btnFixSim, btnIgnSim, btnFixAll, btnGoto;
        // thẻ Phương trình
        TextBox tbCt;
        Panel ctBox;
        Label lblKind, lblCtMsg;
        ComboBox cbOpt;
        EquationPreview inlinePreview, displayPreview;
        CheckBox insertInline, insertDisplay;
        ComposeResult cur;
        bool ctBusy;

        public static string GetTitle(Version v = null)
        {
            if (v == null) v = Assembly.GetExecutingAssembly().GetName().Version;
            if (v == null) return "Trợ lý Kỹ thuật - Rev0.9.5.7";
            string ver = (v.Revision > 0)
                ? string.Format("{0}.{1}.{2}.{3}", v.Major, v.Minor, v.Build, v.Revision)
                : string.Format("{0}.{1}.{2}", v.Major, v.Minor, v.Build);
            return "Trợ lý Kỹ thuật - Rev" + ver;
        }

        public PaneForm(object word)
        {
            rv = new WordReview(word);
            wc = new WordCompose(word);
            Text = GetTitle();
            DpiUtil.ConfigureForm(this, 11f);
            Width = 560; Height = 900;
            MinimumSize = new Size(480, 600);
            StartPosition = FormStartPosition.Manual;
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            Left = wa.Right - Width; Top = wa.Top; Height = wa.Height;
            ShowInTaskbar = true;
            tip.AutoPopDelay = 20000; tip.InitialDelay = 400; tip.ShowAlways = true;

            lblTokenFooter = new Label();
            lblTokenFooter.Dock = DockStyle.Bottom;
            lblTokenFooter.Height = 24;
            lblTokenFooter.BackColor = Color.FromArgb(245, 247, 250);
            lblTokenFooter.ForeColor = Color.FromArgb(75, 85, 99);
            lblTokenFooter.Font = DpiUtil.CreateUiFont(9f, FontStyle.Regular);
            lblTokenFooter.TextAlign = ContentAlignment.MiddleLeft;
            lblTokenFooter.Padding = new Padding(6, 0, 6, 0);
            lblTokenFooter.Text = Lang.Msg(AiSessionUsage.FormatFooter());
            lblTokenFooter.AccessibleName = "Số token AI đã dùng trong phiên Word hiện tại";
            Controls.Add(lblTokenFooter);

            status = new TextBox();
            status.Multiline = true; status.ReadOnly = true; status.ScrollBars = ScrollBars.Vertical;
            status.Dock = DockStyle.Bottom; status.Height = 84; status.BackColor = Color.FromArgb(238, 245, 253);
            status.AccessibleName = "Thông báo kết quả"; status.AccessibleRole = AccessibleRole.StaticText;
            status.Text = Lang.T("Sẵn sàng.");
            Controls.Add(status);

            main = new TabControl(); main.Dock = DockStyle.Fill; main.AccessibleName = "Nhóm chức năng"; main.ShowToolTips = true;
            Controls.Add(main); main.BringToFront();
            pRv = new TabPage("Chuẩn hóa kỹ thuật"); pRv.ToolTipText = "Rà soát công thức, đơn vị, số kỹ thuật";
            pCt = new TabPage("Phương trình"); pCt.ToolTipText = "Soạn phương trình hóa học, toán, lý";
            pRw = new TabPage("Công cụ AI"); pRw.ToolTipText = "Dùng AI cho thuật ngữ, văn phong, ngữ pháp, từ đồng nghĩa, viết lại và dịch";
            main.TabPages.Add(pRv); main.TabPages.Add(pCt); main.TabPages.Add(pRw);
            BuildReview(pRv);
            scopeBoxes.Add(cbScope);
            cbScope.SelectedIndexChanged += delegate { SyncScope(cbScope); };
            BuildTextTabs(word);
            BuildCompose(pCt);
            AiSessionUsage.Changed += OnAiUsageChanged;
            main.SelectedIndexChanged += delegate { if (main.SelectedTab == pRw) RefreshAiNotice(); };
            LoadToUi();
            ReflowAll();
            Lang.Apply(this, tip);   // Rev0.9.6: dịch giao diện theo ngôn ngữ đã chọn
            // nhãn và danh sách đổi nội dung lúc chạy (kết quả quét, bộ đếm từ...): dịch lại định kỳ khi đang ở tiếng Anh
            langTimer = new System.Windows.Forms.Timer(); langTimer.Interval = 400;
            langTimer.Tick += delegate { if (Lang.IsEn && Visible && !IsDisposed) Lang.Apply(this, tip); };
            langTimer.Start();
            FormClosed += delegate { try { langTimer.Stop(); langTimer.Dispose(); } catch { } };
            Resize += delegate { ReflowAll(); };
            // Rev0.9.5.2: nếu màn hình đang hiện có tỷ lệ khác bố cục (ví dụ dựng ở màn 100% rồi hiện ở màn 125-200%), co giãn chữ và nút cho đúng
            Shown += delegate { DpiUtil.SyncDpi(this); Log.W("Bảng điều khiển hiện: " + DpiUtil.Describe(this)); };
            UiZoom.Attach(this, main, "pane", ReflowAll);
            VisibleChanged += delegate { if (!Visible) ClearAllMarksQuiet(); };
            Activated += delegate { ReloadIfChanged(); };
            Deactivate += delegate { SaveSettings(); };
            FormClosed += delegate { AiSessionUsage.Changed -= OnAiUsageChanged; };
            FormClosing += delegate(object s, FormClosingEventArgs e)
            {
                SaveSettings(); ClearAllMarksQuiet();
                if (!AllowClose && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); }
            };
        }

        /// <summary>Rev0.9.5.2: theo dõi DPI từng màn hình (kéo cửa sổ sang màn hình khác tỷ lệ thì chữ và nút co giãn theo, không bị nhỏ hoặc mờ).</summary>
        protected override void WndProc(ref Message m)
        {
            if (DpiUtil.HandleDpiMessage(this, ref m)) return;
            base.WndProc(ref m);
        }

        public void ShowAndFocus(string tab, bool scan)
        {
            if (!Visible) Show();
            if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
            if (tab == "rv") main.SelectedTab = pRv;
            else if (tab == "ct") { main.SelectedTab = pCt; tbCt.Focus(); }
            Activate();
            ReloadIfChanged();
            if (tab == "ai" || tab == "term" || tab == "style" || tab == "gram" || tab == "syn") { OpenTextTab(tab); return; }
            if (scan) DoScan();
        }

        /// <summary>Mở cửa sổ Cài đặt riêng (chạy trên luồng của bảng điều khiển).</summary>
        public void ShowSettingsDialog()
        {
            SaveSettings();
            using (DpiUtil.EnterPerMonitorV2())
            using (var f = new SettingsForm()) f.ShowDialog();
            ReloadIfChanged();
        }

        /// <summary>Rev0.9.6: đổi ngôn ngữ giao diện khi bảng điều khiển đang mở: dịch lại mọi nhãn, chú thích, mục chọn từ bản gốc tiếng Việt.</summary>
        public void ApplyLanguage()
        {
            try
            {
                Lang.Apply(this, tip);
                if (lblTokenFooter != null) lblTokenFooter.Text = Lang.Msg(AiSessionUsage.FormatFooter());
                if (status != null && (status.Text == "Sẵn sàng." || status.Text == "Ready.")) status.Text = Lang.T("Sẵn sàng.");
                ReflowAll();
            }
            catch (Exception ex) { Log.W("ApplyLanguage loi: " + ex.Message); }
        }

        void OnAiUsageChanged()
        {
            try
            {
                if (IsDisposed || !IsHandleCreated) return;
                BeginInvoke(new MethodInvoker(delegate { if (!IsDisposed && lblTokenFooter != null) lblTokenFooter.Text = Lang.Msg(AiSessionUsage.FormatFooter()); }));
            }
            catch { }
        }

        /// <summary>Thông báo từ nút trên Ribbon: hiện kết quả và nạp lại cài đặt nếu Ribbon vừa đổi.</summary>
        public void Say(string msg)
        {
            msg = Lang.Msg(msg);
            status.Text = msg;
            status.AccessibleName = "Thông báo kết quả: " + msg;
            ReloadIfChanged();
        }

        // ------------------------------------------------------------ dựng giao diện
        FlowLayoutPanel Stack(Control page)
        {
            Panel scroll = new Panel(); scroll.Dock = DockStyle.Fill; scroll.AutoScroll = true;
            FlowLayoutPanel f = new FlowLayoutPanel();
            f.FlowDirection = FlowDirection.TopDown; f.WrapContents = false; f.AutoSize = true; f.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            f.Padding = new Padding(8);
            scroll.Controls.Add(f); page.Controls.Add(scroll);
            scrolls.Add(scroll); flows.Add(f);
            scroll.Resize += delegate { Reflow(scroll, f); };
            return f;
        }

        /// <summary>Cho mọi ô nhập, nút, nhãn dãn theo bề rộng khung hiện tại để nội dung giãn ra khi kéo rộng cửa sổ.</summary>
        void Reflow(Panel scroll, FlowLayoutPanel f)
        {
            int avail = Math.Max(240, scroll.ClientSize.Width - f.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth - 10);
            W = avail;
            foreach (Control c in f.Controls)
            {
                if (c is NumericUpDown) continue;
                if (c is Label || c is CheckBox) c.MaximumSize = new Size(avail, 0);
                else if (c is Button) { c.MinimumSize = new Size(avail, 40); c.MaximumSize = new Size(avail, 0); }
                else if (c is ListView)
                {
                    ListView lvx = (ListView)c; lvx.Width = avail;
                    lvx.Height = Math.Max(150, scroll.ClientSize.Height / 4);
                    if (lvx.Columns.Count == 6)
                    {
                        // Rev0.9.5.2: bề rộng cột cố định lấy từ chính các cột (đã co theo DPI), không dùng số cứng 96 + 84 + 40 + 40 của màn hình 100%
                        int fixedW = lvx.Columns[0].Width + lvx.Columns[1].Width + lvx.Columns[4].Width + lvx.Columns[5].Width + SystemInformation.VerticalScrollBarWidth + 8;
                        int rest = Math.Max((int)(120 * DpiUtil.LayoutScale(this)), avail - fixedW);
                        lvx.Columns[2].Width = rest / 2; lvx.Columns[3].Width = rest - rest / 2;
                    }
                }
                else c.Width = avail;
            }
        }

        void ReflowAll()
        {
            for (int i = 0; i < scrolls.Count; i++) Reflow(scrolls[i], flows[i]);
        }

        /// <summary>Gắn chú thích: hiện khi rê chuột và trình đọc màn hình đọc được như mô tả.</summary>
        T Tip<T>(T c, string text) where T : Control
        {
            tip.SetToolTip(c, text); c.AccessibleDescription = text; return c;
        }

        Label Lbl(FlowLayoutPanel f, string t)
        {
            Label l = new Label(); l.Text = t; l.AutoSize = true; l.MaximumSize = new Size(W, 0); l.Margin = new Padding(3, 10, 3, 2);
            f.Controls.Add(l); return l;
        }
        Label Head(FlowLayoutPanel f, string t)
        {
            Label l = Lbl(f, t); l.Font = new Font(Font, FontStyle.Bold); l.Margin = new Padding(3, 18, 3, 2); return l;
        }
        ComboBox Combo(FlowLayoutPanel f, string name, Item[] items, string tipText)
        {
            Lbl(f, name);
            ComboBox c = new ComboBox(); c.DropDownStyle = ComboBoxStyle.DropDownList; c.Width = W - 20;
            c.Items.AddRange(items); c.SelectedIndex = 0; c.AccessibleName = name;
            if (tipText != null) Tip(c, tipText);
            f.Controls.Add(c); return c;
        }
        CheckBox Chk(FlowLayoutPanel f, string t, string tipText)
        {
            CheckBox c = new CheckBox(); c.Text = t; c.AutoSize = true; c.MaximumSize = new Size(W, 0); c.Margin = new Padding(3, 8, 3, 2);
            c.AccessibleName = t; if (tipText != null) Tip(c, tipText); f.Controls.Add(c); return c;
        }
        /// <summary>Một hàng nút ngắn xếp ngang, mỗi nút chiếm đều bề rộng.</summary>
        Button[] BtnRow(FlowLayoutPanel f, string[] texts, Action[] acts, string[] tips)
        {
            TableLayoutPanel t = new TableLayoutPanel(); t.ColumnCount = texts.Length; t.RowCount = 1; t.Height = 48; t.Width = W - 10; t.Margin = new Padding(3, 6, 3, 2);
            for (int i = 0; i < texts.Length; i++) t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / texts.Length));
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            var bs = new Button[texts.Length];
            for (int i = 0; i < texts.Length; i++)
            {
                Button b = new Button(); b.Text = texts[i]; b.Dock = DockStyle.Fill; b.AccessibleName = texts[i]; Tip(b, tips[i]);
                Action a = acts[i]; b.Click += delegate { a(); };
                t.Controls.Add(b, i, 0); bs[i] = b;
            }
            f.Controls.Add(t);
            return bs;
        }

        // ------------------------------------------------------------ thẻ Chuẩn hóa
        void BuildReview(TabPage page)
        {
            FlowLayoutPanel f = Stack(page);
            cbScope = Combo(f, "Phạm vi", new Item[] { new Item("Toàn văn bản", "all"), new Item("Phần đang bôi đen", "sel") }, null);
            Head(f, "Quét");
            var box = new TableLayoutPanel
            {
                Name = "ReviewOptions", ColumnCount = 2, RowCount = 3, Width = W - 10, Margin = Padding.Empty,
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            box.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            box.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            for (int i = 0; i < 3; i++) box.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            f.Controls.Add(box);
            chRvChem = Chk(f, "Hóa học", "Công thức, ion, phương trình, cân bằng, hóa trị. Phát hiện công thức và phương trình viết sai và đề xuất cách sửa.");
            chRvMath = Chk(f, "Toán", "Biểu thức gõ phẳng như x^2, a_i, mã LaTeX hiện lộ như \\Delta, \\frac{a}{b}, Re_{cyl}^{-0,3}.");
            chRvSym = Chk(f, "Ký hiệu, chỉ số", "Biến theo ngữ cảnh (độ dày t, áp suất P...), chỉ số dưới như P_CO2, R_cyl.");
            chRvNum = Chk(f, "Số và đơn vị", "Dấu phân cách, chữ số thập phân, cách viết đơn vị (rpm, kPa, MPa...), thiếu dấu cách giữa số và đơn vị. Tự bỏ qua số thứ tự đầu dòng, số chú thích bảng hình, năm, số trang, số hiệu, ngày.");
            chRvNumTbl = Chk(f, "Gồm số trong bảng", null);
            {
                CheckBox[] cbs = new CheckBox[] { chRvChem, chRvMath, chRvSym, chRvNum, chRvNumTbl };
                for (int i = 0; i < cbs.Length; i++) { cbs[i].Margin = new Padding(3, 4, 3, 4); box.Controls.Add(cbs[i], i % 2, i / 2); }
            }
            cbRvNumStyle = Combo(f, "Dấu phân cách số khi rà soát", new[]
            {
                new Item("Tự động theo ngôn ngữ quanh số", "auto"),
                new Item("Việt 1.234,56", "vn"),
                new Item("Anh 1,234.56", "us"),
                new Item("SI 1 234,56", "si")
            }, "Tự động nhận tiếng Việt hoặc tiếng Anh gần từng số. Có thể chọn một kiểu cố định cho lần rà soát.");
            chGlyph = Chk(f, "Ký tự nhỏ, thấp", "Công thức cùng dòng với chữ: chọn thì dùng ký tự chỉ số nhỏ, thấp của font Cambria; bỏ chọn thì dùng định dạng chỉ số trên, dưới. Công thức riêng một dòng hoặc có phân số luôn là Equation. Đổi nhanh trên Ribbon bằng nút Kiểu chỉ số.");
            Button[] b0 = BtnRow(f, new string[] { "Quét", "Hiện lại" }, new Action[] { delegate { DoScan(); }, delegate { Run(delegate { return rv.ClearIgnored(); }); } },
                new string[] { "Quét văn bản theo các mục đã chọn", "Hiện lại các mục đã bỏ qua" });
            lblRv = Lbl(f, "");
            lv = new ListView();
            lv.View = View.Details; lv.FullRowSelect = true; lv.MultiSelect = false; lv.HideSelection = false; lv.GridLines = true;
            lv.Width = W - 10; lv.Height = 280; lv.AccessibleName = "Danh sách các mục cần xem";
            lv.Columns.Add("Mức", 96); lv.Columns.Add("Loại", 84); lv.Columns.Add("Hiện tại", 130); lv.Columns.Add("Đề xuất", 130);
            lv.Columns.Add("Sửa", 48); lv.Columns.Add("Bỏ", 52);
            lv.SelectedIndexChanged += delegate { OnPick(); };
            lv.MouseUp += OnListClick;
            lv.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter) { e.Handled = true; e.SuppressKeyPress = true; DoFixOne(); }
                else if (e.KeyCode == Keys.Delete) { e.Handled = true; DoIgnore(); }
            };
            Tip(lv, "Mục đang chọn được tô sáng màu vàng trong văn bản. Bấm dấu ✔ để sửa mục, dấu ✖ để bỏ qua. Bàn phím: Enter để sửa, Delete để bỏ qua.");
            f.Controls.Add(lv);
            tbReason = new TextBox(); tbReason.Multiline = true; tbReason.ReadOnly = true; tbReason.ScrollBars = ScrollBars.Vertical;
            tbReason.Width = W - 10; tbReason.Height = 100; tbReason.AccessibleName = "Lý do của mục đang chọn";
            f.Controls.Add(tbReason);
            cbFix = new ComboBox(); cbFix.DropDownStyle = ComboBoxStyle.DropDownList; cbFix.Width = W - 10; cbFix.AccessibleName = "Cách sửa";
            Tip(cbFix, "Chọn cách sửa khi có nhiều cách. Dùng cho dấu ✔ và nút Sửa tương tự.");
            f.Controls.Add(cbFix);
            Button[] b1 = BtnRow(f, new string[] { "Sửa tương tự", "Bỏ qua tương tự", "Sửa hết" },
                new Action[] { delegate { DoFixSimilar(); }, delegate { DoIgnoreSimilar(); }, delegate { DoFixAll(); } },
                new string[] { "Sửa mọi chỗ giống mục đang chọn (cùng loại, cùng chữ) theo cách đang chọn", "Bỏ qua mọi chỗ giống mục đang chọn", "Sửa tất cả các mục, trừ mục nghi ngờ. Một lần Ctrl+Z hoàn tác một lần sửa." });
            btnFixSim = b1[0]; btnIgnSim = b1[1]; btnFixAll = b1[2];
            btnGoto = BtnRow(f, new string[] { "Tới vị trí" }, new Action[] { delegate { DoGoto(); } }, new string[] { "Chuyển tới vị trí của mục trong văn bản" })[0];
            SetRvEnabled();
        }

        void OnListClick(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo h = lv.HitTest(e.Location);
            if (h.Item == null || h.SubItem == null) return;
            int col = h.Item.SubItems.IndexOf(h.SubItem);
            if (col != 4 && col != 5) return;
            h.Item.Selected = true; h.Item.Focused = true;
            if (col == 4) DoFixOne(); else DoIgnore();
        }

        // ------------------------------------------------------------ thẻ Phương trình
        void BuildCompose(TabPage page)
        {
            FlowLayoutPanel f = Stack(page);
            Lbl(f, "Phương trình");
            ctBox = new Panel(); ctBox.Height = 64; ctBox.Width = W - 20; ctBox.BackColor = Color.White; ctBox.BorderStyle = BorderStyle.FixedSingle;
            tbCt = new TextBox(); tbCt.BorderStyle = BorderStyle.None; tbCt.TextAlign = HorizontalAlignment.Center; tbCt.Font = new Font("Segoe UI", 14f); tbCt.ForeColor = Color.Black; tbCt.BackColor = Color.White;   // Cambria Math có chiều cao dòng rất lớn nên chữ bị cắt, nhìn như mất chữ
            tbCt.AccessibleName = "Phương trình cần soạn";
            Tip(tbCt, "Hóa học: 2H2 + O2 -> 2H2O, SO4^2-, Fe^3+, H_2O. Toán, lý: x^2 + y^2 = r^2, t_{req} = (P D)/(2 S E), \\frac{a}{b}, sqrt(x), \\Delta G. Tự nhận hóa học hay toán.");
            ctBox.Controls.Add(tbCt);
            ctBox.Resize += delegate { CenterInput(); };
            ctBox.Click += delegate { tbCt.Focus(); };
            f.Controls.Add(ctBox);
            CenterInput();
            lblKind = Lbl(f, "");
            lblKind.Font = new Font(Font, FontStyle.Bold);
            cbOpt = new ComboBox(); cbOpt.DropDownStyle = ComboBoxStyle.DropDownList; cbOpt.Width = W - 20; cbOpt.AccessibleName = "Cách hiểu và sửa";
            Tip(cbOpt, "Các cách hiểu hoặc sửa phương trình. Mặc định là cách được đề xuất.");
            f.Controls.Add(cbOpt);
            inlinePreview = AddEquationPreview(f, "Chèn công thức cùng dòng", false, out insertInline);
            displayPreview = AddEquationPreview(f, "Chèn công thức kiểu MathType (dòng riêng, phân số nhiều tầng)", true, out insertDisplay);
            lblCtMsg = Lbl(f, ""); lblCtMsg.Visible = false;
            tbCt.TextChanged += delegate { UpdateCompose(); };
            cbOpt.SelectedIndexChanged += delegate { if (!ctBusy && cur != null) { cur.Select(cbOpt.SelectedIndex); ShowPreview(cur); } };
            UpdateCompose();
        }

        EquationPreview AddEquationPreview(FlowLayoutPanel flow, string label, bool display, out CheckBox toggle)
        {
            var frame = new Panel { Width = W - 20, Height = display ? 164 : 106,
                BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(6) };
            var check = new CheckBox { Text = label, AccessibleName = label, Dock = DockStyle.Top,
                AutoSize = true, Padding = new Padding(2, 4, 2, 6), Enabled = false };
            Tip(check, "Tích để chèn ngay công thức đang xem vào tài liệu Word.");
            var preview = new EquationPreview(display) { Dock = DockStyle.Fill, AccessibleRole = AccessibleRole.StaticText };
            frame.Controls.Add(preview);
            frame.Controls.Add(check);
            frame.Resize += delegate { check.MaximumSize = new Size(Math.Max(1, frame.ClientSize.Width - frame.Padding.Horizontal), 0); };
            check.CheckedChanged += delegate
            {
                if (!check.Checked) return;
                check.Checked = false; // Có thể tích lại để chèn lần sau.
                if (cur == null || !cur.Ok || cur.Fix == null) return;
                ComposeResult formula = cur;
                string error;
                string message = OnWord(delegate { return wc.InsertPreview(formula, display); }, out error);
                Report(error ?? message);
            };
            flow.Controls.Add(frame);
            toggle = check;
            return preview;
        }

        /// <summary>Đặt ô nhập vào giữa khung theo chiều dọc; chữ trong ô đã căn giữa theo chiều ngang.</summary>
        void CenterInput()
        {
            if (tbCt == null || ctBox == null) return;
            tbCt.Width = Math.Max(40, ctBox.ClientSize.Width - 16);
            tbCt.Left = 8; tbCt.Top = Math.Max(0, (ctBox.ClientSize.Height - tbCt.Height) / 2);
        }

        void UpdateCompose()
        {
            ctBusy = true;
            try
            {
                cur = Compose.Build(tbCt.Text);
                cbOpt.Items.Clear();
                foreach (ComposeOption o in cur.Options) cbOpt.Items.Add(Lang.Msg(o.Label + ": " + o.Show));
                if (cur.Ok && cbOpt.Items.Count > 0) cbOpt.SelectedIndex = Math.Max(0, Math.Min(cur.Selected, cbOpt.Items.Count - 1));
                cbOpt.Enabled = cbOpt.Items.Count > 1;
                lblKind.Text = !cur.Ok ? "" : (cur.Kind == "chem" ? "Nhận ra: hóa học" : "Nhận ra: toán, lý");
                lblCtMsg.Text = cur.Ok ? "" : cur.Message;
                lblCtMsg.Visible = !cur.Ok && !string.IsNullOrEmpty(cur.Message);
                ShowPreview(cur);
            }
            finally { ctBusy = false; }
        }

        void ShowPreview(ComposeResult r)
        {
            MNode node = r.Ok && r.Fix != null ? r.Fix.GetNode() : null;
            inlinePreview.Set(node, r.Show);
            displayPreview.Set(node, r.Show);
            insertInline.Enabled = insertDisplay.Enabled = node != null;
        }

        void ClearMarkQuiet()
        {
            if (AllowClose || !rv.HasMark) return;      // đang tắt Word hoặc không có gì để trả lại: không gọi sang luồng chính
            try { string err; OnWord(delegate { rv.ClearMark(); return 0; }, out err); } catch { }
        }

        // ------------------------------------------------------------ đồng bộ cài đặt
        // Bảng điều khiển chỉ giữ các thông số của thẻ Chuẩn hóa và Phương trình. Thông số khác (bảng, số, khóa API...) nằm ở Ribbon
        // và cửa sổ Cài đặt, nên luôn nạp tệp mới nhất rồi chỉ ghi đè phần của mình để không làm mất thay đổi từ nơi khác.
        void LoadToUi()
        {
            loading = true;
            try
            {
                UiUtil.Pick(cbScope, st.RvScope);
                chGlyph.Checked = st.EqMode == "glyph";
                chRvChem.Checked = st.RvChem; chRvMath.Checked = st.RvMath; chRvSym.Checked = st.RvSym; chRvNum.Checked = st.RvNum; chRvNumTbl.Checked = st.RvNumTbl;
                UiUtil.Pick(cbRvNumStyle, st.RvNumStyle);
                if (tabTerm != null) { foreach (ComboBox c in scopeBoxes) UiUtil.Pick(c, st.RvScope); tabTerm.Load(st); tabStyle.Load(st); tabGram.Load(st); }
                if (cbRwTone != null && !string.IsNullOrEmpty(st.AiTone)) UiUtil.Pick(cbRwTone, st.AiTone);
                if (chkRwRewrite != null)
                {
                    chkRwRewrite.Checked = st.RwDoRewrite; chkRwTranslate.Checked = st.RwDoTranslate;
                    UiUtil.Pick(cbRwLang, st.RwTargetLang ?? "en");
                    cbRwLang.Enabled = st.RwDoTranslate; if (cbRwTone != null) cbRwTone.Enabled = st.RwDoRewrite;
                    if (st.RwInH >= 60) tbRwOrig.Height = Math.Max(rwInBase, st.RwInH);
                    if (st.RwOutH >= 60) tbRwOut.Height = Math.Max(rwOutBase, st.RwOutH);
                    UpdateRwRunLabel();
                }
                RefreshAiNotice();
                TopMost = st.TopMost;
            }
            finally { loading = false; }
        }

        void SaveSettings()
        {
            try
            {
                AppSettings fresh = AppSettings.Load();
                fresh.RvScope = UiUtil.Val(cbScope); fresh.EqMode = chGlyph.Checked ? "glyph" : "fmt";
                fresh.RvChem = chRvChem.Checked; fresh.RvMath = chRvMath.Checked; fresh.RvSym = chRvSym.Checked; fresh.RvNum = chRvNum.Checked; fresh.RvNumTbl = chRvNumTbl.Checked;
                fresh.RvNumStyle = UiUtil.Val(cbRvNumStyle);
                if (tabTerm != null) { tabTerm.Save(fresh); tabStyle.Save(fresh); tabGram.Save(fresh); }
                if (cbRwTone != null && cbRwTone.SelectedItem is Item) fresh.AiTone = UiUtil.Val(cbRwTone);
                if (chkRwRewrite != null)
                {
                    fresh.RwDoRewrite = chkRwRewrite.Checked; fresh.RwDoTranslate = chkRwTranslate.Checked;
                    if (cbRwLang.SelectedItem is Item) fresh.RwTargetLang = UiUtil.Val(cbRwLang);
                    fresh.RwInH = tbRwOrig.Height > rwInBase + 4 ? tbRwOrig.Height : 0;
                    fresh.RwOutH = tbRwOut.Height > rwOutBase + 4 ? tbRwOut.Height : 0;
                }
                fresh.Save(); st = fresh; savedAt = FileTime();
            }
            catch { }
        }

        static DateTime FileTime()
        {
            try { return File.GetLastWriteTimeUtc(AppSettings.FilePath); } catch { return DateTime.MinValue; }
        }

        /// <summary>Ribbon hoặc cửa sổ Cài đặt có thể đã đổi cài đặt. Nạp lại nếu tệp mới hơn lần lưu của bảng điều khiển.</summary>
        void ReloadIfChanged()
        {
            DateTime t = FileTime();
            if (t > savedAt) { st = AppSettings.Load(); LoadToUi(); savedAt = t; }
        }

        // ------------------------------------------------------------ chạy lệnh an toàn
        void Report(string msg)
        {
            msg = Lang.Msg(msg);   // Rev0.9.6
            status.Text = msg;
            status.AccessibleName = Lang.T("Thông báo kết quả") + ": " + msg;
            // Rev0.7: không hiện hộp thoại; nội dung nằm ở ô trạng thái và thanh trạng thái của Word.
        }

        T OnWord<T>(Func<T> job, out string error)
        {
            error = null;
            Cursor old = Cursor.Current; Cursor.Current = Cursors.WaitCursor;
            try { return Host.OnWord(job); }
            catch (Exception ex) { error = Host.Explain(ex); Log.W("Lenh loi: " + ex); return default(T); }
            finally { Cursor.Current = old; }
        }

        void Run(Func<string> job)
        {
            SaveSettings();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            string err;
            string msg = OnWord(job, out err);
            sw.Stop();
            Report((err ?? msg) + " (" + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0", CultureInfo.CurrentCulture) + " giây)");
        }

        // ------------------------------------------------------------ chuẩn hóa
        void Fill(RvResult r) { Fill(r, 0); }

        void Fill(RvResult r, int keepAt)
        {
            lv.BeginUpdate();
            lv.Items.Clear();
            Font sym = new Font("Segoe UI Symbol", 12f, FontStyle.Bold);
            foreach (RvRow row in r.Rows)
            {
                var it = new ListViewItem(new string[] { row.SevText, row.CatText, row.Before, row.Show, "✔", "✖" });
                it.UseItemStyleForSubItems = false;
                it.SubItems[4].ForeColor = Color.FromArgb(30, 130, 60); it.SubItems[4].Font = sym;
                it.SubItems[5].ForeColor = Color.FromArgb(190, 45, 40); it.SubItems[5].Font = sym;
                it.Tag = row; lv.Items.Add(it);
            }
            lv.EndUpdate();
            if (lv.Items.Count > 0)
            {
                int at = Math.Max(0, Math.Min(keepAt, lv.Items.Count - 1));
                lv.Items[at].Selected = true; lv.Items[at].Focused = true;
            }
            else { tbReason.Text = ""; cbFix.Items.Clear(); }
            SetRvEnabled();
        }

        void SetRvEnabled()
        {
            bool any = lv.Items.Count > 0;
            btnFixSim.Enabled = any; btnIgnSim.Enabled = any; btnFixAll.Enabled = any; btnGoto.Enabled = any; cbFix.Enabled = any;
        }

        RvRow Current() { return lv.SelectedItems.Count > 0 ? (RvRow)lv.SelectedItems[0].Tag : null; }

        void OnPick()
        {
            RvRow r = Current();
            if (r == null) return;
            tbReason.Text = Lang.Msg(r.SevText + ". " + r.CatText + ". " + r.Reason);
            cbFix.Items.Clear();
            for (int i = 0; i < r.FixLabels.Length; i++)
                cbFix.Items.Add(Lang.Msg(r.FixLabels[i] + ": " + r.Before + " thành " + r.FixShow[i]));
            if (cbFix.Items.Count > 0) cbFix.SelectedIndex = 0;
            string err;
            OnWord(delegate { return rv.Goto(r.Index); }, out err);   // lỗi khi chuyển vị trí không cần báo
        }

        void DoScan()
        {
            SaveSettings();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            string err; RvResult r = OnWord(delegate { return rv.Scan(st); }, out err);
            sw.Stop();
            if (err != null || r == null) { Report(err ?? "Không quét được."); return; }
            Fill(r);
            string msg = r.Message + " (" + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0", CultureInfo.CurrentCulture) + " giây)";
            lblRv.Text = msg; Report(msg);
        }

        void DoGoto()
        {
            RvRow r = Current(); if (r == null) return;
            string err; string m = OnWord(delegate { return rv.Goto(r.Index); }, out err);
            Report(err ?? m);
        }

        void RemoveCurrent()
        {
            if (lv.SelectedItems.Count == 0) return;
            int idx = lv.SelectedItems[0].Index;
            lv.Items.RemoveAt(idx);
            if (idx < lv.Items.Count) { lv.Items[idx].Selected = true; lv.Items[idx].Focused = true; }
            else if (lv.Items.Count > 0) { lv.Items[lv.Items.Count - 1].Selected = true; }
            else { tbReason.Text = Lang.Msg("Đã xử lý hết các mục."); cbFix.Items.Clear(); }
            SetRvEnabled();
            lblRv.Text = "Còn " + lv.Items.Count + " mục.";
        }

        void RefreshPending()
        {
            int at = lv.SelectedItems.Count > 0 ? lv.SelectedItems[0].Index : 0;
            string err2; RvResult p = OnWord(delegate { return rv.Pending(); }, out err2);
            if (p != null) Fill(p, at);
            lblRv.Text = "Còn " + lv.Items.Count + " mục.";
        }

        void DoFixOne()
        {
            RvRow r = Current(); if (r == null) return;
            SaveSettings();
            int fi = cbFix.SelectedIndex < 0 ? 0 : cbFix.SelectedIndex;
            string err; string m = OnWord(delegate { return rv.FixItem(r.Index, fi, st); }, out err);
            if (err != null) { Report(err); return; }
            Report(m); RemoveCurrent();
        }

        void DoIgnore()
        {
            RvRow r = Current(); if (r == null) return;
            string err; string m = OnWord(delegate { return rv.Ignore(r.Index); }, out err);
            if (err != null) { Report(err); return; }
            Report(m); RemoveCurrent();
        }

        void DoFixSimilar()
        {
            RvRow r = Current(); if (r == null) return;
            SaveSettings();
            int fi = cbFix.SelectedIndex < 0 ? 0 : cbFix.SelectedIndex;
            string err; string m = OnWord(delegate { return rv.FixSimilar(r.Index, fi, st); }, out err);
            if (err != null) { Report(err); return; }
            Report(m); RefreshPending();
        }

        void DoIgnoreSimilar()
        {
            RvRow r = Current(); if (r == null) return;
            string err; string m = OnWord(delegate { return rv.IgnoreSimilar(r.Index); }, out err);
            if (err != null) { Report(err); return; }
            Report(m); RefreshPending();
        }

        void DoFixAll()
        {
            SaveSettings();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            string err; string m = OnWord(delegate { return rv.FixAll(st); }, out err);
            sw.Stop();
            if (err != null) { Report(err); return; }
            string err2; RvResult p = OnWord(delegate { return rv.Pending(); }, out err2);
            if (p != null) Fill(p);
            string msg = m + " (" + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0", CultureInfo.CurrentCulture) + " giây)";
            lblRv.Text = msg; Report(msg);
        }

        // ------------------------------------------------------------ thẻ Viết lại (AI)
        void BuildRewrite(Control page)
        {
            FlowLayoutPanel f = Stack(page);
            Head(f, "Công cụ AI: viết lại và dịch");
            Label note = Lbl(f, "Viết lại và dịch thực hiện bằng AI trực tuyến (cài đặt khóa & model trong Cài đặt). Chỉ áp dụng cho đoạn văn được bôi đen không quá 500 từ.");
            note.ForeColor = Color.FromArgb(55, 65, 81);
            note.Font = DpiUtil.CreateUiFont(9.5f, FontStyle.Regular);

            btnRwRead = BtnRow(f, new string[] { "Đọc đoạn đang bôi đen" }, new Action[] { delegate { DoRwRead(false); } },
                new string[] { "Lấy đoạn văn đang chọn trong Word, đếm số từ" })[0];

            lblRwCount = Lbl(f, "Chưa đọc vùng chọn (tối đa 500 từ).");
            tbRwOrig = new TextBox(); tbRwOrig.Multiline = true; tbRwOrig.ReadOnly = true; tbRwOrig.ScrollBars = ScrollBars.Vertical;
            tbRwOrig.Width = W - 10; tbRwOrig.Height = 96; tbRwOrig.AccessibleName = "Đoạn gốc đang bôi đen";
            f.Controls.Add(tbRwOrig);
            rwInBase = tbRwOrig.Height;
            AddGrip(f, tbRwOrig, true);

            FlowLayoutPanel rwRow = new FlowLayoutPanel();
            rwRow.FlowDirection = FlowDirection.LeftToRight; rwRow.WrapContents = true; rwRow.AutoSize = true; rwRow.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            rwRow.Margin = new Padding(3, 8, 3, 2); rwRow.Width = W - 10;
            chkRwRewrite = new CheckBox { Text = "Viết lại", AutoSize = true, Checked = true, Margin = new Padding(3, 6, 14, 3), AccessibleName = "Viết lại" };
            chkRwTranslate = new CheckBox { Text = "Dịch", AutoSize = true, Margin = new Padding(3, 6, 14, 3), AccessibleName = "Dịch" };
            Label lblRwLang = new Label { Text = "Ngôn ngữ đích", AutoSize = true, Margin = new Padding(3, 8, 4, 3) };
            cbRwLang = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Math.Max(150, (W - 10) / 2), Margin = new Padding(3, 3, 3, 3), AccessibleName = "Ngôn ngữ đích" };
            foreach (string[] lg in AiClient.TargetLangs)
                cbRwLang.Items.Add(new Item(lg[1] == lg[2] ? lg[1] : lg[1] + " (" + lg[2] + ")", lg[0]));
            cbRwLang.SelectedIndex = 1;
            Tip(chkRwRewrite, "Viết lại theo văn phong đã chọn. Nếu không chọn Dịch thì giữ nguyên ngôn ngữ gốc.");
            Tip(chkRwTranslate, "Dịch sang ngôn ngữ đích. Nếu không chọn Viết lại thì chỉ dịch, giữ nguyên văn phong và ý.");
            Tip(cbRwLang, "Ngôn ngữ đích khi chọn Dịch.");
            rwRow.Controls.Add(chkRwRewrite); rwRow.Controls.Add(chkRwTranslate); rwRow.Controls.Add(lblRwLang); rwRow.Controls.Add(cbRwLang);
            f.Controls.Add(rwRow);
            EventHandler onRwOpt = delegate { if (loading) return; cbRwLang.Enabled = chkRwTranslate.Checked; if (cbRwTone != null) cbRwTone.Enabled = chkRwRewrite.Checked; RefreshRwRun(); SaveSettings(); };
            chkRwRewrite.CheckedChanged += onRwOpt; chkRwTranslate.CheckedChanged += onRwOpt; cbRwLang.SelectedIndexChanged += onRwOpt;
            cbRwLang.Enabled = false;

            cbRwTone = Combo(f, "Văn phong (Tone)", new Item[]
            {
                new Item("Chuyên nghiệp kỹ thuật", "Chuyên nghiệp kỹ thuật"),
                new Item("Trang trọng", "Trang trọng"),
                new Item("Ngắn gọn súc tích", "Ngắn gọn súc tích"),
                new Item("Tự nhiên mạch lạc", "Tự nhiên mạch lạc"),
                new Item("Học thuật / Báo cáo khoa học", "Học thuật / Báo cáo khoa học")
            }, "Chọn văn phong mong muốn khi viết lại.");

            btnRwRun = BtnRow(f, new string[] { "Viết lại" }, new Action[] { delegate { DoRwRun(); } },
                new string[] { "Gửi phần bôi đen tới AI để viết lại hoặc dịch" })[0];
            btnRwRun.Enabled = false;

            Lbl(f, "Bản xem trước (có thể chỉnh sửa trước khi áp dụng)");
            tbRwOut = new TextBox(); tbRwOut.Multiline = true; tbRwOut.ScrollBars = ScrollBars.Vertical;
            tbRwOut.Width = W - 10; tbRwOut.Height = 140; tbRwOut.AccessibleName = "Kết quả viết lại xem trước";
            f.Controls.Add(tbRwOut);
            rwOutBase = tbRwOut.Height;
            AddGrip(f, tbRwOut, false);

            Button[] acts = BtnRow(f, new string[] { "Áp dụng (giữ định dạng)", "Hoàn tác trong Word" },
                new Action[] { delegate { DoRwApply(); }, delegate { DoRwUndo(); } },
                new string[] { "Áp dụng bản viết lại vào Word vào vùng đã bôi đen (giữ kiểu đoạn và định dạng)", "Hoàn tác bước viết lại vừa áp dụng" });
            btnRwApply = acts[0];
            btnRwUndo = acts[1];
            btnRwApply.Enabled = false;
            btnRwUndo.Enabled = false;
        }

        void UpdateRwState(RewriteSnapshot snap)
        {
            rwSnap = snap;
            rwResultSnap = null;
            btnRwApply.Enabled = false;
            tbRwOut.Clear();
            int wc = snap != null ? snap.WordCount : 0;
            tbRwOrig.Text = snap != null ? snap.Text : "";
            if (wc == 0)
            {
                lblRwCount.ForeColor = Color.FromArgb(185, 28, 28);
                lblRwCount.Text = "Chưa có chữ nào được bôi đen (cần 1–500 từ).";
                btnRwRun.Enabled = false;
            }
            else if (wc > AiClient.MaxRewriteWords)
            {
                lblRwCount.ForeColor = Color.FromArgb(185, 28, 28);
                lblRwCount.Text = "CẢNH BÁO: Đang chọn " + wc + " từ, vượt quá giới hạn " + AiClient.MaxRewriteWords + " từ! Nút Viết lại đã bị khóa.";
                btnRwRun.Enabled = false;
            }
            else
            {
                lblRwCount.ForeColor = Color.FromArgb(22, 101, 52);
                lblRwCount.Text = "Đang chọn: " + wc + " / " + AiClient.MaxRewriteWords + " từ (hợp lệ).";
                btnRwRun.Enabled = !rwBusy && RwModeValid();
            }
            UpdateRwRunLabel();
        }

        bool RwModeValid() { return chkRwRewrite != null && chkRwTranslate != null && (chkRwRewrite.Checked || chkRwTranslate.Checked); }

        void UpdateRwRunLabel()
        {
            if (btnRwRun == null || chkRwRewrite == null) return;
            string t = chkRwRewrite.Checked && chkRwTranslate.Checked ? "Dịch và viết lại" : (chkRwTranslate.Checked ? "Dịch" : (chkRwRewrite.Checked ? "Viết lại" : "Chọn Viết lại hoặc Dịch"));
            btnRwRun.Text = t; btnRwRun.AccessibleName = t;
        }

        void RefreshRwRun()
        {
            if (btnRwRun == null) return;
            int wc = rwSnap != null ? rwSnap.WordCount : 0;
            btnRwRun.Enabled = !rwBusy && wc > 0 && wc <= AiClient.MaxRewriteWords && RwModeValid();
            UpdateRwRunLabel();
        }

        /// <summary>Thanh nắm kéo mỏng dưới ô văn bản: kéo để đổi chiều cao, nhấp đúp để mở rộng gấp ba hoặc thu lại.</summary>
        void AddGrip(FlowLayoutPanel f, TextBox tb, bool isIn)
        {
            Panel grip = new Panel();
            grip.Height = Math.Max(7, (int)Math.Round(7 * DeviceDpi / 96.0));
            grip.Width = tb.Width; grip.Cursor = Cursors.SizeNS; grip.BackColor = Color.FromArgb(205, 210, 218); grip.Margin = new Padding(3, 0, 3, 8);
            Tip(grip, "Kéo để thay đổi chiều cao; nhấp đúp để mở rộng hoặc thu lại");
            int startY = 0, startH = 0; bool drag = false;
            grip.MouseDown += delegate { drag = true; startY = Cursor.Position.Y; startH = tb.Height; grip.Capture = true; };
            grip.MouseMove += delegate
            {
                if (!drag) return;
                int baseH = isIn ? rwInBase : rwOutBase;
                int maxH = (int)Math.Round(600 * DeviceDpi / 96.0);
                tb.Height = Math.Max(baseH, Math.Min(maxH, startH + Cursor.Position.Y - startY));
            };
            grip.MouseUp += delegate { drag = false; grip.Capture = false; if (!loading) SaveSettings(); };
            grip.DoubleClick += delegate
            {
                int baseH = isIn ? rwInBase : rwOutBase;
                tb.Height = tb.Height > baseH + 4 ? baseH : baseH * 3;
                if (!loading) SaveSettings();
            };
            f.Controls.Add(grip);
        }

        void DoRwRead(bool quiet)
        {
            if (rwBusy) return;
            string err;
            RewriteSnapshot snap = OnWord(delegate { return wtx.GetRewriteSnapshot(); }, out err);
            if (err != null)
            {
                UpdateRwState(null);
                if (!quiet) Report(err);
                return;
            }
            UpdateRwState(snap);
            if (!quiet)
                Report(snap.WordCount == 0 ? "Hãy bôi đen đoạn văn cần viết lại trong Word (≤ 500 từ)." : "Đã đọc " + snap.WordCount + " từ trong vùng bôi đen.");
        }

        void RunAi<T>(Func<T> action, Action<T, string> onComplete)
        {
            ThreadPool.QueueUserWorkItem(delegate
            {
                T result = default(T);
                string error = null;
                try { result = action(); }
                catch (UserMsg ex) { error = ex.Message; }
                catch (Exception ex) { Log.W("AI: " + ex.GetType().Name); error = "Không xử lý được yêu cầu AI."; }
                try
                {
                    if (!IsDisposed && IsHandleCreated)
                        BeginInvoke(new MethodInvoker(delegate { if (!IsDisposed) onComplete(result, error); }));
                }
                catch (InvalidOperationException) { }
            });
        }

        void DoRwRun()
        {
            if (rwBusy || !RequireAiKey()) return;
            SaveSettings();
            // Đọc lại vùng chọn ngay lúc bấm để chắc chắn cập nhật số từ mới nhất
            string err;
            RewriteSnapshot snap = OnWord(delegate { return wtx.GetRewriteSnapshot(); }, out err);
            if (err != null) { Report(err); return; }
            UpdateRwState(snap);
            if (snap.WordCount == 0 || snap.WordCount > AiClient.MaxRewriteWords)
            {
                Report(lblRwCount.Text);
                return;
            }
            string tone = cbRwTone.SelectedItem is Item ? UiUtil.Val(cbRwTone) : "Chuyên nghiệp kỹ thuật";
            bool doRw = chkRwRewrite.Checked, doTr = chkRwTranslate.Checked;
            string tl = cbRwLang.SelectedItem is Item ? UiUtil.Val(cbRwLang) : "en";
            if (!doRw && !doTr) { Report("Hãy chọn Viết lại hoặc Dịch (hoặc cả hai)."); return; }
            rwBusy = true;
            btnRwRun.Enabled = false;
            btnRwApply.Enabled = false;
            Report("Đang gọi AI (" + (doRw && doTr ? "dịch và viết lại" : (doTr ? "dịch" : "viết lại")) + ", " + snap.WordCount + " từ" + (doRw ? ", tone: " + tone : "") + ")...");
            AppSettings settings = st;
            RunAi(delegate { return AiClient.TransformText(settings, snap.Text, tone, "", doRw, doTr, tl); },
                delegate(AiChatResult result, string error)
                {
                    rwBusy = false;
                    UpdateRwState(snap);
                    if (error != null || !result.Ok)
                    {
                        Report("Công cụ AI thất bại: " + (error ?? result.Error));
                        return;
                    }
                    rwResultSnap = snap;
                    rwResultTarget = doTr ? tl : null;
                    tbRwOut.Text = result.Text;
                    btnRwApply.Enabled = true;
                    Report("AI đã xử lý xong. Kiểm tra bản xem trước rồi bấm Áp dụng.");
                });
        }

        void DoRwApply()
        {
            if (rwResultSnap == null || rwBusy) { Report("Hãy chạy AI trên vùng chọn trước."); return; }
            SaveSettings();
            string err;
            string m = OnWord(delegate { return wtx.ApplyRewrite(rwResultSnap, tbRwOut.Text, rwResultTarget); }, out err);
            if (err != null) { Report(err); return; }
            rwResultSnap = null;
            btnRwApply.Enabled = false;
            btnRwUndo.Enabled = true;
            Report(m);
        }

        void DoRwUndo()
        {
            string err;
            string m = OnWord(delegate { return wtx.UndoLast(); }, out err);
            if (err != null) { Report(err); return; }
            btnRwUndo.Enabled = false;
            Report(m);
        }
    }
}
