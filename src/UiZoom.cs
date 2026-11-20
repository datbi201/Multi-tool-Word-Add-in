using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace TroLyKyThuat
{
    internal static class UiZoom
    {
        const int Minimum = 50, Maximum = 130, Step = 10, Default = 100;

        internal static void Attach(Form form, Control content, string key, Action reflow)
        {
            AttachWithFile(form, content, Path.Combine(AppSettings.Dir, "zoom-" + key + ".txt"), reflow);
        }

        internal static void AttachWithFile(Form form, Control content, string path, Action reflow)
        {
            DpiUtil.NoMnemonics(form); form.HandleCreated += delegate { DpiUtil.NoMnemonics(form); }; form.Shown += delegate { DpiUtil.NoMnemonics(form); };
            int percent = Load(path);
            var bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.RightToLeft, WrapContents = false
            };
            var group = new TableLayoutPanel
            {
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 3, RowCount = 1, Margin = new Padding(4)
            };
            for (int i = 0; i < 3; i++) group.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            group.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var plus = ZoomButton("ZoomIn", "+", "Phóng to giao diện");
            var minus = ZoomButton("ZoomOut", "−", "Thu nhỏ giao diện");
            var label = new Label
            {
                Name = "ZoomPercent", AutoSize = false, Size = new Size(65, 32), Anchor = AnchorStyles.None,
                TextAlign = ContentAlignment.MiddleCenter, Margin = Padding.Empty
            };
            group.Controls.Add(minus, 0, 0);
            group.Controls.Add(label, 1, 0);
            group.Controls.Add(plus, 2, 0);
            bar.Controls.Add(group);
            form.Controls.Add(bar);
            content.BringToFront();

            Action update = delegate
            {
                label.Text = percent + "%";
                plus.Enabled = percent < Maximum;
                minus.Enabled = percent > Minimum;
            };
            Action<int> change = delegate(int delta)
            {
                int next = Math.Max(Minimum, Math.Min(Maximum, percent + delta));
                if (next == percent) return;
                DpiUtil.SetZoom(form, next, reflow);
                percent = next;
                update();
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                    File.WriteAllText(path, percent.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
                catch (IOException ex) { Log.W("Zoom lưu: " + ex.Message); }
                catch (UnauthorizedAccessException ex) { Log.W("Zoom lưu: " + ex.Message); }
            };
            plus.Click += delegate { change(Step); };
            minus.Click += delegate { change(-Step); };
            form.Shown += delegate { DpiUtil.SetZoom(form, percent, reflow); };
            update();
        }

        static Button ZoomButton(string name, string text, string description)
        {
            return new Button
            {
                Name = name, Text = text, Size = new Size(32, 32), AccessibleName = description,
                TextAlign = ContentAlignment.MiddleCenter, Margin = Padding.Empty, Padding = Padding.Empty,
                Anchor = AnchorStyles.None, UseCompatibleTextRendering = false
            };
        }

        static int Load(string path)
        {
            try
            {
                int percent;
                if (int.TryParse(File.ReadAllText(path), out percent) &&
                    percent >= Minimum && percent <= Maximum && percent % Step == 0) return percent;
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            return Default;
        }
    }
}
