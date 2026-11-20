// Trợ lý Kỹ thuật - Rev0.9.6: đổi ngôn ngữ giao diện Việt, Anh (Technical Assistant: Vietnamese / English interface switch).
// Mọi chữ trong mã nguồn vẫn viết bằng tiếng Việt. Lớp này dịch ở "biên giao diện": Ribbon (nhãn, chú thích), các cửa sổ WinForms và thông báo kết quả.
// Chuỗi chưa có bản dịch được giữ nguyên tiếng Việt (không bao giờ gây lỗi). Dữ liệu dịch nằm trong LangData.cs.
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace TroLyKyThuat
{
    public static class Lang
    {
        public const string Vi = "vi", En = "en";
        static string current = Vi;
        static readonly object gate = new object();
        static Dictionary<string, string> exact;
        static Dictionary<char, List<KeyValuePair<string, string>>> frags;

        /// <summary>Phát ra khi ngôn ngữ thật sự đổi (Ribbon làm mới, bảng điều khiển dịch lại).</summary>
        public static event Action Changed;

        public static string Current { get { return current; } }
        public static bool IsEn { get { return current == En; } }

        public static string Normalize(string code) { return code == En ? En : Vi; }

        /// <summary>Đặt ngôn ngữ. Trả về true nếu có đổi.</summary>
        public static bool Set(string code)
        {
            string n = Normalize(code);
            if (n == current) return false;
            current = n;
            Action h = Changed;
            if (h != null) { try { h(); } catch (Exception ex) { Log.W("Lang.Changed loi: " + ex.Message); } }
            return true;
        }

        static void Init()
        {
            if (exact != null) return;
            lock (gate)
            {
                if (exact != null) return;
                var ex = new Dictionary<string, string>(StringComparer.Ordinal);
                string[] e = LangData.Exact;
                for (int i = 0; i + 1 < e.Length; i += 2) ex[e[i]] = e[i + 1];
                var fr = new Dictionary<char, List<KeyValuePair<string, string>>>();
                string[] f = LangData.Frag;
                for (int i = 0; i + 1 < f.Length; i += 2)
                {
                    if (string.IsNullOrEmpty(f[i])) continue;
                    List<KeyValuePair<string, string>> l;
                    if (!fr.TryGetValue(f[i][0], out l)) { l = new List<KeyValuePair<string, string>>(); fr[f[i][0]] = l; }
                    l.Add(new KeyValuePair<string, string>(f[i], f[i + 1]));
                }
                foreach (List<KeyValuePair<string, string>> l in fr.Values) l.Sort(delegate(KeyValuePair<string, string> a, KeyValuePair<string, string> b) { return b.Key.Length.CompareTo(a.Key.Length); });
                frags = fr; exact = ex;
            }
        }

        /// <summary>Tra khớp nguyên chuỗi (nhãn, chú thích, tiêu đề). Ngôn ngữ Việt hoặc chuỗi lạ thì trả về nguyên văn.</summary>
        public static string T(string vi)
        {
            if (current != En || string.IsNullOrEmpty(vi)) return vi;
            Init();
            string en;
            return exact.TryGetValue(vi, out en) ? en : vi;
        }

        /// <summary>Thông báo kết quả có số xen giữa: khớp nguyên chuỗi, nếu không thì dịch theo mảnh (mảnh dài trước), giữ số và tên riêng ở nguyên vị trí.</summary>
        public static string Msg(string vi)
        {
            if (current != En || string.IsNullOrEmpty(vi)) return vi;
            Init();
            string en;
            if (exact.TryGetValue(vi, out en)) return en;
            var sb = new StringBuilder(vi.Length + 16);
            int i = 0;
            while (i < vi.Length)
            {
                List<KeyValuePair<string, string>> l;
                bool hit = false;
                if (frags.TryGetValue(vi[i], out l))
                {
                    for (int k = 0; k < l.Count; k++)
                    {
                        string key = l[k].Key;
                        if (string.CompareOrdinal(vi, i, key, 0, key.Length) == 0 && WordEdgesOk(vi, i, key))
                        {
                            sb.Append(l[k].Value); i += key.Length; hit = true; break;
                        }
                    }
                }
                if (!hit) { sb.Append(vi[i]); i++; }
            }
            return sb.ToString();
        }

        /// <summary>Mảnh bắt đầu hoặc kết thúc bằng chữ cái thì phải nằm trọn từ: không khớp giữa một từ khác (ví dụ "ô" trong "kilôpascan").</summary>
        static bool WordEdgesOk(string text, int pos, string key)
        {
            if (char.IsLetter(key[0]) && pos > 0 && char.IsLetter(text[pos - 1])) return false;
            int end = pos + key.Length;
            if (char.IsLetter(key[key.Length - 1]) && end < text.Length && char.IsLetter(text[end])) return false;
            return true;
        }

        /// <summary>Văn bản nhiều dòng (hướng dẫn): dịch từng dòng, dòng lạ giữ nguyên.</summary>
        public static string Doc(string vi)
        {
            if (current != En || string.IsNullOrEmpty(vi)) return vi;
            string[] lines = vi.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++) lines[i] = Msg(lines[i]);
            return string.Join("\r\n", lines);
        }

        /// <summary>Chuỗi có chữ cái tiếng Việt có dấu (dùng cho kiểm thử độ phủ).</summary>
        public static bool HasVietnamese(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            const string marks = "àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ";
            foreach (char c in s) if (marks.IndexOf(char.ToLowerInvariant(c)) >= 0) return true;
            return false;
        }

        // ------------------------------------------------------------------------------------------ Ribbon
        /// <summary>
        /// Đổi nhãn, tên gọi, chú thích tĩnh trong Ribbon XML thành lệnh gọi lại (getLabel, getScreentip, getSupertip) để đổi ngôn ngữ ngay,
        /// không cần mở lại Word. Nội dung gốc (tiếng Việt) được lưu vào bảng, khóa "mã|thuộc tính".
        /// </summary>
        public static string TransformRibbon(string xml, IDictionary<string, string> store)
        {
            store.Clear();
            var doc = new XmlDocument();
            doc.LoadXml(xml);
            string[][] attrs = { new[] { "label", "getLabel", "GetLbl" }, new[] { "screentip", "getScreentip", "GetScr" }, new[] { "supertip", "getSupertip", "GetSup" } };
            foreach (XmlElement e in doc.SelectNodes("//*[@id]"))
            {
                if (e.LocalName == "item") continue;
                string id = e.GetAttribute("id");
                foreach (string[] a in attrs)
                {
                    if (!e.HasAttribute(a[0]) || e.HasAttribute(a[1])) continue;
                    string vi = e.GetAttribute(a[0]);
                    e.RemoveAttribute(a[0]);
                    e.SetAttribute(a[1], a[2]);
                    store[id + "|" + a[0]] = vi;
                }
            }
            return doc.OuterXml;
        }

        // ------------------------------------------------------------------------------------------ WinForms
        sealed class Memo { public string Orig, Applied; }
        static readonly ConditionalWeakTable<object, Dictionary<string, Memo>> memos = new ConditionalWeakTable<object, Dictionary<string, Memo>>();

        /// <summary>Dịch (hoặc trả về tiếng Việt) một thuộc tính chữ của đối tượng, nhớ bản gốc để đổi qua lại nhiều lần.</summary>
        static string Resolve(object owner, string prop, string cur, Func<string, string> tr)
        {
            if (cur == null) return null;
            Dictionary<string, Memo> d = memos.GetOrCreateValue(owner);
            Memo m;
            string orig = cur;
            if (d.TryGetValue(prop, out m) && m.Applied == cur) orig = m.Orig;
            string res = tr(orig);
            if (m == null) { m = new Memo(); d[prop] = m; }
            m.Orig = orig; m.Applied = res;
            return res;
        }

        /// <summary>
        /// Dịch toàn bộ cây điều khiển: chữ của nhãn, nút, ô tích, tiêu đề, thẻ, cột; chú thích (ToolTip); tên truy cập; mục ComboBox.
        /// Không đụng tới ô nhập của người dùng. Gọi lại khi đổi ngôn ngữ: dịch lại từ bản gốc tiếng Việt.
        /// </summary>
        public static void Apply(Control root, params ToolTip[] tips)
        {
            if (root == null) return;
            ApplyOne(root, tips);
        }

        static void ApplyOne(Control c, ToolTip[] tips)
        {
            try
            {
                if (c is Label || c is ButtonBase || c is GroupBox || c is TabPage || c is Form)
                {
                    string s = Resolve(c, "Text", c.Text, Msg);
                    if (s != c.Text) c.Text = s;
                }
                string an = Resolve(c, "AName", c.AccessibleName, Msg);
                if (an != c.AccessibleName) c.AccessibleName = an;
                string ad = Resolve(c, "ADesc", c.AccessibleDescription, Msg);
                if (ad != c.AccessibleDescription) c.AccessibleDescription = ad;
                if (tips != null)
                    for (int ti = 0; ti < tips.Length; ti++)
                    {
                        string t = tips[ti].GetToolTip(c);
                        if (!string.IsNullOrEmpty(t)) { string r = Resolve(c, "Tip" + ti, t, Msg); if (r != t) tips[ti].SetToolTip(c, r); }
                    }
                TabPage tp = c as TabPage;
                if (tp != null && !string.IsNullOrEmpty(tp.ToolTipText))
                {
                    string r = Resolve(tp, "TTT", tp.ToolTipText, Msg); if (r != tp.ToolTipText) tp.ToolTipText = r;
                }
                TextBox tb = c as TextBox;
                if (tb != null && tb.ReadOnly && tb.Multiline && (tb.Tag as string) == "doc")
                {
                    string r = Resolve(tb, "Doc", tb.Text, Doc); if (r != tb.Text) { tb.Text = r; tb.Select(0, 0); }
                }
                ComboBox cb = c as ComboBox;
                if (cb != null && cb.Items.Count > 0 && !(cb.Items[0] is string)) RefreshCombo(cb);   // mục là đối tượng (Item) tự dịch khi hiện
                ListView lv = c as ListView;
                if (lv != null)
                {
                    foreach (ColumnHeader h in lv.Columns) { string r = Resolve(h, "Text", h.Text, Msg); if (r != h.Text) h.Text = r; }
                    // hai cột đầu của danh sách rà soát (Mức, Loại) là nhãn; các cột sau là nội dung văn bản của người dùng nên không dịch
                    foreach (ListViewItem it in lv.Items)
                        for (int si = 0; si < it.SubItems.Count && si < 2; si++)
                        {
                            ListViewItem.ListViewSubItem sub = it.SubItems[si];
                            string r = Resolve(sub, "Text", sub.Text, Msg); if (r != sub.Text) sub.Text = r;
                        }
                }
            }
            catch (Exception ex) { Log.W("Lang.Apply loi: " + ex.Message); }
            foreach (Control child in c.Controls) ApplyOne(child, tips);
        }

        static readonly MethodInfo refreshItems = typeof(ComboBox).GetMethod("RefreshItems", BindingFlags.Instance | BindingFlags.NonPublic);
        static void RefreshCombo(ComboBox cb)
        {
            try { if (refreshItems != null) refreshItems.Invoke(cb, null); cb.Invalidate(); } catch { }
        }
    }
}
