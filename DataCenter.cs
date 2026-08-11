using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Linq;

namespace DataCenter
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    class Record
    {
        public string Number, Country, Type, Channel, Source;
        public DateTime Added;
    }

    class MainForm : Form
    {
        List<Record> data = new List<Record>();
        List<Record> filtered = new List<Record>();
        string dataPath;
        DataGridView grid;
        ComboBox cbCountry, cbType, cbChannel;
        TextBox tbSearch;
        Label lStat, lSummary;

        public MainForm()
        {
            Text = "龙哥数据中心";
            Size = new Size(1200, 700);
            Icon = System.Drawing.SystemIcons.Application;
            dataPath = Path.Combine(Application.StartupPath, "data.tsv");
            LoadData();
            BuildUI();
            ApplyFilter();
        }

        void LoadData()
        {
            if (!File.Exists(dataPath)) return;
            try
            {
                foreach (var line in File.ReadAllLines(dataPath, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var p = line.Split('\t');
                    if (p.Length < 5) continue;
                    data.Add(new Record { Number = p[0], Country = p[1], Type = p[2], Channel = p[3], Source = p[4], Added = DateTime.Now });
                }
            }
            catch { }
        }

        void SaveData()
        {
            try
            {
                var sb = new StringBuilder();
                foreach (var r in data)
                    sb.Append(r.Number).Append('\t').Append(r.Country).Append('\t').Append(r.Type).Append('\t').Append(r.Channel).Append('\t').Append(r.Source).AppendLine();
                File.WriteAllText(dataPath, sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        void BuildUI()
        {
            var top = new Panel { Dock = DockStyle.Top, Height = 56, BackColor = Color.FromArgb(0x22, 0x22, 0x22) };
            Controls.Add(top);
            Btn("导入TXT", 12, 10, 80, 36, (s, e) => OnImport(), top);
            Btn("粘贴", 98, 10, 70, 36, (s, e) => OnPaste(), top);
            Btn("删除", 174, 10, 70, 36, (s, e) => OnDelete(), top);
            Btn("清空", 250, 10, 70, 36, (s, e) => OnClear(), top);
            Btn("导出", 326, 10, 70, 36, (s, e) => OnExport(), top);
            lStat = new Label { Left = 410, Top = 18, ForeColor = Color.LightGray, Font = new Font("Microsoft YaHei", 9), AutoSize = true };
            top.Controls.Add(lStat);

            cbCountry = Cbb(530, 14, 120, top, "全部国家");
            cbCountry.SelectedIndexChanged += (s, e) => ApplyFilter();
            cbType = Cbb(660, 14, 110, top, "全部类型");
            cbType.SelectedIndexChanged += (s, e) => ApplyFilter();
            cbChannel = Cbb(780, 14, 110, top, "全部渠道");
            cbChannel.SelectedIndexChanged += (s, e) => ApplyFilter();
            tbSearch = new TextBox { Left = 900, Top = 14, Width = 140, Font = new Font("Microsoft YaHei", 9) };
            tbSearch.TextChanged += (s, e) => ApplyFilter();
            top.Controls.Add(tbSearch);

            grid = new DataGridView { Dock = DockStyle.Fill, VirtualMode = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersWidth = 40, BackgroundColor = Color.White, BorderStyle = BorderStyle.None };
            grid.CellValueNeeded += (s, e) => { if (e.RowIndex < filtered.Count) { var r = filtered[e.RowIndex]; switch (e.ColumnIndex) { case 0: e.Value = (e.RowIndex + 1).ToString(); break; case 1: e.Value = r.Number; break; case 2: e.Value = r.Country; break; case 3: e.Value = r.Type; break; case 4: e.Value = r.Channel; break; case 5: e.Value = r.Source; break; } } };
            grid.Columns.Add("序号", "序号");
            grid.Columns.Add("号码", "号码");
            grid.Columns.Add("国家", "国家");
            grid.Columns.Add("类型", "类型");
            grid.Columns.Add("渠道", "渠道");
            grid.Columns.Add("来源", "来源");
            grid.Columns[0].FillWeight = 5;
            grid.Columns[1].FillWeight = 20;
            grid.Columns[2].FillWeight = 15;
            grid.Columns[3].FillWeight = 10;
            grid.Columns[4].FillWeight = 15;
            grid.Columns[5].FillWeight = 20;
            Controls.Add(grid);

            var right = new Panel { Dock = DockStyle.Right, Width = 220, BackColor = Color.FromArgb(0xF5, 0xF5, 0xF5) };
            lSummary = new Label { Left = 10, Top = 10, Width = 200, Height = 600, Font = new Font("Microsoft YaHei", 9), ForeColor = Color.FromArgb(0x33, 0x33, 0x33) };
            right.Controls.Add(lSummary);
            Controls.Add(right);
        }

        Button Btn(string t, int x, int y, int w, int h, EventHandler cb, Panel p)
        {
            var b = new Button { Text = t, Left = x, Top = y, Width = w, Height = h, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0x44, 0x44, 0x44), ForeColor = Color.White, Font = new Font("Microsoft YaHei", 9) };
            b.FlatAppearance.BorderSize = 0;
            b.Click += cb;
            p.Controls.Add(b);
            return b;
        }

        ComboBox Cbb(int x, int y, int w, Panel p, string placeholder)
        {
            var cb = new ComboBox { Left = x, Top = y, Width = w, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft YaHei", 9) };
            cb.Items.Add(placeholder);
            cb.SelectedIndex = 0;
            p.Controls.Add(cb);
            return cb;
        }

        string Clean(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
                if (ch >= '0' && ch <= '9') sb.Append(ch);
                else if (ch == '+') sb.Append(ch);
            return sb.ToString().TrimStart('0');
        }

        bool IsDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (char ch in s)
                if (ch < '0' || ch > '9') return false;
            return true;
        }

        string DetectCountry(string num)
        {
            if (num.StartsWith("+1")) return "美国/加拿大";
            if (num.StartsWith("+44")) return "英国";
            if (num.StartsWith("+81")) return "日本";
            if (num.StartsWith("+82")) return "韩国";
            if (num.StartsWith("+91")) return "印度";
            if (num.StartsWith("+92")) return "巴基斯坦";
            if (num.StartsWith("+7")) return "俄罗斯";
            if (num.StartsWith("+61")) return "澳大利亚";
            if (num.StartsWith("+49")) return "德国";
            if (num.StartsWith("+33")) return "法国";
            if (num.StartsWith("+39")) return "意大利";
            if (num.StartsWith("+34")) return "西班牙";
            if (num.StartsWith("+55")) return "巴西";
            if (num.StartsWith("+52")) return "墨西哥";
            if (num.StartsWith("+86")) return "中国";
            if (num.StartsWith("+84")) return "越南";
            if (num.StartsWith("+66")) return "泰国";
            if (num.StartsWith("+62")) return "印尼";
            if (num.StartsWith("+63")) return "菲律宾";
            if (num.StartsWith("+60")) return "马来西亚";
            if (num.StartsWith("+65")) return "新加坡";
            if (num.StartsWith("+852")) return "香港";
            if (num.StartsWith("+886")) return "台湾";
            if (num.StartsWith("+971")) return "阿联酋";
            if (num.StartsWith("+966")) return "沙特";
            if (num.StartsWith("+90")) return "土耳其";
            if (num.StartsWith("+234")) return "尼日利亚";
            if (num.StartsWith("+254")) return "肯尼亚";
            if (num.StartsWith("+27")) return "南非";
            if (num.StartsWith("+20")) return "埃及";
            if (num.StartsWith("+380")) return "乌克兰";
            if (num.StartsWith("+48")) return "波兰";
            if (num.StartsWith("+31")) return "荷兰";
            if (num.StartsWith("+46")) return "瑞典";
            if (num.StartsWith("+41")) return "瑞士";
            if (num.StartsWith("+47")) return "挪威";
            if (num.StartsWith("+45")) return "丹麦";
            if (num.StartsWith("+351")) return "葡萄牙";
            if (num.StartsWith("+353")) return "爱尔兰";
            if (num.StartsWith("+43")) return "奥地利";
            if (num.StartsWith("+32")) return "比利时";
            if (num.StartsWith("+64")) return "新西兰";
            if (num.StartsWith("+54")) return "阿根廷";
            if (num.StartsWith("+56")) return "智利";
            if (num.StartsWith("+57")) return "哥伦比亚";
            if (num.StartsWith("+51")) return "秘鲁";
            if (num.StartsWith("+98")) return "伊朗";
            if (num.StartsWith("+964")) return "伊拉克";
            if (num.StartsWith("+972")) return "以色列";
            if (num.StartsWith("+95")) return "缅甸";
            if (num.StartsWith("+880")) return "孟加拉";
            if (num.StartsWith("+855")) return "柬埔寨";
            if (num.StartsWith("+977")) return "尼泊尔";
            if (num.StartsWith("+94")) return "斯里兰卡";
            if (num.StartsWith("+374")) return "亚美尼亚";
            if (num.StartsWith("+994")) return "阿塞拜疆";
            if (num.StartsWith("+995")) return "格鲁吉亚";
            if (num.StartsWith("+998")) return "乌兹别克";
            if (num.StartsWith("+373")) return "摩尔多瓦";
            if (num.StartsWith("+40")) return "罗马尼亚";
            if (num.StartsWith("+359")) return "保加利亚";
            if (num.StartsWith("+36")) return "匈牙利";
            if (num.StartsWith("+420")) return "捷克";
            if (num.StartsWith("+381")) return "塞尔维亚";
            if (num.StartsWith("+385")) return "克罗地亚";
            if (num.StartsWith("+30")) return "希腊";
            if (num.StartsWith("+352")) return "卢森堡";
            if (num.StartsWith("+356")) return "马耳他";
            if (num.StartsWith("+357")) return "塞浦路斯";
            if (num.StartsWith("+212")) return "摩洛哥";
            if (num.StartsWith("+213")) return "阿尔及利亚";
            if (num.StartsWith("+216")) return "突尼斯";
            if (num.StartsWith("+218")) return "利比亚";
            if (num.StartsWith("+225")) return "科特迪瓦";
            if (num.StartsWith("+233")) return "加纳";
            if (num.StartsWith("+256")) return "乌干达";
            if (num.StartsWith("+255")) return "坦桑尼亚";
            if (num.StartsWith("+251")) return "埃塞俄比亚";
            return "未知";
        }

        string AutoDetectType(string num)
        {
            var raw = Clean(num);
            if (raw.Length == 0) return "未知";
            if (raw.StartsWith("+") || raw.Length > 11) return "国际";
            if (raw.Length == 11 && (raw.StartsWith("1") || raw.StartsWith("13") || raw.StartsWith("14") || raw.StartsWith("15") || raw.StartsWith("16") || raw.StartsWith("17") || raw.StartsWith("18") || raw.StartsWith("19"))) return "手机";
            if (raw.Length >= 10 && raw.Length <= 12) return "固话";
            return "其他";
        }

        void ApplyFilter()
        {
            filtered.Clear();
            var kw = (tbSearch != null ? tbSearch.Text.Trim().ToLower() : "");
            var selCountry = cbCountry != null && cbCountry.SelectedIndex > 0 ? cbCountry.SelectedItem.ToString() : null;
            var selType = cbType != null && cbType.SelectedIndex > 0 ? cbType.SelectedItem.ToString() : null;
            var selChannel = cbChannel != null && cbChannel.SelectedIndex > 0 ? cbChannel.SelectedItem.ToString() : null;
            foreach (var r in data)
            {
                if (selCountry != null && r.Country != selCountry) continue;
                if (selType != null && r.Type != selType) continue;
                if (selChannel != null && r.Channel != selChannel) continue;
                if (kw.Length > 0 && r.Number.ToLower().IndexOf(kw) < 0 && r.Country.ToLower().IndexOf(kw) < 0) continue;
                filtered.Add(r);
            }
            RefreshGrid();
            RefreshSummary();
        }

        void RefreshGrid()
        {
            if (grid == null) return;
            grid.RowCount = filtered.Count;
            grid.Refresh();
            if (lStat != null) lStat.Text = string.Format("共 {0} 条", filtered.Count);
        }

        void RefreshSummary()
        {
            if (lSummary == null) return;
            var byCountry = new Dictionary<string, int>();
            var byType = new Dictionary<string, int>();
            var byChannel = new Dictionary<string, int>();
            foreach (var r in data)
            {
                Inc(byCountry, r.Country);
                Inc(byType, r.Type);
                Inc(byChannel, r.Channel);
            }
            var uCountries = new HashSet<string>();
            var uTypes = new HashSet<string>();
            var uChannels = new HashSet<string>();
            var sb = new StringBuilder();
            sb.AppendLine(string.Format("总计: {0} 条", data.Count));
            sb.AppendLine();
            sb.AppendLine("--- 国家 ---");
            foreach (var kv in byCountry) { sb.AppendLine(string.Format("  {0}: {1}", kv.Key, kv.Value)); uCountries.Add(kv.Key); }
            sb.AppendLine();
            sb.AppendLine("--- 类型 ---");
            foreach (var kv in byType) { sb.AppendLine(string.Format("  {0}: {1}", kv.Key, kv.Value)); uTypes.Add(kv.Key); }
            sb.AppendLine();
            sb.AppendLine("--- 渠道 ---");
            foreach (var kv in byChannel) { sb.AppendLine(string.Format("  {0}: {1}", kv.Key, kv.Value)); uChannels.Add(kv.Key); }
            lSummary.Text = sb.ToString();
            UpdateCombo(cbCountry, uCountries, "全部国家");
            UpdateCombo(cbType, uTypes, "全部类型");
            UpdateCombo(cbChannel, uChannels, "全部渠道");
        }

        void Inc(Dictionary<string, int> d, string k)
        {
            int v;
            d.TryGetValue(k, out v);
            d[k] = v + 1;
        }

        void UpdateCombo(ComboBox cb, HashSet<string> items, string placeholder)
        {
            if (cb == null) return;
            var sel = cb.SelectedIndex > 0 ? cb.SelectedItem.ToString() : "";
            cb.Items.Clear();
            cb.Items.Add(placeholder);
            var sorted = new List<string>(items);
            sorted.Sort();
            foreach (var it in sorted) cb.Items.Add(it);
            if (!string.IsNullOrEmpty(sel))
            {
                int idx = cb.Items.IndexOf(sel);
                cb.SelectedIndex = idx >= 0 ? idx : 0;
            }
        }

        void OnImport()
        {
            var dlg = new OpenFileDialog { Filter = "文本文件|*.txt;*.csv", Multiselect = true };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            foreach (var fpath in dlg.FileNames)
            {
                var source = Path.GetFileNameWithoutExtension(fpath);
                ImportNumbers(ReadFile(fpath), source);
            }
            ApplyFilter();
        }

        void OnPaste()
        {
            try
            {
                var text = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(text)) { MessageBox.Show("剪贴板为空"); return; }
                ImportNumbers(text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList(), "粘贴导入");
                ApplyFilter();
            }
            catch { MessageBox.Show("读取剪贴板失败"); }
        }

        List<string> ReadFile(string path)
        {
            var ext = Path.GetExtension(path).ToLower();
            Encoding enc = Encoding.UTF8;
            try
            {
                var raw = File.ReadAllBytes(path);
                if (raw.Length >= 2 && raw[0] == 0xFF && raw[1] == 0xFE) enc = Encoding.Unicode;
                else if (raw.Length >= 2 && raw[0] == 0xFE && raw[1] == 0xFF) enc = Encoding.BigEndianUnicode;
                else if (raw.Length >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF) enc = Encoding.UTF8;
                else enc = Encoding.Default;
                return new List<string>(File.ReadAllLines(path, enc));
            }
            catch { return new List<string>(); }
        }

        void ImportNumbers(List<string> lines, string source)
        {
            foreach (var line in lines)
            {
                var raw = line.Trim();
                if (string.IsNullOrEmpty(raw)) continue;
                var num = Clean(raw);
                if (num.Length < 5) continue;
                if (!raw.StartsWith("+")) num = "+" + num;
                else num = raw;
                var country = DetectCountry(num);
                var type = AutoDetectType(num);
                var channel = source == "粘贴导入" ? "短信" : "短信";
                data.Add(new Record { Number = num, Country = country, Type = type, Channel = channel, Source = source, Added = DateTime.Now });
            }
        }

        void OnDelete()
        {
            if (grid == null || grid.RowCount == 0) return;
            var sel = grid.SelectedRows;
            if (sel.Count == 0) { MessageBox.Show("请先选择要删除的行"); return; }
            var toRemove = new List<Record>();
            foreach (DataGridViewRow row in sel)
            {
                if (row.Index < filtered.Count)
                    toRemove.Add(filtered[row.Index]);
            }
            foreach (var r in toRemove) data.Remove(r);
            ApplyFilter();
        }

        void OnClear()
        {
            if (data.Count == 0) return;
            if (MessageBox.Show("确定清空全部数据？", "确认", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                data.Clear();
                try { File.Delete(dataPath); } catch { }
                ApplyFilter();
            }
        }

        void OnExport()
        {
            if (filtered.Count == 0) { MessageBox.Show("无数据可导出"); return; }
            var dlg = new SaveFileDialog { Filter = "TSV文件|*.tsv|TXT文件|*.txt", FileName = "export.tsv" };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            try
            {
                var sb = new StringBuilder();
                foreach (var r in filtered)
                    sb.Append(r.Number).Append('\t').Append(r.Country).Append('\t').Append(r.Type).Append('\t').Append(r.Channel).Append('\t').Append(r.Source).AppendLine();
                File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show(string.Format("已导出 {0} 条", filtered.Count));
            }
            catch { MessageBox.Show("导出失败"); }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveData();
            base.OnFormClosing(e);
        }
    }
}
