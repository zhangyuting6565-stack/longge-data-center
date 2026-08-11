using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
        TreeView tree;
        TextBox txtSearch;
        ComboBox cbCountry;
        Button btnPaste, btnClear, btnExport, btnDelete;
        Label lblStatus;

        public MainForm()
        {
            Text = "龙哥数据中心";
            Size = new Size(1200, 700);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei", 9f);
            dataPath = Path.Combine(Application.StartupPath, "data.tsv");
            LoadData();
            BuildUI();
            RefreshAll();
        }

        void LoadData()
        {
            if (!File.Exists(dataPath)) return;
            try
            {
                foreach (var line in File.ReadAllLines(dataPath, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var parts = line.Split('\t');
                    if (parts.Length < 4) continue;
                    DateTime dt;
                    if (!DateTime.TryParse(parts.Length > 5 ? parts[5] : "", out dt))
                        dt = DateTime.Now;
                    data.Add(new Record
                    {
                        Number = Clean(parts[0]),
                        Country = parts.Length > 1 ? parts[1] : "",
                        Type = parts.Length > 2 ? parts[2] : "",
                        Channel = parts.Length > 3 ? parts[3] : "短信",
                        Source = parts.Length > 4 ? parts[4] : "",
                        Added = dt
                    });
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
                    sb.AppendFormat("{0}\t{1}\t{2}\t{3}\t{4}\t{5:yyyy-MM-dd}\n",
                        r.Number, r.Country, r.Type, r.Channel, r.Source, r.Added);
                File.WriteAllText(dataPath, sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        void BuildUI()
        {
            var top = new Panel { Dock = DockStyle.Top, Height = 46, BackColor = Color.FromArgb(0x22, 0x22, 0x22) };
            Controls.Add(top);

            var btnImport = new Button { Text = "导入文件", Left = 12, Top = 10, Width = 80, Height = 28, FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(0x33, 0x33, 0x33), FlatAppearance = { BorderSize = 1 } };
            btnImport.Click += OnImport;
            top.Controls.Add(btnImport);

            btnPaste = new Button { Text = "粘贴导入", Left = 100, Top = 10, Width = 80, Height = 28, FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(0x33, 0x33, 0x33), FlatAppearance = { BorderSize = 1 } };
            btnPaste.Click += OnPaste;
            top.Controls.Add(btnPaste);

            btnClear = new Button { Text = "清空数据", Left = 188, Top = 10, Width = 80, Height = 28, FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(0x33, 0x33, 0x33), FlatAppearance = { BorderSize = 1 } };
            btnClear.Click += OnClear;
            top.Controls.Add(btnClear);

            txtSearch = new TextBox { Left = 290, Top = 12, Width = 200 };
            txtSearch.TextChanged += (s, e) => ApplyFilter();
            top.Controls.Add(txtSearch);
            top.Controls.Add(new Label { Text = "搜索:", Left = 278, Top = 16, ForeColor = Color.White, AutoSize = true });

            cbCountry = new ComboBox { Left = 500, Top = 12, Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            cbCountry.SelectedIndexChanged += (s, e) => ApplyFilter();
            top.Controls.Add(cbCountry);
            top.Controls.Add(new Label { Text = "国家:", Left = 500, Top = 16, ForeColor = Color.White, AutoSize = true });

            btnExport = new Button { Text = "导出视图", Left = 640, Top = 10, Width = 80, Height = 28, FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(0x33, 0x33, 0x33), FlatAppearance = { BorderSize = 1 } };
            btnExport.Click += OnExport;
            top.Controls.Add(btnExport);

            var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 720, Panel1MinSize = 400 };
            Controls.Add(split);

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None
            };
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Number", HeaderText = "号码", MinimumWidth = 120 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Country", HeaderText = "国家", MinimumWidth = 80 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", HeaderText = "类型", MinimumWidth = 60 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Channel", HeaderText = "渠道", MinimumWidth = 60 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Source", HeaderText = "来源", MinimumWidth = 100 });
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = true;
            split.Panel1.Controls.Add(grid);

            var rightPanel = new Panel { Dock = DockStyle.Fill };
            var lblTitle = new Label { Text = "国家汇总", Dock = DockStyle.Top, Height = 30, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft YaHei", 11f, FontStyle.Bold) };
            rightPanel.Controls.Add(lblTitle);

            tree = new TreeView { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, Font = new Font("Microsoft YaHei", 9f) };
            tree.AfterSelect += (s, e) =>
            {
                if (e.Node != null && e.Node.Level == 0)
                {
                    cbCountry.SelectedItem = e.Node.Text.Split(new[] { ' ' })[0];
                }
            };
            rightPanel.Controls.Add(tree);
            split.Panel2.Controls.Add(rightPanel);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 34, BackColor = Color.FromArgb(0xF0, 0xF0, 0xF0) };

            lblStatus = new Label { Left = 12, Top = 8, AutoSize = true, Text = "共 0 条" };
            bottom.Controls.Add(lblStatus);

            btnDelete = new Button { Text = "删除已选", Left = 200, Top = 5, Width = 80, Height = 24, FlatStyle = FlatStyle.Flat, ForeColor = Color.FromArgb(0xCC, 0x33, 0x33), BackColor = Color.White };
            btnDelete.Click += OnDelete;
            bottom.Controls.Add(btnDelete);

            Controls.Add(bottom);
        }

        string Clean(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var arr = s.Trim().ToCharArray();
            var result = new char[arr.Length];
            int j = 0;
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] >= '0' && arr[i] <= '9') { result[j++] = arr[i]; continue; }
            }
            if (j == 0) return "";
            var num = new string(result, 0, j);
            // deduplicate repeating patterns (12345-12345 → 12345)
            int len = num.Length;
            if (len >= 4 && len % 2 == 0)
            {
                int half = len / 2;
                if (num.Substring(0, half) == num.Substring(half, half))
                    num = num.Substring(0, half);
            }
            return num;
        }

        bool IsDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            for (int i = 0; i < s.Length; i++)
                if (s[i] < '0' || s[i] > '9') return false;
            return true;
        }

        string DetectCountry(string n)
        {
            if (string.IsNullOrEmpty(n)) return "未知";
            if (n.StartsWith("852")) return "香港";
            if (n.StartsWith("853")) return "澳门";
            if (n.StartsWith("886")) return "台湾";
            if (n.StartsWith("86")) return "中国";
            if (n.StartsWith("7")) return "俄罗斯";
            if (n.StartsWith("1")) return "美国/加拿大";
            if (n.StartsWith("81")) return "日本";
            if (n.StartsWith("82")) return "韩国";
            if (n.StartsWith("84")) return "越南";
            if (n.StartsWith("65")) return "新加坡";
            if (n.StartsWith("60")) return "马来西亚";
            if (n.StartsWith("66")) return "泰国";
            if (n.StartsWith("62")) return "印尼";
            if (n.StartsWith("63")) return "菲律宾";
            if (n.StartsWith("91")) return "印度";
            if (n.StartsWith("92")) return "巴基斯坦";
            if (n.StartsWith("90")) return "土耳其";
            if (n.StartsWith("98")) return "伊朗";
            if (n.StartsWith("95")) return "缅甸";
            if (n.StartsWith("94")) return "斯里兰卡";
            if (n.StartsWith("93")) return "阿富汗";
            if (n.StartsWith("880")) return "孟加拉";
            if (n.StartsWith("855")) return "柬埔寨";
            if (n.StartsWith("856")) return "老挝";
            if (n.StartsWith("20")) return "埃及";
            if (n.StartsWith("27")) return "南非";
            if (n.StartsWith("234")) return "尼日利亚";
            if (n.StartsWith("254")) return "肯尼亚";
            if (n.StartsWith("30")) return "希腊";
            if (n.StartsWith("31")) return "荷兰";
            if (n.StartsWith("32")) return "比利时";
            if (n.StartsWith("33")) return "法国";
            if (n.StartsWith("34")) return "西班牙";
            if (n.StartsWith("36")) return "匈牙利";
            if (n.StartsWith("39")) return "意大利";
            if (n.StartsWith("40")) return "罗马尼亚";
            if (n.StartsWith("41")) return "瑞士";
            if (n.StartsWith("43")) return "奥地利";
            if (n.StartsWith("44")) return "英国";
            if (n.StartsWith("45")) return "丹麦";
            if (n.StartsWith("46")) return "瑞典";
            if (n.StartsWith("47")) return "挪威";
            if (n.StartsWith("48")) return "波兰";
            if (n.StartsWith("49")) return "德国";
            if (n.StartsWith("51")) return "秘鲁";
            if (n.StartsWith("52")) return "墨西哥";
            if (n.StartsWith("53")) return "古巴";
            if (n.StartsWith("54")) return "阿根廷";
            if (n.StartsWith("55")) return "巴西";
            if (n.StartsWith("56")) return "智利";
            if (n.StartsWith("57")) return "哥伦比亚";
            if (n.StartsWith("58")) return "委内瑞拉";
            if (n.StartsWith("61")) return "澳大利亚";
            if (n.StartsWith("64")) return "新西兰";
            if (n.StartsWith("355")) return "阿尔巴尼亚";
            if (n.StartsWith("213")) return "阿尔及利亚";
            if (n.StartsWith("376")) return "安道尔";
            if (n.StartsWith("244")) return "安哥拉";
            if (n.StartsWith("374")) return "亚美尼亚";
            if (n.StartsWith("297")) return "阿鲁巴";
            if (n.StartsWith("994")) return "阿塞拜疆";
            if (n.StartsWith("973")) return "巴林";
            if (n.StartsWith("375")) return "白俄罗斯";
            if (n.StartsWith("501")) return "伯利兹";
            if (n.StartsWith("229")) return "贝宁";
            if (n.StartsWith("975")) return "不丹";
            if (n.StartsWith("591")) return "玻利维亚";
            if (n.StartsWith("387")) return "波黑";
            if (n.StartsWith("267")) return "博茨瓦纳";
            if (n.StartsWith("673")) return "文莱";
            if (n.StartsWith("359")) return "保加利亚";
            if (n.StartsWith("226")) return "布基纳法索";
            if (n.StartsWith("257")) return "布隆迪";
            if (n.StartsWith("855")) return "柬埔寨";
            if (n.StartsWith("237")) return "喀麦隆";
            if (n.StartsWith("235")) return "乍得";
            if (n.StartsWith("269")) return "科摩罗";
            if (n.StartsWith("242")) return "刚果";
            if (n.StartsWith("506")) return "哥斯达黎加";
            if (n.StartsWith("225")) return "科特迪瓦";
            if (n.StartsWith("385")) return "克罗地亚";
            if (n.StartsWith("357")) return "塞浦路斯";
            if (n.StartsWith("420")) return "捷克";
            if (n.StartsWith("253")) return "吉布提";
            if (n.StartsWith("593")) return "厄瓜多尔";
            if (n.StartsWith("503")) return "萨尔瓦多";
            if (n.StartsWith("240")) return "赤道几内亚";
            if (n.StartsWith("291")) return "厄立特里亚";
            if (n.StartsWith("372")) return "爱沙尼亚";
            if (n.StartsWith("251")) return "埃塞俄比亚";
            if (n.StartsWith("679")) return "斐济";
            if (n.StartsWith("358")) return "芬兰";
            if (n.StartsWith("241")) return "加蓬";
            if (n.StartsWith("220")) return "冈比亚";
            if (n.StartsWith("995")) return "格鲁吉亚";
            if (n.StartsWith("233")) return "加纳";
            if (n.StartsWith("502")) return "危地马拉";
            if (n.StartsWith("224")) return "几内亚";
            if (n.StartsWith("592")) return "圭亚那";
            if (n.StartsWith("509")) return "海地";
            if (n.StartsWith("504")) return "洪都拉斯";
            if (n.StartsWith("354")) return "冰岛";
            if (n.StartsWith("964")) return "伊拉克";
            if (n.StartsWith("353")) return "爱尔兰";
            if (n.StartsWith("972")) return "以色列";
            if (n.StartsWith("962")) return "约旦";
            if (n.StartsWith("7")) return "哈萨克斯坦";
            if (n.StartsWith("686")) return "基里巴斯";
            if (n.StartsWith("965")) return "科威特";
            if (n.StartsWith("996")) return "吉尔吉斯斯坦";
            if (n.StartsWith("371")) return "拉脱维亚";
            if (n.StartsWith("961")) return "黎巴嫩";
            if (n.StartsWith("266")) return "莱索托";
            if (n.StartsWith("231")) return "利比里亚";
            if (n.StartsWith("218")) return "利比亚";
            if (n.StartsWith("423")) return "列支敦士登";
            if (n.StartsWith("370")) return "立陶宛";
            if (n.StartsWith("352")) return "卢森堡";
            if (n.StartsWith("389")) return "马其顿";
            if (n.StartsWith("261")) return "马达加斯加";
            if (n.StartsWith("265")) return "马拉维";
            if (n.StartsWith("960")) return "马尔代夫";
            if (n.StartsWith("223")) return "马里";
            if (n.StartsWith("356")) return "马耳他";
            if (n.StartsWith("222")) return "毛里塔尼亚";
            if (n.StartsWith("230")) return "毛里求斯";
            if (n.StartsWith("373")) return "摩尔多瓦";
            if (n.StartsWith("377")) return "摩纳哥";
            if (n.StartsWith("976")) return "蒙古";
            if (n.StartsWith("382")) return "黑山";
            if (n.StartsWith("212")) return "摩洛哥";
            if (n.StartsWith("258")) return "莫桑比克";
            if (n.StartsWith("264")) return "纳米比亚";
            if (n.StartsWith("977")) return "尼泊尔";
            if (n.StartsWith("505")) return "尼加拉瓜";
            if (n.StartsWith("227")) return "尼日尔";
            if (n.StartsWith("968")) return "阿曼";
            if (n.StartsWith("507")) return "巴拿马";
            if (n.StartsWith("675")) return "巴布亚新几内亚";
            if (n.StartsWith("595")) return "巴拉圭";
            if (n.StartsWith("351")) return "葡萄牙";
            if (n.StartsWith("974")) return "卡塔尔";
            if (n.StartsWith("7")) return "俄罗斯";
            if (n.StartsWith("250")) return "卢旺达";
            if (n.StartsWith("966")) return "沙特阿拉伯";
            if (n.StartsWith("221")) return "塞内加尔";
            if (n.StartsWith("381")) return "塞尔维亚";
            if (n.StartsWith("232")) return "塞拉利昂";
            if (n.StartsWith("421")) return "斯洛伐克";
            if (n.StartsWith("386")) return "斯洛文尼亚";
            if (n.StartsWith("252")) return "索马里";
            if (n.StartsWith("211")) return "南苏丹";
            if (n.StartsWith("249")) return "苏丹";
            if (n.StartsWith("597")) return "苏里南";
            if (n.StartsWith("268")) return "斯威士兰";
            if (n.StartsWith("963")) return "叙利亚";
            if (n.StartsWith("992")) return "塔吉克斯坦";
            if (n.StartsWith("255")) return "坦桑尼亚";
            if (n.StartsWith("228")) return "多哥";
            if (n.StartsWith("216")) return "突尼斯";
            if (n.StartsWith("256")) return "乌干达";
            if (n.StartsWith("380")) return "乌克兰";
            if (n.StartsWith("971")) return "阿联酋";
            if (n.StartsWith("598")) return "乌拉圭";
            if (n.StartsWith("998")) return "乌兹别克斯坦";
            if (n.StartsWith("678")) return "瓦努阿图";
            if (n.StartsWith("967")) return "也门";
            if (n.StartsWith("260")) return "赞比亚";
            if (n.StartsWith("263")) return "津巴布韦";
            return "未知";
        }

        string AutoDetectType(string n)
        {
            if (string.IsNullOrEmpty(n)) return "未知";
            if (!n.StartsWith("86") && n.Length > 2) return "国际";
            if (n.StartsWith("86")) n = n.Substring(2);
            if (n.Length >= 11 && n[0] == '1' && n[1] >= '3' && n[1] <= '9') return "手机";
            if (n.Length >= 10 && n[0] == '0' && n[1] >= '1') return "固话";
            return "未知";
        }

        void RefreshAll()
        {
            RefreshGrid();
            RefreshTree();
            UpdateStatus();
            PopulateCountryCombo();
        }

        void ApplyFilter()
        {
            var search = (txtSearch.Text ?? "").Trim().ToLower();
            var country = cbCountry.SelectedItem != null ? (string)cbCountry.SelectedItem : "";

            filtered = data.Where(r =>
            {
                if (!string.IsNullOrEmpty(country) && r.Country != country) return false;
                if (!string.IsNullOrEmpty(search))
                {
                    if (r.Number.ToLower().Contains(search)) return true;
                    if (r.Country.ToLower().Contains(search)) return true;
                    if (r.Type.ToLower().Contains(search)) return true;
                    if (r.Channel.ToLower().Contains(search)) return true;
                    return false;
                }
                return true;
            }).ToList();

            RefreshGrid();
            RefreshTree();
            UpdateStatus();
        }

        void RefreshGrid()
        {
            grid.Rows.Clear();
            foreach (var r in filtered)
                grid.Rows.Add(r.Number, r.Country, r.Type, r.Channel, r.Source);
        }

        void RefreshTree()
        {
            tree.Nodes.Clear();
            var groups = filtered
                .GroupBy(r => r.Country)
                .OrderBy(g => g.Count() == 0 ? 0 : -g.Count());
            foreach (var cg in groups)
            {
                var countryNode = new TreeNode(string.Format("{0} ({1})", cg.Key, cg.Count()));
                var types = cg.GroupBy(r => r.Type).OrderBy(g => -g.Count());
                foreach (var tg in types)
                {
                    var typeNode = new TreeNode(string.Format("{0} ({1})", tg.Key, tg.Count()));
                    var channels = tg.GroupBy(r => r.Channel).OrderBy(g => -g.Count());
                    foreach (var chg in channels)
                        typeNode.Nodes.Add(string.Format("{0} ({1})", chg.Key, chg.Count()));
                    countryNode.Nodes.Add(typeNode);
                }
                tree.Nodes.Add(countryNode);
            }
            if (tree.Nodes.Count > 0)
            {
                tree.Nodes[0].Expand();
                if (tree.Nodes[0].Nodes.Count > 0)
                    tree.Nodes[0].Nodes[0].Expand();
            }
        }

        void UpdateStatus()
        {
            lblStatus.Text = string.Format("共 {0} 条  |  显示 {1} 条  |  已选 {2} 条", data.Count, filtered.Count, grid.SelectedRows.Count);
        }

        void PopulateCountryCombo()
        {
            var current = cbCountry.SelectedItem != null ? (string)cbCountry.SelectedItem : "";
            cbCountry.Items.Clear();
            cbCountry.Items.Add("");
            foreach (var c in data.Select(r => r.Country).Distinct().OrderBy(s => s))
                cbCountry.Items.Add(c);
            if (cbCountry.Items.Contains(current))
                cbCountry.SelectedItem = current;
            else
                cbCountry.SelectedIndex = 0;
        }

        void OnImport(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "文本文件|*.txt;*.csv",
                Multiselect = true,
                Title = "选择数据文件（TXT / CSV）"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            int total = 0;
            foreach (var path in dlg.FileNames)
            {
                var nums = ReadFile(path);
                total += ImportNumbers(nums, Path.GetFileName(path));
            }
            SaveData(); RefreshAll();
            MessageBox.Show(string.Format("成功导入 {0} 条", total), "导入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void OnPaste(object sender, EventArgs e)
        {
            try
            {
                var text = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(text)) { MessageBox.Show("剪贴板为空"); return; }
                var nums = new List<string>();
                foreach (var line in text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var n = Clean(line);
                    if (!string.IsNullOrEmpty(n)) nums.Add(n);
                }
                int added = ImportNumbers(nums, "粘贴导入");
                if (added > 0) { SaveData(); RefreshAll(); }
                MessageBox.Show(string.Format("成功导入 {0} 条", added), "导入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { MessageBox.Show("粘贴失败"); }
        }

        int ImportNumbers(List<string> nums, string source)
        {
            var existing = new HashSet<string>(data.Select(r => r.Number));
            int added = 0;
            foreach (var n in nums)
            {
                if (string.IsNullOrEmpty(n) || existing.Contains(n)) continue;
                var country = DetectCountry(n);
                var type = AutoDetectType(n);
                data.Add(new Record
                {
                    Number = n,
                    Country = country,
                    Type = type,
                    Channel = "短信",
                    Source = source,
                    Added = DateTime.Now
                });
                existing.Add(n);
                added++;
            }
            return added;
        }

        List<string> ReadFile(string path)
        {
            var nums = new List<string>();
            try
            {
                foreach (var line in File.ReadAllLines(path, DetectEncoding(path)))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var n = Clean(line);
                    if (!string.IsNullOrEmpty(n)) nums.Add(n);
                }
            }
            catch { }
            return nums;
        }

        Encoding DetectEncoding(string path)
        {
            try
            {
                var bytes = File.ReadAllBytes(path);
                if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return Encoding.Unicode;
                if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return Encoding.UTF8;
                return Encoding.Default;
            }
            catch { return Encoding.Default; }
        }

        void OnDelete(object sender, EventArgs e)
        {
            if (grid.SelectedRows.Count == 0) { MessageBox.Show("请先选择要删除的行"); return; }
            var toRemove = new List<Record>();
            foreach (DataGridViewRow row in grid.SelectedRows)
            {
                if (row.Index < filtered.Count)
                    toRemove.Add(filtered[row.Index]);
            }
            if (MessageBox.Show(string.Format("确认删除 {0} 条?", toRemove.Count), "确认", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            foreach (var r in toRemove) data.Remove(r);
            SaveData(); RefreshAll();
        }

        void OnClear(object sender, EventArgs e)
        {
            if (data.Count == 0) return;
            if (MessageBox.Show(string.Format("确认清空全部 {0} 条数据?", data.Count), "确认", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            data.Clear();
            SaveData(); RefreshAll();
        }

        void OnExport(object sender, EventArgs e)
        {
            if (filtered.Count == 0) { MessageBox.Show("无数据可导出"); return; }
            var dlg = new SaveFileDialog { Filter = "文本文件|*.txt", FileName = "export.txt" };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            try
            {
                var sb = new StringBuilder();
                foreach (var r in filtered)
                    sb.AppendLine(r.Number);
                File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show(string.Format("成功导出 {0} 条", filtered.Count), "导出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show("导出失败: " + ex.Message); }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveData();
            base.OnFormClosing(e);
        }
    }
}
