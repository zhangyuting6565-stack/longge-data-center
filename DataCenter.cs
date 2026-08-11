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
        }

        string Clean(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
                if (ch >= '0' && ch <= '9') { sb.Append(ch); }
                else if (ch == '+' && sb.Length == 0) { sb.Append(ch); }
            return sb.ToString();
        }

        string DetectCountry(string num)
        {
            if (string.IsNullOrEmpty(num) || !num.StartsWith("+")) return "未知";
            if (num.StartsWith("+998")) return "乌兹别克";
            if (num.StartsWith("+996")) return "吉尔吉斯斯坦";
            if (num.StartsWith("+995")) return "格鲁吉亚";
            if (num.StartsWith("+994")) return "阿塞拜疆";
            if (num.StartsWith("+993")) return "土库曼斯坦";
            if (num.StartsWith("+992")) return "塔吉克斯坦";
            if (num.StartsWith("+977")) return "尼泊尔";
            if (num.StartsWith("+976")) return "蒙古";
            if (num.StartsWith("+975")) return "不丹";
            if (num.StartsWith("+974")) return "卡塔尔";
            if (num.StartsWith("+973")) return "巴林";
            if (num.StartsWith("+972")) return "以色列";
            if (num.StartsWith("+971")) return "阿联酋";
            if (num.StartsWith("+970")) return "巴勒斯坦";
            if (num.StartsWith("+968")) return "阿曼";
            if (num.StartsWith("+967")) return "也门";
            if (num.StartsWith("+966")) return "沙特";
            if (num.StartsWith("+965")) return "科威特";
            if (num.StartsWith("+964")) return "伊拉克";
            if (num.StartsWith("+963")) return "叙利亚";
            if (num.StartsWith("+962")) return "约旦";
            if (num.StartsWith("+961")) return "黎巴嫩";
            if (num.StartsWith("+960")) return "马尔代夫";
            if (num.StartsWith("+886")) return "台湾";
            if (num.StartsWith("+880")) return "孟加拉";
            if (num.StartsWith("+856")) return "老挝";
            if (num.StartsWith("+855")) return "柬埔寨";
            if (num.StartsWith("+853")) return "澳门";
            if (num.StartsWith("+852")) return "香港";
            if (num.StartsWith("+850")) return "朝鲜";
            if (num.StartsWith("+692")) return "马绍尔";
            if (num.StartsWith("+691")) return "密克罗尼西亚";
            if (num.StartsWith("+690")) return "托克劳";
            if (num.StartsWith("+689")) return "法属波利尼西亚";
            if (num.StartsWith("+688")) return "图瓦卢";
            if (num.StartsWith("+687")) return "新喀里多尼亚";
            if (num.StartsWith("+686")) return "基里巴斯";
            if (num.StartsWith("+685")) return "萨摩亚";
            if (num.StartsWith("+683")) return "纽埃";
            if (num.StartsWith("+682")) return "库克群岛";
            if (num.StartsWith("+681")) return "瓦利斯";
            if (num.StartsWith("+680")) return "帕劳";
            if (num.StartsWith("+679")) return "斐济";
            if (num.StartsWith("+678")) return "瓦努阿图";
            if (num.StartsWith("+677")) return "所罗门";
            if (num.StartsWith("+676")) return "汤加";
            if (num.StartsWith("+675")) return "巴布亚新几内亚";
            if (num.StartsWith("+674")) return "瑙鲁";
            if (num.StartsWith("+673")) return "文莱";
            if (num.StartsWith("+672")) return "南极洲";
            if (num.StartsWith("+670")) return "东帝汶";
            if (num.StartsWith("+599")) return "荷属安的列斯";
            if (num.StartsWith("+598")) return "乌拉圭";
            if (num.StartsWith("+597")) return "苏里南";
            if (num.StartsWith("+596")) return "马提尼克";
            if (num.StartsWith("+595")) return "巴拉圭";
            if (num.StartsWith("+594")) return "法属圭亚那";
            if (num.StartsWith("+593")) return "厄瓜多尔";
            if (num.StartsWith("+592")) return "圭亚那";
            if (num.StartsWith("+591")) return "玻利维亚";
            if (num.StartsWith("+590")) return "瓜德罗普";
            if (num.StartsWith("+509")) return "海地";
            if (num.StartsWith("+508")) return "圣皮埃尔";
            if (num.StartsWith("+507")) return "巴拿马";
            if (num.StartsWith("+506")) return "哥斯达黎加";
            if (num.StartsWith("+505")) return "尼加拉瓜";
            if (num.StartsWith("+504")) return "洪都拉斯";
            if (num.StartsWith("+503")) return "萨尔瓦多";
            if (num.StartsWith("+502")) return "危地马拉";
            if (num.StartsWith("+501")) return "伯利兹";
            if (num.StartsWith("+500")) return "福克兰";
            if (num.StartsWith("+423")) return "列支敦士登";
            if (num.StartsWith("+421")) return "斯洛伐克";
            if (num.StartsWith("+420")) return "捷克";
            if (num.StartsWith("+389")) return "北马其顿";
            if (num.StartsWith("+387")) return "波黑";
            if (num.StartsWith("+386")) return "斯洛文尼亚";
            if (num.StartsWith("+385")) return "克罗地亚";
            if (num.StartsWith("+383")) return "科索沃";
            if (num.StartsWith("+382")) return "黑山";
            if (num.StartsWith("+381")) return "塞尔维亚";
            if (num.StartsWith("+380")) return "乌克兰";
            if (num.StartsWith("+378")) return "圣马力诺";
            if (num.StartsWith("+377")) return "摩纳哥";
            if (num.StartsWith("+376")) return "安道尔";
            if (num.StartsWith("+375")) return "白俄罗斯";
            if (num.StartsWith("+374")) return "亚美尼亚";
            if (num.StartsWith("+373")) return "摩尔多瓦";
            if (num.StartsWith("+372")) return "爱沙尼亚";
            if (num.StartsWith("+371")) return "拉脱维亚";
            if (num.StartsWith("+370")) return "立陶宛";
            if (num.StartsWith("+359")) return "保加利亚";
            if (num.StartsWith("+358")) return "芬兰";
            if (num.StartsWith("+357")) return "塞浦路斯";
            if (num.StartsWith("+356")) return "马耳他";
            if (num.StartsWith("+355")) return "阿尔巴尼亚";
            if (num.StartsWith("+354")) return "冰岛";
            if (num.StartsWith("+353")) return "爱尔兰";
            if (num.StartsWith("+352")) return "卢森堡";
            if (num.StartsWith("+351")) return "葡萄牙";
            if (num.StartsWith("+350")) return "直布罗陀";
            if (num.StartsWith("+299")) return "格陵兰";
            if (num.StartsWith("+298")) return "法罗群岛";
            if (num.StartsWith("+297")) return "阿鲁巴";
            if (num.StartsWith("+291")) return "厄立特里亚";
            if (num.StartsWith("+290")) return "圣赫勒拿";
            if (num.StartsWith("+269")) return "科摩罗";
            if (num.StartsWith("+268")) return "斯威士兰";
            if (num.StartsWith("+267")) return "博茨瓦纳";
            if (num.StartsWith("+266")) return "莱索托";
            if (num.StartsWith("+265")) return "马拉维";
            if (num.StartsWith("+264")) return "纳米比亚";
            if (num.StartsWith("+263")) return "津巴布韦";
            if (num.StartsWith("+262")) return "留尼汪";
            if (num.StartsWith("+261")) return "马达加斯加";
            if (num.StartsWith("+260")) return "赞比亚";
            if (num.StartsWith("+258")) return "莫桑比克";
            if (num.StartsWith("+257")) return "布隆迪";
            if (num.StartsWith("+256")) return "乌干达";
            if (num.StartsWith("+255")) return "坦桑尼亚";
            if (num.StartsWith("+254")) return "肯尼亚";
            if (num.StartsWith("+253")) return "吉布提";
            if (num.StartsWith("+252")) return "索马里";
            if (num.StartsWith("+251")) return "埃塞俄比亚";
            if (num.StartsWith("+250")) return "卢旺达";
            if (num.StartsWith("+249")) return "苏丹";
            if (num.StartsWith("+248")) return "塞舌尔";
            if (num.StartsWith("+247")) return "阿森松";
            if (num.StartsWith("+246")) return "迪戈加西亚";
            if (num.StartsWith("+245")) return "几内亚比绍";
            if (num.StartsWith("+244")) return "安哥拉";
            if (num.StartsWith("+243")) return "刚果金";
            if (num.StartsWith("+242")) return "刚果布";
            if (num.StartsWith("+241")) return "加蓬";
            if (num.StartsWith("+240")) return "赤道几内亚";
            if (num.StartsWith("+239")) return "圣多美";
            if (num.StartsWith("+238")) return "佛得角";
            if (num.StartsWith("+237")) return "喀麦隆";
            if (num.StartsWith("+236")) return "中非";
            if (num.StartsWith("+235")) return "乍得";
            if (num.StartsWith("+234")) return "尼日利亚";
            if (num.StartsWith("+233")) return "加纳";
            if (num.StartsWith("+232")) return "塞拉利昂";
            if (num.StartsWith("+231")) return "利比里亚";
            if (num.StartsWith("+230")) return "毛里求斯";
            if (num.StartsWith("+229")) return "贝宁";
            if (num.StartsWith("+228")) return "多哥";
            if (num.StartsWith("+227")) return "尼日尔";
            if (num.StartsWith("+226")) return "布基纳法索";
            if (num.StartsWith("+225")) return "科特迪瓦";
            if (num.StartsWith("+224")) return "几内亚";
            if (num.StartsWith("+223")) return "马里";
            if (num.StartsWith("+222")) return "毛里塔尼亚";
            if (num.StartsWith("+221")) return "塞内加尔";
            if (num.StartsWith("+220")) return "冈比亚";
            if (num.StartsWith("+218")) return "利比亚";
            if (num.StartsWith("+216")) return "突尼斯";
            if (num.StartsWith("+213")) return "阿尔及利亚";
            if (num.StartsWith("+212")) return "摩洛哥";
            if (num.StartsWith("+98")) return "伊朗";
            if (num.StartsWith("+95")) return "缅甸";
            if (num.StartsWith("+94")) return "斯里兰卡";
            if (num.StartsWith("+93")) return "阿富汗";
            if (num.StartsWith("+92")) return "巴基斯坦";
            if (num.StartsWith("+91")) return "印度";
            if (num.StartsWith("+90")) return "土耳其";
            if (num.StartsWith("+86")) return "中国";
            if (num.StartsWith("+84")) return "越南";
            if (num.StartsWith("+82")) return "韩国";
            if (num.StartsWith("+81")) return "日本";
            if (num.StartsWith("+66")) return "泰国";
            if (num.StartsWith("+65")) return "新加坡";
            if (num.StartsWith("+64")) return "新西兰";
            if (num.StartsWith("+63")) return "菲律宾";
            if (num.StartsWith("+62")) return "印尼";
            if (num.StartsWith("+61")) return "澳洲";
            if (num.StartsWith("+60")) return "马来西亚";
            if (num.StartsWith("+58")) return "委内瑞拉";
            if (num.StartsWith("+57")) return "哥伦比亚";
            if (num.StartsWith("+56")) return "智利";
            if (num.StartsWith("+55")) return "巴西";
            if (num.StartsWith("+54")) return "阿根廷";
            if (num.StartsWith("+53")) return "古巴";
            if (num.StartsWith("+52")) return "墨西哥";
            if (num.StartsWith("+51")) return "秘鲁";
            if (num.StartsWith("+49")) return "德国";
            if (num.StartsWith("+48")) return "波兰";
            if (num.StartsWith("+47")) return "挪威";
            if (num.StartsWith("+46")) return "瑞典";
            if (num.StartsWith("+45")) return "丹麦";
            if (num.StartsWith("+44")) return "英国";
            if (num.StartsWith("+43")) return "奥地利";
            if (num.StartsWith("+41")) return "瑞士";
            if (num.StartsWith("+40")) return "罗马尼亚";
            if (num.StartsWith("+39")) return "意大利";
            if (num.StartsWith("+36")) return "匈牙利";
            if (num.StartsWith("+34")) return "西班牙";
            if (num.StartsWith("+33")) return "法国";
            if (num.StartsWith("+32")) return "比利时";
            if (num.StartsWith("+31")) return "荷兰";
            if (num.StartsWith("+30")) return "希腊";
            if (num.StartsWith("+27")) return "南非";
            if (num.StartsWith("+20")) return "埃及";
            if (num.StartsWith("+7")) return "俄罗斯";
            if (num.StartsWith("+1")) return "美国/加拿大";
            return "其他";
        }

        string AutoDetectType(string num)
        {
            if (string.IsNullOrEmpty(num)) return "未知";
            if (num.StartsWith("+") && !num.StartsWith("+86")) return "国际";
            if (num.Length >= 11 && num.StartsWith("+86"))
            {
                var mobile = num.Length >= 13 ? num.Substring(3) : num;
                if (mobile.Length == 11 && mobile[0] == '1' && mobile[1] >= '3' && mobile[1] <= '9')
                    return "手机";
                return "固话";
            }
            if (num.Length == 11 && num[0] == '1' && num[1] >= '3' && num[1] <= '9') return "手机";
            if (num.Length >= 10 && num[0] == '0') return "固话";
            return "未知";
        }

        void RefreshAll()
        {
            ApplyFilter();
        }

        void ApplyFilter()
        {
            var search = (txtSearch.Text ?? "").Trim().ToLower();
            var country = cbCountry.SelectedItem as string;
            if (country == "全部") country = null;

            filtered = data.Where(r =>
            {
                if (country != null && r.Country != country) return false;
                if (!string.IsNullOrEmpty(search))
                {
                    return r.Number.ToLower().Contains(search) ||
                           r.Country.ToLower().Contains(search) ||
                           r.Type.ToLower().Contains(search) ||
                           r.Channel.ToLower().Contains(search);
                }
                return true;
            }).ToList();
            RefreshGrid();
            RefreshTree();
        }

        void RefreshGrid()
        {
            grid.Rows.Clear();
            foreach (var r in filtered)
                grid.Rows.Add(r.Number, r.Country, r.Type, r.Channel, r.Source);
            lblStatus.Text = string.Format("共 {0:N0} 条", filtered.Count);
        }

        void RefreshTree()
        {
            tree.Nodes.Clear();
            var byCountry = data.GroupBy(r => r.Country)
                .OrderByDescending(g => g.Count());
            var allCountryNames = byCountry.Select(g => g.Key).ToList();
            allCountryNames.Insert(0, "全部");
            cbCountry.Items.Clear();
            foreach (var name in allCountryNames)
                cbCountry.Items.Add(name);
            if (cbCountry.Items.Count > 0) cbCountry.SelectedIndex = 0;

            foreach (var g in byCountry)
            {
                var countryNode = new TreeNode(string.Format("{0} ({1:N0})", g.Key, g.Count()));
                countryNode.ForeColor = Color.FromArgb(0x00, 0x70, 0xC0);
                countryNode.NodeFont = new Font(tree.Font, FontStyle.Bold);
                var byType = g.GroupBy(r => r.Type).OrderByDescending(t => t.Count());
                foreach (var t in byType)
                {
                    var typeNode = new TreeNode(string.Format("{0} ({1:N0})", t.Key, t.Count()));
                    typeNode.ForeColor = Color.FromArgb(0x33, 0x33, 0x33);
                    var byChannel = t.GroupBy(r => r.Channel).OrderByDescending(c => c.Count());
                    foreach (var ch in byChannel)
                    {
                        var chNode = new TreeNode(string.Format("{0} ({1:N0})", ch.Key, ch.Count()));
                        chNode.ForeColor = Color.FromArgb(0x66, 0x66, 0x66);
                        typeNode.Nodes.Add(chNode);
                    }
                    countryNode.Nodes.Add(typeNode);
                }
                tree.Nodes.Add(countryNode);
            }
            if (tree.Nodes.Count > 0) tree.Nodes[0].Expand();
        }

        void OnImport(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "文本文件|*.txt;*.csv|Excel文件|*.xlsx|所有文件|*.*",
                Multiselect = true
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            int added = 0;
            foreach (var path in dlg.FileNames)
            {
                var source = Path.GetFileNameWithoutExtension(path);
                var nums = ReadFile(path);
                added += ImportNumbers(nums, source);
            }
            if (added > 0) { SaveData(); RefreshAll(); }
            MessageBox.Show(string.Format("成功导入 {0} 条", added), "导入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                var ext = Path.GetExtension(path).ToLower();
                if (ext == ".xlsx") return ReadXlsx(path);
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

        List<string> ReadXlsx(string path)
        {
            var nums = new List<string>();
            try
            {
                var bytes = File.ReadAllBytes(path);
                using (var ms = new MemoryStream(bytes))
                using (var pkg = System.IO.Packaging.Package.Open(ms))
                {
                    var uri = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
                    var part = pkg.GetPart(uri);
                    var xml = "";
                    using (var sr = new StreamReader(part.GetStream()))
                        xml = sr.ReadToEnd();
                    int idx = 0;
                    while ((idx = xml.IndexOf("<t>", idx)) >= 0)
                    {
                        var end = xml.IndexOf("</t>", idx);
                        if (end < 0) break;
                        var val = xml.Substring(idx + 3, end - idx - 3);
                        var n = Clean(val);
                        if (!string.IsNullOrEmpty(n)) nums.Add(n);
                        idx = end + 4;
                    }
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
