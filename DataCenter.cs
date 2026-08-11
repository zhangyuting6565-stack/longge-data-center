using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

using System.Reflection;

namespace DataCenter
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            // 从嵌入资源释放 SQLite.Interop.dll
            ExtractInterop();
            // 注册 AssemblyResolve 加载嵌入的 System.Data.SQLite.dll
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            string name = new AssemblyName(args.Name).Name;
            if (name == "System.Data.SQLite")
            {
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("System.Data.SQLite.dll"))
                {
                    if (stream != null)
                    {
                        byte[] raw = new byte[stream.Length];
                        stream.Read(raw, 0, raw.Length);
                        return Assembly.Load(raw);
                    }
                }
            }
            return null;
        }

        static void ExtractInterop()
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            string path = Path.Combine(dir, "SQLite.Interop.dll");
            // 已存在则跳过（避免覆盖）
            if (File.Exists(path)) return;

            try
            {
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("SQLite.Interop.dll"))
                {
                    if (stream != null)
                    {
                        using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                        {
                            byte[] buf = new byte[8192];
                            int n;
                            while ((n = stream.Read(buf, 0, buf.Length)) > 0)
                                fs.Write(buf, 0, n);
                        }
                    }
                }
            }
            catch
            {
                // 释放失败：依赖外部 DLL，用户需自行放置
            }
        }
    }

    class MainForm : Form
    {
        // ── 数据库 ──
        string dbPath;
        SQLiteConnection db;
        DataTable dtCache;       // 当前查询结果缓存
        long totalCount;

        // ── UI 控件 ──
        DataGridView grid;
        Label lStat, lSummary;
        TextBox tbSearch;
        ComboBox cbCountry, cbType, cbChannel;
        Button bImport, bPaste, bExport, bClear, bDel, bRefresh, bSetType, bSetChannel;
        Panel topPanel, rightPanel;

        public MainForm()
        {
            Text = "龙哥数据中心";
            Size = new Size(1300, 750);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(0x1a, 0x1a, 0x1a);
            ForeColor = Color.White;
            MinimumSize = new Size(1000, 550);

            // 数据库路径: exe 同目录
            dbPath = Path.Combine(Application.StartupPath, "data.db");
            InitDB();
            BuildUI();
            RefreshStats();
        }

        // ═══════════════════════════════════════════════
        // 数据库初始化
        // ═══════════════════════════════════════════════
        void InitDB()
        {
            db = new SQLiteConnection(string.Format("Data Source={0};Version=3;", dbPath));
            db.Open();
            using (var cmd = new SQLiteCommand(@"
                CREATE TABLE IF NOT EXISTS numbers (
                    id      INTEGER PRIMARY KEY AUTOINCREMENT,
                    number  TEXT NOT NULL UNIQUE,
                    country TEXT,
                    type    TEXT DEFAULT '',
                    channel TEXT DEFAULT '短信',
                    prefix  TEXT,
                    source  TEXT,
                    added   TEXT DEFAULT (datetime('now','localtime'))
                );
                CREATE INDEX IF NOT EXISTS idx_country ON numbers(country);
                CREATE INDEX IF NOT EXISTS idx_type   ON numbers(type);
                CREATE INDEX IF NOT EXISTS idx_channel ON numbers(channel);
                CREATE INDEX IF NOT EXISTS idx_prefix  ON numbers(prefix);
                CREATE INDEX IF NOT EXISTS idx_added   ON numbers(added);
            ", db))
            {
                cmd.ExecuteNonQuery();
            }

            // 迁移: 旧库添加 type/channel 列
            try { ExecSQL("ALTER TABLE numbers ADD COLUMN type TEXT DEFAULT ''"); } catch { }
            try { ExecSQL("ALTER TABLE numbers ADD COLUMN channel TEXT DEFAULT '短信'"); } catch { }
        }

        void ExecSQL(string sql)
        {
            using (var cmd = new SQLiteCommand(sql, db))
                cmd.ExecuteNonQuery();
        }

        DataTable QuerySQL(string sql)
        {
            var dt = new DataTable();
            using (var cmd = new SQLiteCommand(sql, db))
            using (var adapter = new SQLiteDataAdapter(cmd))
                adapter.Fill(dt);
            return dt;
        }

        object ScalarSQL(string sql)
        {
            using (var cmd = new SQLiteCommand(sql, db))
                return cmd.ExecuteScalar();
        }

        // ═══════════════════════════════════════════════
        // UI 构建
        // ═══════════════════════════════════════════════
        void BuildUI()
        {
            // 顶部面板
            topPanel = new Panel { Dock = DockStyle.Top, Height = 90, BackColor = Color.FromArgb(0x22, 0x22, 0x22) };
            Controls.Add(topPanel);

            int y1 = 8;

            // 导入按钮
            bImport = Btn("导入文件", 10, y1, 85, 28, OnImportFile);
            bPaste = Btn("粘贴导入", 100, y1, 85, 28, OnPasteImport);
            bClear = Btn("清空库", 190, y1, 75, 28, OnClearAll);

            // 搜索控件
            cbCountry = new ComboBox { Left = 280, Top = y1 + 1, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft YaHei", 9) };
            cbCountry.BackColor = Color.FromArgb(0x33, 0x33, 0x33);
            cbCountry.ForeColor = Color.White;
            LoadCountries();
            topPanel.Controls.Add(cbCountry);

            cbType = new ComboBox { Left = 388, Top = y1 + 1, Width = 88, DropDownStyle = ComboBoxStyle.DropDownList, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft YaHei", 9) };
            cbType.BackColor = Color.FromArgb(0x33, 0x33, 0x33);
            cbType.ForeColor = Color.White;
            LoadTypes();
            topPanel.Controls.Add(cbType);

            cbChannel = new ComboBox { Left = 484, Top = y1 + 1, Width = 88, DropDownStyle = ComboBoxStyle.DropDown, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft YaHei", 9) };
            cbChannel.BackColor = Color.FromArgb(0x33, 0x33, 0x33);
            cbChannel.ForeColor = Color.White;
            cbChannel.Items.AddRange(new object[] { "短信", "WS", "TG", "RCS", "iOS群发专用" });
            cbChannel.Text = "短信";
            topPanel.Controls.Add(cbChannel);

            tbSearch = new TextBox { Left = 580, Top = y1 + 1, Width = 130, Font = new Font("Microsoft YaHei", 9) };
            tbSearch.BackColor = Color.FromArgb(0x33, 0x33, 0x33);
            tbSearch.ForeColor = Color.White;
            tbSearch.BorderStyle = BorderStyle.FixedSingle;
            topPanel.Controls.Add(tbSearch);

            var bSearch = Btn("搜索", 718, y1, 60, 28, OnSearch);

            // 第一行右侧按钮
            bRefresh = Btn("刷新", 788, y1, 60, 28, (s, e) => RefreshAll());
            bExport = Btn("导出", 856, y1, 60, 28, OnExportResult);
            bDel = Btn("删选中", 924, y1, 70, 28, OnDeleteSelected);

            // 第二行: 批量操作
            int y2 = 44;
            var l = new Label { Text = "选中行批量:", Left = 10, Top = y2 + 5, AutoSize = true, ForeColor = Color.FromArgb(0x88, 0x88, 0x88), Font = new Font("Microsoft YaHei", 9) };
            topPanel.Controls.Add(l);

            bSetType = Btn("改类型", 90, y2, 75, 26, OnBatchSetType);
            bSetChannel = Btn("改渠道", 172, y2, 75, 26, OnBatchSetChannel);

            // 统计标签
            lStat = new Label
            {
                Left = 260, Top = y2 + 5, AutoSize = true,
                ForeColor = Color.FromArgb(0x88, 0x88, 0x88),
                Text = "共 0 条", Font = new Font("Microsoft YaHei", 9)
            };
            topPanel.Controls.Add(lStat);

            // 右侧汇总面板
            rightPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 280,
                BackColor = Color.FromArgb(0x22, 0x22, 0x22),
                Padding = new Padding(8)
            };
            lSummary = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei", 9),
                Text = "加载中…"
            };
            rightPanel.Controls.Add(lSummary);
            Controls.Add(rightPanel);

            // 数据表格
            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.FromArgb(0x1a, 0x1a, 0x1a),
                ForeColor = Color.White,
                GridColor = Color.FromArgb(0x33, 0x33, 0x33),
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                EnableHeadersVisualStyles = false,
                VirtualMode = true,
                RowCount = 0
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0x2a, 0x2a, 0x2a);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei", 9, FontStyle.Bold);
            grid.DefaultCellStyle.BackColor = Color.FromArgb(0x1a, 0x1a, 0x1a);
            grid.DefaultCellStyle.ForeColor = Color.White;
            grid.DefaultCellStyle.Font = new Font("Consolas", 10);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0x33, 0x66, 0x99);
            grid.Columns.Add("col_id", "ID");
            grid.Columns["col_id"].Width = 55;
            grid.Columns["col_id"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            grid.Columns.Add("col_number", "号码");
            grid.Columns["col_number"].Width = 140;
            grid.Columns["col_number"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            grid.Columns.Add("col_country", "国家");
            grid.Columns["col_country"].Width = 80;
            grid.Columns["col_country"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            grid.Columns.Add("col_type", "类型");
            grid.Columns["col_type"].Width = 65;
            grid.Columns["col_type"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            grid.Columns.Add("col_channel", "渠道");
            grid.Columns["col_channel"].Width = 65;
            grid.Columns["col_channel"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            grid.Columns.Add("col_prefix", "前缀");
            grid.Columns.Add("col_source", "来源");
            grid.Columns.Add("col_added", "时间");
            grid.CellValueNeeded += OnCellValueNeeded;
            Controls.Add(grid);
        }

        Button Btn(string text, int x, int y, int w, int h, EventHandler handler)
        {
            var b = new Button
            {
                Text = text, Left = x, Top = y, Width = w, Height = h,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0x44, 0x44, 0x44),
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei", 9)
            };
            b.FlatAppearance.BorderSize = 0;
            b.Click += handler;
            topPanel.Controls.Add(b);
            return b;
        }

        // ═══════════════════════════════════════════════
        // 国家列表
        // ═══════════════════════════════════════════════
        void LoadCountries()
        {
            cbCountry.Items.Clear();
            cbCountry.Items.Add("-- 全部国家 --");
            try
            {
                var dt = QuerySQL("SELECT DISTINCT country FROM numbers WHERE country IS NOT NULL AND country != '' ORDER BY country");
                foreach (DataRow row in dt.Rows)
                    cbCountry.Items.Add(row[0].ToString());
            }
            catch { }
            cbCountry.SelectedIndex = 0;
        }

        // ═══════════════════════════════════════════════
        // 类型列表 (自动从库中收集 + 预设)
        // ═══════════════════════════════════════════════
        void LoadTypes()
        {
            cbType.Items.Clear();
            cbType.Items.Add("-- 全部类型 --");
            try
            {
                var dt = QuerySQL("SELECT DISTINCT type FROM numbers WHERE type IS NOT NULL AND type != '' ORDER BY type");
                foreach (DataRow row in dt.Rows)
                    cbType.Items.Add(row[0].ToString());
            }
            catch { }
            cbType.SelectedIndex = 0;
        }

        // ═══════════════════════════════════════════════
        // 号码类型自动检测
        // ═══════════════════════════════════════════════
        static string AutoDetectType(string num)
        {
            if (string.IsNullOrEmpty(num)) return "";
            // 11位 1[3-9] → 手机
            if (num.Length == 11 && num[0] == '1' && num[1] >= '3' && num[1] <= '9')
                return "手机";
            // 0开头 ≥10位 → 固话
            if (num.StartsWith("0") && num.Length >= 10)
                return "固话";
            // + / 00 开头 → 国际
            if (num.StartsWith("+") || num.StartsWith("00"))
                return "国际";
            // 10-12位 非0/1开头 → 国际
            if (num.Length >= 8 && num.Length <= 15)
            {
                char first = num[0];
                if (first != '0' && first != '1' && num.StartsWith("86") == false)
                    return "国际";
            }
            return "其他";
        }

        // ═══════════════════════════════════════════════
        // 汇总面板
        // ═══════════════════════════════════════════════
        void RefreshSummary()
        {
            try
            {
                var sb = new StringBuilder();

                // 标题
                sb.AppendLine("◆ 数据中心汇总");
                sb.AppendLine();

                // 总计
                long total = Convert.ToInt64(ScalarSQL("SELECT COUNT(*) FROM numbers"));
                sb.AppendLine(string.Format("总计: {0:n0} 条", total));
                sb.AppendLine();

                // 按国家
                sb.AppendLine("── 按国家 ──");
                var dtC = QuerySQL("SELECT country, COUNT(*) cnt FROM numbers WHERE country IS NOT NULL AND country != '' GROUP BY country ORDER BY cnt DESC LIMIT 20");
                foreach (DataRow row in dtC.Rows)
                    sb.AppendLine(string.Format("  {0}: {1}", row["country"], row["cnt"]));
                sb.AppendLine();

                // 按类型
                sb.AppendLine("── 按类型 ──");
                var dtT = QuerySQL("SELECT type, COUNT(*) cnt FROM numbers WHERE type IS NOT NULL AND type != '' GROUP BY type ORDER BY cnt DESC");
                foreach (DataRow row in dtT.Rows)
                    sb.AppendLine(string.Format("  {0}: {1}", row["type"], row["cnt"]));
                sb.AppendLine();

                // 按渠道
                sb.AppendLine("── 按渠道 ──");
                var dtCh = QuerySQL("SELECT channel, COUNT(*) cnt FROM numbers WHERE channel IS NOT NULL AND channel != '' GROUP BY channel ORDER BY cnt DESC");
                foreach (DataRow row in dtCh.Rows)
                    sb.AppendLine(string.Format("  {0}: {1}", row["channel"], row["cnt"]));

                lSummary.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                lSummary.Text = "汇总加载失败: " + ex.Message;
            }
        }

        // ═══════════════════════════════════════════════
        // 国家检测 (复用筛选器逻辑)
        // ═══════════════════════════════════════════════
        static Dictionary<string, string> _countryMap;
        static Dictionary<string, string> CountryMap()
        {
            if (_countryMap != null) return _countryMap;
            var m = new Dictionary<string, string>();
            // 东亚
            m["86"] = "中国"; m["852"] = "香港"; m["853"] = "澳门"; m["886"] = "台湾";
            m["81"] = "日本"; m["82"] = "韩国"; m["850"] = "朝鲜"; m["976"] = "蒙古";
            // 东南亚
            m["84"] = "越南"; m["66"] = "泰国"; m["62"] = "印尼"; m["60"] = "马来西亚";
            m["65"] = "新加坡"; m["63"] = "菲律宾"; m["95"] = "缅甸"; m["855"] = "柬埔寨";
            m["856"] = "老挝"; m["673"] = "文莱"; m["670"] = "东帝汶";
            // 南亚
            m["91"] = "印度"; m["92"] = "巴基斯坦"; m["880"] = "孟加拉国"; m["94"] = "斯里兰卡";
            m["977"] = "尼泊尔"; m["960"] = "马尔代夫"; m["975"] = "不丹";
            // 中亚/独联体
            m["7"] = "俄罗斯/哈萨克斯坦"; m["998"] = "乌兹别克斯坦"; m["993"] = "土库曼斯坦";
            m["992"] = "塔吉克斯坦"; m["996"] = "吉尔吉斯斯坦"; m["994"] = "阿塞拜疆";
            m["374"] = "亚美尼亚"; m["995"] = "格鲁吉亚";
            // 中东
            m["98"] = "伊朗"; m["90"] = "土耳其"; m["964"] = "伊拉克"; m["966"] = "沙特";
            m["971"] = "阿联酋"; m["974"] = "卡塔尔"; m["973"] = "巴林"; m["968"] = "阿曼";
            m["965"] = "科威特"; m["967"] = "也门"; m["962"] = "约旦"; m["961"] = "黎巴嫩";
            m["963"] = "叙利亚"; m["972"] = "以色列"; m["970"] = "巴勒斯坦";
            // 欧洲
            m["44"] = "英国"; m["49"] = "德国"; m["33"] = "法国"; m["39"] = "意大利";
            m["34"] = "西班牙"; m["31"] = "荷兰"; m["32"] = "比利时"; m["41"] = "瑞士";
            m["43"] = "奥地利"; m["46"] = "瑞典"; m["47"] = "挪威"; m["45"] = "丹麦";
            m["358"] = "芬兰"; m["48"] = "波兰"; m["380"] = "乌克兰"; m["40"] = "罗马尼亚";
            m["420"] = "捷克"; m["36"] = "匈牙利"; m["351"] = "葡萄牙"; m["30"] = "希腊";
            m["353"] = "爱尔兰"; m["381"] = "塞尔维亚"; m["385"] = "克罗地亚";
            // 非洲
            m["20"] = "埃及"; m["234"] = "尼日利亚"; m["254"] = "肯尼亚"; m["27"] = "南非";
            m["256"] = "乌干达"; m["255"] = "坦桑尼亚"; m["251"] = "埃塞俄比亚";
            m["233"] = "加纳"; m["225"] = "科特迪瓦"; m["237"] = "喀麦隆";
            m["221"] = "塞内加尔"; m["212"] = "摩洛哥"; m["213"] = "阿尔及利亚";
            m["216"] = "突尼斯"; m["218"] = "利比亚"; m["249"] = "苏丹";
            // 北美
            m["1"] = "美国/加拿大";
            // 拉美
            m["52"] = "墨西哥"; m["55"] = "巴西"; m["54"] = "阿根廷"; m["56"] = "智利";
            m["57"] = "哥伦比亚"; m["51"] = "秘鲁"; m["58"] = "委内瑞拉";
            m["53"] = "古巴"; m["506"] = "哥斯达黎加"; m["507"] = "巴拿马";
            // 大洋洲
            m["61"] = "澳大利亚"; m["64"] = "新西兰";
            _countryMap = m;
            return m;
        }

        static string DetectCountry(string num)
        {
            if (string.IsNullOrEmpty(num)) return "";
            string raw = num.Trim().Replace(" ", "").Replace("-", "");
            string clean = raw.Replace("+", "");
            bool hasIntlPrefix = raw.StartsWith("+") || raw.StartsWith("00");

            if (hasIntlPrefix)
            {
                if (raw.StartsWith("00")) clean = raw.Substring(2).Replace("+", "");
                else clean = raw.Replace("+", "");
                var m = CountryMap();
                var keys = new List<string>(m.Keys);
                keys.Sort((a, b) => b.Length.CompareTo(a.Length));
                foreach (var k in keys)
                {
                    if (clean.StartsWith(k)) return m[k];
                }
                return "未知区域";
            }

            // 无国际前缀: 11位1[3-9]开头→中国
            if (clean.Length == 11 && clean[0] == '1' && clean[1] >= '3' && clean[1] <= '9')
                return "中国";

            // 0开头≥10位→中国固话
            if (clean.StartsWith("0") && clean.Length >= 10) return "中国";

            // 其他前缀匹配
            var m2 = CountryMap();
            var keys2 = new List<string>(m2.Keys);
            keys2.Sort((a, b) => b.Length.CompareTo(a.Length));
            foreach (var k in keys2)
            {
                if (clean.StartsWith(k) && k != "86") return m2[k];
            }
            return "未知区域";
        }

        // ═══════════════════════════════════════════════
        // 号码清洗
        // ═══════════════════════════════════════════════
        static string Clean(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            string trimmed = raw.Trim();
            var sb = new StringBuilder(trimmed.Length);
            for (int i = 0; i < trimmed.Length; i++)
            {
                char c = trimmed[i];
                if (c == ' ' || c == '\t' || c == '-' || c == '+' || c == '(' || c == ')' || c == '.' || c == ',') continue;
                sb.Append(c);
            }
            return sb.ToString();
        }

        static bool IsAllDigits(string s)
        {
            for (int i = 0; i < s.Length; i++) if (s[i] < '0' || s[i] > '9') return false;
            return true;
        }

        // ═══════════════════════════════════════════════
        // 导入 ── 从 List<string> 批量插入
        // ═══════════════════════════════════════════════
        int ImportNumbers(List<string> numbers, string source, string channel)
        {
            if (numbers.Count == 0) return 0;
            int inserted = 0;

            using (var trans = db.BeginTransaction())
            {
                using (var cmd = new SQLiteCommand(
                    "INSERT OR IGNORE INTO numbers (number, country, type, channel, prefix, source) VALUES (@n, @c, @t, @ch, @p, @s)", db))
                {
                    cmd.Parameters.Add("@n", DbType.String);
                    cmd.Parameters.Add("@c", DbType.String);
                    cmd.Parameters.Add("@t", DbType.String);
                    cmd.Parameters.Add("@ch", DbType.String);
                    cmd.Parameters.Add("@p", DbType.String);
                    cmd.Parameters.Add("@s", DbType.String);

                    foreach (var num in numbers)
                    {
                        string country = DetectCountry(num);
                        string type = AutoDetectType(num);
                        string prefix = num.Length >= 5 ? num.Substring(0, 5) : num;
                        cmd.Parameters["@n"].Value = num;
                        cmd.Parameters["@c"].Value = country;
                        cmd.Parameters["@t"].Value = type;
                        cmd.Parameters["@ch"].Value = channel;
                        cmd.Parameters["@p"].Value = prefix;
                        cmd.Parameters["@s"].Value = source;
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0) inserted++;
                    }
                }
                trans.Commit();
            }

            return inserted;
        }

        // ═══════════════════════════════════════════════
        // 导入 ── TXT/CSV 文件
        // ═══════════════════════════════════════════════
        void OnImportFile(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "选择数据文件",
                Filter = "文本文件|*.txt;*.csv|Excel文件|*.xlsx;*.xls|所有文件|*.*"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string ext = Path.GetExtension(dlg.FileName).ToLower();
            List<string> numbers;

            if (ext == ".xlsx" || ext == ".xls")
            {
                numbers = ReadXlsx(dlg.FileName);
            }
            else
            {
                numbers = ReadTextFile(dlg.FileName);
            }

            if (numbers.Count == 0)
            {
                MessageBox.Show("未读取到有效号码", "提示");
                return;
            }

            Enabled = false;
            string source = Path.GetFileNameWithoutExtension(dlg.FileName);
            string channel = cbChannel.Text;
            if (string.IsNullOrEmpty(channel)) channel = "短信";
            int inserted = ImportNumbers(numbers, source, channel);
            Enabled = true;
            LoadCountries();
            LoadTypes();
            RefreshAll();
            MessageBox.Show(string.Format("导入完成: {0:n0} 条 (共读取 {1:n0} 行)", inserted, numbers.Count), "完成");
        }

        List<string> ReadTextFile(string path)
        {
            var list = new List<string>();
            try
            {
                using (var sr = new StreamReader(path, Encoding.UTF8))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        var num = Clean(line);
                        if (num.Length > 0 && IsAllDigits(num)) list.Add(num);
                    }
                }
            }
            catch
            {
                list.Clear();
                using (var sr = new StreamReader(path, Encoding.GetEncoding("GBK")))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        var num = Clean(line);
                        if (num.Length > 0 && IsAllDigits(num)) list.Add(num);
                    }
                }
            }
            return list;
        }

        // ═══════════════════════════════════════════════
        // 导入 ── XLSX (WinBase Package API, .NET 4.0)
        // ═══════════════════════════════════════════════
        List<string> ReadXlsx(string path)
        {
            var numbers = new List<string>();
            try
            {
                using (var package = Package.Open(path, FileMode.Open, FileAccess.Read))
                {
                    // 读取共享字符串表
                    var sharedStrings = new List<string>();
                    var ssUri = PackUriHelper.CreatePartUri(new Uri("/xl/sharedStrings.xml", UriKind.Relative));
                    if (package.PartExists(ssUri))
                    {
                        var ssPart = package.GetPart(ssUri);
                        using (var stream = ssPart.GetStream())
                        {
                            var doc = XDocument.Load(stream);
                            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                            foreach (var si in doc.Descendants(ns + "si"))
                            {
                                var t = si.Element(ns + "t");
                                sharedStrings.Add(t != null ? t.Value : "");
                            }
                        }
                    }

                    // 读取第一个工作表
                    var sheetUri = PackUriHelper.CreatePartUri(new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative));
                    if (!package.PartExists(sheetUri)) return numbers;
                    var sheetPart = package.GetPart(sheetUri);
                    using (var stream = sheetPart.GetStream())
                    {
                        var doc = XDocument.Load(stream);
                        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                        foreach (var row in doc.Descendants(ns + "row"))
                        {
                            var cell = row.Elements(ns + "c").FirstOrDefault();
                            if (cell == null) continue;
                            string val = GetCellValue(cell, sharedStrings, ns);
                            if (!string.IsNullOrEmpty(val))
                            {
                                var num = Clean(val);
                                if (num.Length > 0 && IsAllDigits(num))
                                    numbers.Add(num);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("读取 Excel 失败: " + ex.Message, "错误");
            }
            return numbers;
        }

        string GetCellValue(XElement cell, List<string> sharedStrings, XNamespace ns)
        {
            var v = cell.Element(ns + "v");
            if (v == null) return "";
            string type = (string)cell.Attribute("t");
            if (type == "s")
            {
                int idx;
                if (int.TryParse(v.Value, out idx) && idx >= 0 && idx < sharedStrings.Count)
                    return sharedStrings[idx];
                return "";
            }
            return v.Value;
        }

        // ═══════════════════════════════════════════════
        // 导入 ── 粘贴
        // ═══════════════════════════════════════════════
        void OnPasteImport(object sender, EventArgs e)
        {
            try
            {
                string text = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(text))
                {
                    MessageBox.Show("剪贴板为空", "提示");
                    return;
                }

                var numbers = new List<string>();
                foreach (var line in text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var num = Clean(line);
                    if (num.Length > 0 && IsAllDigits(num)) numbers.Add(num);
                }

                if (numbers.Count == 0)
                {
                    MessageBox.Show("未识别到有效号码", "提示");
                    return;
                }

                string ch = cbChannel.Text;
                if (string.IsNullOrEmpty(ch)) ch = "短信";
                int inserted = ImportNumbers(numbers, "粘贴导入", ch);
                LoadCountries();
                LoadTypes();
                RefreshAll();
                MessageBox.Show(string.Format("导入完成: {0:n0} 条 (共 {1:n0} 行)", inserted, numbers.Count), "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("粘贴失败: " + ex.Message, "错误");
            }
        }

        // ═══════════════════════════════════════════════
        // 搜索
        // ═══════════════════════════════════════════════
        void OnSearch(object sender, EventArgs e)
        {
            string country = cbCountry.SelectedIndex > 0 ? cbCountry.SelectedItem.ToString() : "";
            string type = cbType.SelectedIndex > 0 ? cbType.SelectedItem.ToString() : "";
            string keyword = tbSearch.Text.Trim();

            var sql = new StringBuilder("SELECT * FROM numbers WHERE 1=1");
            if (!string.IsNullOrEmpty(country))
                sql.AppendFormat(" AND country = '{0}'", country.Replace("'", "''"));
            if (!string.IsNullOrEmpty(type))
                sql.AppendFormat(" AND type = '{0}'", type.Replace("'", "''"));
            if (!string.IsNullOrEmpty(keyword))
                sql.AppendFormat(" AND (number LIKE '%{0}%' OR prefix LIKE '{0}%')", keyword.Replace("'", "''"));
            sql.Append(" ORDER BY id DESC LIMIT 50000");

            Enabled = false;
            dtCache = QuerySQL(sql.ToString());
            totalCount = dtCache.Rows.Count;
            grid.RowCount = dtCache.Rows.Count;
            grid.Refresh();
            lStat.Text = string.Format("结果: {0:n0} 条", totalCount);
            Enabled = true;
        }

        // ═══════════════════════════════════════════════
        // VirtualMode 回调
        // ═══════════════════════════════════════════════
        void OnCellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (dtCache == null || e.RowIndex >= dtCache.Rows.Count) return;
            var row = dtCache.Rows[e.RowIndex];
            switch (e.ColumnIndex)
            {
                case 0: e.Value = row["id"]; break;
                case 1: e.Value = row["number"]; break;
                case 2: e.Value = row["country"]; break;
                case 3: e.Value = row["type"]; break;
                case 4: e.Value = row["channel"]; break;
                case 5: e.Value = row["prefix"]; break;
                case 6: e.Value = row["source"]; break;
                case 7: e.Value = row["added"]; break;
            }
        }

        // ═══════════════════════════════════════════════
        // 刷新
        // ═══════════════════════════════════════════════
        void RefreshAll()
        {
            RefreshStats();
            RefreshSummary();
            dtCache = QuerySQL("SELECT * FROM numbers ORDER BY id DESC LIMIT 50000");
            grid.RowCount = dtCache.Rows.Count;
            grid.Refresh();
        }

        void RefreshStats()
        {
            try
            {
                totalCount = Convert.ToInt64(ScalarSQL("SELECT COUNT(*) FROM numbers"));
                lStat.Text = string.Format("共 {0:n0} 条", totalCount);
            }
            catch { }
        }

        // ═══════════════════════════════════════════════
        // 导出
        // ═══════════════════════════════════════════════
        void OnExportResult(object sender, EventArgs e)
        {
            if (dtCache == null || dtCache.Rows.Count == 0)
            {
                MessageBox.Show("无数据可导出", "提示");
                return;
            }

            var dlg = new SaveFileDialog
            {
                Title = "导出结果",
                Filter = "文本文件|*.txt",
                FileName = string.Format("export_{0}.txt", DateTime.Now.ToString("yyyyMMddHHmmss"))
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                var lines = new List<string>();
                foreach (DataRow row in dtCache.Rows)
                    lines.Add(row["number"].ToString());
                File.WriteAllLines(dlg.FileName, lines, Encoding.UTF8);
                MessageBox.Show(string.Format("导出完成: {0:n0} 条", lines.Count), "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出失败: " + ex.Message, "错误");
            }
        }

        // ═══════════════════════════════════════════════
        // 删除选中
        // ═══════════════════════════════════════════════
        void OnDeleteSelected(object sender, EventArgs e)
        {
            if (grid.SelectedRows.Count == 0) return;
            if (MessageBox.Show(string.Format("确认删除选中的 {0} 条?", grid.SelectedRows.Count), "确认", MessageBoxButtons.YesNo) != DialogResult.Yes) return;

            try
            {
                foreach (DataGridViewRow row in grid.SelectedRows)
                {
                    if (row.Index < dtCache.Rows.Count)
                    {
                        long id = Convert.ToInt64(dtCache.Rows[row.Index]["id"]);
                        ExecSQL(string.Format("DELETE FROM numbers WHERE id = {0}", id));
                    }
                }
                LoadCountries();
                LoadTypes();
                RefreshAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show("删除失败: " + ex.Message, "错误");
            }
        }

        // ═══════════════════════════════════════════════
        // 清空库
        // ═══════════════════════════════════════════════
        void OnClearAll(object sender, EventArgs e)
        {
            if (MessageBox.Show("确认清空全部数据? 此操作不可恢复!", "危险操作", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            ExecSQL("DELETE FROM numbers");
            LoadCountries();
            LoadTypes();
            RefreshAll();
        }

        // ═══════════════════════════════════════════════
        // 批量改类型
        // ═══════════════════════════════════════════════
        void OnBatchSetType(object sender, EventArgs e)
        {
            if (grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("请先选中数据行", "提示");
                return;
            }

            var dlg = new Form
            {
                Text = "选择新类型", Size = new Size(220, 140),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false
            };
            var cb = new ComboBox
            {
                Left = 20, Top = 20, Width = 160,
                DropDownStyle = ComboBoxStyle.DropDown,
                Text = "手机"
            };
            cb.Items.AddRange(new object[] { "手机", "固话", "国际", "其他" });
            dlg.Controls.Add(cb);
            var ok = new Button { Text = "确定", Left = 50, Top = 60, Width = 80 };
            ok.Click += (s2, e2) => dlg.DialogResult = DialogResult.OK;
            dlg.Controls.Add(ok);
            dlg.AcceptButton = ok;

            if (dlg.ShowDialog() != DialogResult.OK) return;
            string newType = cb.Text;
            if (string.IsNullOrEmpty(newType)) return;

            try
            {
                foreach (DataGridViewRow row in grid.SelectedRows)
                {
                    if (row.Index < dtCache.Rows.Count)
                    {
                        long id = Convert.ToInt64(dtCache.Rows[row.Index]["id"]);
                        ExecSQL(string.Format("UPDATE numbers SET type = '{0}' WHERE id = {1}", newType.Replace("'", "''"), id));
                    }
                }
                LoadTypes();
                RefreshAll();
                MessageBox.Show(string.Format("已更新 {0} 条类型为 {1}", grid.SelectedRows.Count, newType), "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("批量修改失败: " + ex.Message, "错误");
            }
        }

        // ═══════════════════════════════════════════════
        // 批量改渠道
        // ═══════════════════════════════════════════════
        void OnBatchSetChannel(object sender, EventArgs e)
        {
            if (grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("请先选中数据行", "提示");
                return;
            }

            var dlg = new Form
            {
                Text = "选择新渠道", Size = new Size(220, 140),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false
            };
            var cb = new ComboBox
            {
                Left = 20, Top = 20, Width = 160,
                DropDownStyle = ComboBoxStyle.DropDown,
                Text = "短信"
            };
            cb.Items.AddRange(new object[] { "短信", "WS", "TG", "RCS", "iOS群发专用" });
            dlg.Controls.Add(cb);
            var ok = new Button { Text = "确定", Left = 50, Top = 60, Width = 80 };
            ok.Click += (s2, e2) => dlg.DialogResult = DialogResult.OK;
            dlg.Controls.Add(ok);
            dlg.AcceptButton = ok;

            if (dlg.ShowDialog() != DialogResult.OK) return;
            string newChannel = cb.Text;
            if (string.IsNullOrEmpty(newChannel)) return;

            try
            {
                foreach (DataGridViewRow row in grid.SelectedRows)
                {
                    if (row.Index < dtCache.Rows.Count)
                    {
                        long id = Convert.ToInt64(dtCache.Rows[row.Index]["id"]);
                        ExecSQL(string.Format("UPDATE numbers SET channel = '{0}' WHERE id = {1}", newChannel.Replace("'", "''"), id));
                    }
                }
                LoadTypes();
                RefreshAll();
                MessageBox.Show(string.Format("已更新 {0} 条渠道为 {1}", grid.SelectedRows.Count, newChannel), "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show("批量修改失败: " + ex.Message, "错误");
            }
        }

        // ═══════════════════════════════════════════════
        // 关闭
        // ═══════════════════════════════════════════════
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (db != null && db.State == ConnectionState.Open)
                db.Close();
            base.OnFormClosing(e);
        }
    }
}