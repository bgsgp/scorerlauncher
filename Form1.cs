using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;

namespace scorerlauncher;

public partial class Form1 : Form
{
    private const string KeyFile = "kei.json";
    private const string ExcelFile = "score.xlsx";
    private const string NoticeFile = "notice.json";

    private string? currentOperator;
    private bool isSeasonStopped;
    private bool isDarkMode;

    private readonly BindingList<ScoreEntryModel> scoreEntries = new() { new ScoreEntryModel() };

    private Panel? loginPanel;
    private TextBox? passwordTextBox;
    private Button? confirmButton;
    private Button? logoutButton;

    private Panel? entryPanel;
    private TextBox? itemTextBox;
    private DataGridView? scoreDataGridView;
    private Button? addButton;
    private Button? submitButton;

    private DoubleBufferedListBox? noticeListBox;
    private PictureBox? seasonOverlayPictureBox;
    private CheckBox? chkDarkMode;

    public Form1()
    {
        DoubleBuffered = true;
        SetStyle(ControlStyles.OptimizedDoubleBuffer |
                 ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.UserPaint, true);

        string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ico.ico");
        if (File.Exists(iconPath))
        {
            try { Icon = new Icon(iconPath); } catch { }
        }

        InitializeDatabaseFromKeyFile();

        SuspendLayout();
        InitializeComponents();
        ResumeLayout(false);

        AcceptButton = confirmButton;   // 登录阶段回车确认

        ApplyColorMode(IsSystemDarkMode());
        LoadNotices();
        SetEntryAreaEnabled(false);
        SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;
    }

    private void InitializeDatabaseFromKeyFile()
    {
        using var context = new ScoreContext();
        context.Database.EnsureCreated();

        if (context.Operators.Any()) return;

        string keyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, KeyFile);
        if (!File.Exists(keyPath))
        {
            MessageBox.Show($"缺少配置文件 {KeyFile}，请通过安装程序正确安装。", "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Environment.Exit(1);
        }

        var keyData = JsonConvert.DeserializeObject<KeyFileModel>(File.ReadAllText(keyPath));
        if (keyData == null)
        {
            MessageBox.Show($"{KeyFile} 格式错误，无法继续。", "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Environment.Exit(1);
        }

        // 管理员
        context.Operators.Add(new OperatorAccount
        {
            Username = "管理员",
            PasswordHash = BCrypt.Net.BCrypt.HashPassword(keyData.adminPassword),
            IsMaster = true,
            IsAdmin = true
        });

        // 总密码
        if (keyData.masterAccounts != null)
        {
            foreach (var acc in keyData.masterAccounts)
            {
                context.Operators.Add(new OperatorAccount
                {
                    Username = acc.username,
                    PasswordHash = BCrypt.Net.BCrypt.HashPassword(acc.password),
                    IsMaster = true
                });
            }
        }

        // 周密码
        if (keyData.dailyAccounts != null)
        {
            foreach (var kv in keyData.dailyAccounts)
            {
                if (Enum.TryParse<DayOfWeek>(kv.Key, true, out var day))
                {
                    context.Operators.Add(new OperatorAccount
                    {
                        Username = kv.Value.username,
                        PasswordHash = BCrypt.Net.BCrypt.HashPassword(kv.Value.password),
                        AssignedDay = day
                    });
                }
            }
        }

        context.SaveChanges();

        // 初始化完成，永久删除 kei.json
        try { File.Delete(keyPath); } catch { }
    }

    private void InitializeComponents()
    {
        Text = "乞分君";
        ClientSize = new Size(1100, 650);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        BackColor = Color.FromArgb(240, 240, 240);
        MinimumSize = new Size(1000, 600);

        // ---------- 登录区域 ----------
        loginPanel = new Panel
        {
            Location = new Point(20, 20),
            Size = new Size(420, 130),
            BackColor = Color.FromArgb(248, 249, 250),
            BorderStyle = BorderStyle.FixedSingle
        };
        Controls.Add(loginPanel);

        Label lblLogin = new Label
        {
            Text = "记分员登录",
            Font = new Font("微软雅黑", 12, FontStyle.Bold),
            Location = new Point(12, 12),
            AutoSize = true
        };
        loginPanel.Controls.Add(lblLogin);

        passwordTextBox = new TextBox
        {
            Location = new Point(12, 45),
            Size = new Size(180, 23),
            UseSystemPasswordChar = true,
            Font = new Font("微软雅黑", 10)
        };
        loginPanel.Controls.Add(passwordTextBox);

        confirmButton = new Button
        {
            Text = "确认",
            Location = new Point(200, 43),
            Size = new Size(80, 28),
            BackColor = Color.FromArgb(74, 158, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("微软雅黑", 9, FontStyle.Bold)
        };
        confirmButton.FlatAppearance.BorderSize = 0;
        confirmButton.Click += BtnConfirm_Click;
        loginPanel.Controls.Add(confirmButton);

        logoutButton = new Button
        {
            Text = "退出",
            Location = new Point(290, 43),
            Size = new Size(80, 28),
            BackColor = Color.LightGray,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("微软雅黑", 9, FontStyle.Bold),
            Enabled = false
        };
        logoutButton.FlatAppearance.BorderSize = 0;
        logoutButton.Click += BtnLogout_Click;
        loginPanel.Controls.Add(logoutButton);

        // ---------- 录入区域 ----------
        entryPanel = new Panel
        {
            Location = new Point(20, 170),
            Size = new Size(680, 420),
            BackColor = Color.FromArgb(248, 249, 250),
            BorderStyle = BorderStyle.FixedSingle
        };
        Controls.Add(entryPanel);

        Label lblEntry = new Label
        {
            Text = "积分录入",
            Font = new Font("微软雅黑", 14, FontStyle.Bold),
            ForeColor = Color.FromArgb(74, 158, 255),
            Location = new Point(12, 12),
            AutoSize = true
        };
        entryPanel.Controls.Add(lblEntry);

        Label lblItem = new Label
        {
            Text = "项目：",
            Location = new Point(12, 50),
            AutoSize = true,
            Font = new Font("微软雅黑", 10)
        };
        entryPanel.Controls.Add(lblItem);

        itemTextBox = new TextBox
        {
            Location = new Point(60, 48),
            Size = new Size(220, 23),
            Font = new Font("微软雅黑", 10)
        };
        entryPanel.Controls.Add(itemTextBox);

        // DataGridView
        scoreDataGridView = new DataGridView
        {
            Location = new Point(12, 85),
            Size = new Size(650, 260),
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            RowHeadersVisible = false,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            Font = new Font("微软雅黑", 10),
            EditMode = DataGridViewEditMode.EditOnEnter
        };
        scoreDataGridView.EditingControlShowing += ScoreDataGridView_EditingControlShowing;
        entryPanel.Controls.Add(scoreDataGridView);

        var nameCol = new DataGridViewTextBoxColumn { HeaderText = "姓名", DataPropertyName = "Name", Width = 160 };
        var groupCol = new DataGridViewComboBoxColumn
        {
            HeaderText = "组别",
            DataPropertyName = "Group",
            DataSource = Enumerable.Range(1, 12).Select(i => i.ToString()).ToList(),
            Width = 90,
            FlatStyle = FlatStyle.Flat
        };
        var scoreCol = new DataGridViewTextBoxColumn { HeaderText = "分数", DataPropertyName = "Score", Width = 120 };
        scoreDataGridView.Columns.Add(nameCol);
        scoreDataGridView.Columns.Add(groupCol);
        scoreDataGridView.Columns.Add(scoreCol);
        scoreDataGridView.DataSource = scoreEntries;

        addButton = new Button
        {
            Text = "+ 增加",
            Location = new Point(12, 360),
            Size = new Size(90, 30),
            BackColor = Color.FromArgb(74, 158, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("微软雅黑", 9, FontStyle.Bold)
        };
        addButton.FlatAppearance.BorderSize = 0;
        addButton.Click += BtnAdd_Click;
        entryPanel.Controls.Add(addButton);

        submitButton = new Button
        {
            Text = "写入 Excel",
            Location = new Point(115, 360),
            Size = new Size(110, 30),
            BackColor = Color.FromArgb(74, 158, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("微软雅黑", 9, FontStyle.Bold)
        };
        submitButton.FlatAppearance.BorderSize = 0;
        submitButton.Click += BtnSubmit_Click;
        entryPanel.Controls.Add(submitButton);

        // ---------- 右侧通知区域 ----------
        Panel rightPanel = new Panel
        {
            Location = new Point(720, 20),
            Size = new Size(350, 570),
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };
        Controls.Add(rightPanel);

        noticeListBox = new DoubleBufferedListBox
        {
            Location = new Point(5, 5),
            Size = new Size(rightPanel.ClientSize.Width - 10, rightPanel.ClientSize.Height - 10),
            BorderStyle = BorderStyle.None,
            BackColor = Color.White,
            Font = new Font("微软雅黑", 9),
            DrawMode = DrawMode.OwnerDrawVariable,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };
        noticeListBox.DrawItem += NoticeListBox_DrawItem;
        noticeListBox.MeasureItem += NoticeListBox_MeasureItem;
        rightPanel.Controls.Add(noticeListBox);

        seasonOverlayPictureBox = new PictureBox
        {
            Location = new Point(5, 5),
            Size = noticeListBox.Size,
            SizeMode = PictureBoxSizeMode.Zoom,
            Visible = false,
            BackColor = Color.White,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };
        rightPanel.Controls.Add(seasonOverlayPictureBox);

        rightPanel.Resize += (s, e) =>
        {
            if (noticeListBox != null)
                noticeListBox.Size = new Size(rightPanel.ClientSize.Width - 10, rightPanel.ClientSize.Height - 10);
            if (seasonOverlayPictureBox != null)
                seasonOverlayPictureBox.Size = noticeListBox?.Size ?? new Size(340, 560);
        };

        chkDarkMode = new CheckBox
        {
            Text = "暗夜模式",
            Location = new Point(950, 600),
            AutoSize = true,
            BackColor = Color.Transparent
        };
        chkDarkMode.CheckedChanged += (s, e) =>
        {
            if (chkDarkMode.Checked != isDarkMode)
            {
                ApplyColorMode(chkDarkMode.Checked);
            }
        };
        Controls.Add(chkDarkMode);
    }

    private void ScoreDataGridView_EditingControlShowing(object? sender, DataGridViewEditingControlShowingEventArgs e)
    {
        if (scoreDataGridView?.CurrentCell is DataGridViewComboBoxCell && e.Control is ComboBox cb)
        {
            cb.FlatStyle = FlatStyle.Flat;
            typeof(Control).InvokeMember("DoubleBuffered",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
                null, cb, new object[] { true });
            if (isDarkMode)
            {
                cb.BackColor = Color.FromArgb(45, 45, 45);
                cb.ForeColor = Color.White;
            }
            else
            {
                cb.BackColor = Color.White;
                cb.ForeColor = Color.Black;
            }
        }
    }

    private static bool IsSystemDarkMode()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            return key?.GetValue("AppsUseLightTheme") is int val && val == 0;
        }
        catch { return false; }
    }

    private void OnUserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
    {
        if (e.Category == UserPreferenceCategory.General)
            BeginInvoke(() => ApplyColorMode(IsSystemDarkMode()));
    }

    private void ApplyColorMode(bool dark)
    {
        isDarkMode = dark;
        SuspendLayout();

        var bgColor = dark ? Color.FromArgb(30, 30, 30) : Color.FromArgb(240, 240, 240);
        var panelBg = dark ? Color.FromArgb(45, 45, 45) : Color.FromArgb(248, 249, 250);
        var textColor = dark ? Color.White : Color.Black;

        BackColor = bgColor;
        if (loginPanel != null) loginPanel.BackColor = panelBg;
        if (entryPanel != null) entryPanel.BackColor = panelBg;

        if (scoreDataGridView != null)
        {
            scoreDataGridView.BackgroundColor = dark ? Color.FromArgb(30, 30, 30) : Color.White;
            scoreDataGridView.DefaultCellStyle.BackColor = dark ? Color.FromArgb(45, 45, 45) : Color.White;
            scoreDataGridView.DefaultCellStyle.ForeColor = textColor;
            scoreDataGridView.ColumnHeadersDefaultCellStyle.BackColor = dark ? Color.FromArgb(60, 60, 60) : SystemColors.Control;
            scoreDataGridView.ColumnHeadersDefaultCellStyle.ForeColor = textColor;
            scoreDataGridView.EnableHeadersVisualStyles = false;
            scoreDataGridView.GridColor = dark ? Color.Gray : Color.LightGray;
            if (scoreDataGridView.Columns["Group"] is DataGridViewComboBoxColumn groupCol)
            {
                groupCol.FlatStyle = FlatStyle.Flat;
                groupCol.DefaultCellStyle.BackColor = dark ? Color.FromArgb(45, 45, 45) : Color.White;
                groupCol.DefaultCellStyle.ForeColor = textColor;
            }
            typeof(Control).InvokeMember("DoubleBuffered",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
                null, scoreDataGridView, new object[] { true });
        }

        if (noticeListBox != null)
        {
            noticeListBox.BackColor = dark ? Color.FromArgb(30, 30, 30) : Color.White;
            noticeListBox.ForeColor = textColor;
        }
        if (seasonOverlayPictureBox != null) seasonOverlayPictureBox.BackColor = dark ? Color.FromArgb(30, 30, 30) : Color.White;
        if (chkDarkMode != null)
        {
            chkDarkMode.ForeColor = textColor;
            if (chkDarkMode.Checked != isDarkMode)
            {
                chkDarkMode.CheckedChanged -= null; 
                chkDarkMode.Checked = isDarkMode;
                chkDarkMode.CheckedChanged += (s, e) =>
                {
                    if (chkDarkMode.Checked != isDarkMode)
                        ApplyColorMode(chkDarkMode.Checked);
                };
            }
        }

        ResumeLayout(false);
        noticeListBox?.Invalidate();
        scoreDataGridView?.Invalidate();
    }

    private void LoadNotices()
    {
        string noticePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, NoticeFile);
        if (!File.Exists(noticePath)) return;

        try
        {
            string json = File.ReadAllText(noticePath);
            var data = JsonConvert.DeserializeObject<NoticeDataWrapper>(json);
            if (data == null) return;

            isSeasonStopped = data.IsSeasonStop;

            if (isSeasonStopped)
            {
                var stop = data.Notices.FirstOrDefault();
                if (stop?.Image != null)
                {
                    string imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, stop.Image);
                    if (File.Exists(imgPath) && seasonOverlayPictureBox != null)
                        seasonOverlayPictureBox.Image = Image.FromFile(imgPath);
                }
                if (seasonOverlayPictureBox != null) seasonOverlayPictureBox.Visible = true;
                if (noticeListBox != null) noticeListBox.Visible = false;
                SetEntryAreaEnabled(false);
            }
            else
            {
                if (seasonOverlayPictureBox != null) seasonOverlayPictureBox.Visible = false;
                if (noticeListBox != null) noticeListBox.Visible = true;

                var items = data.Notices.Select(n => new NoticeItem
                {
                    Title = n.Title,
                    Subtitle = n.Subtitle,
                    Content = n.Content,
                    TitleFontSize = n.TitleFontSize,
                    SubtitleFontSize = n.SubtitleFontSize,
                    ContentFontSize = n.ContentFontSize,
                    TitleColor = n.TitleColor,
                    SubtitleColor = n.SubtitleColor,
                    ContentColor = n.ContentColor,
                    Icon = n.Type == NoticeType.Warning ? SystemIcons.Warning.ToBitmap() : SystemIcons.Information.ToBitmap()
                }).ToList();
                if (noticeListBox != null)
                {
                    noticeListBox.DataSource = null;
                    noticeListBox.DataSource = items;
                }
            }
        }
        catch (Exception ex) { MessageBox.Show($"加载通知失败：{ex.Message}"); }
    }

    private static Brush AdjustBrushForDarkMode(Brush original, bool isDarkMode)
    {
        if (!isDarkMode) return original;
        if (original is SolidBrush sb)
        {
            Color c = sb.Color;
            float brightness = (c.R * 0.299f + c.G * 0.587f + c.B * 0.114f) / 255f;
            if (brightness < 0.25f) return Brushes.White;
        }
        return original;
    }

    private void NoticeListBox_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || noticeListBox?.Items[e.Index] is not NoticeItem item) return;
        e.DrawBackground();
        var g = e.Graphics;
        var bounds = e.Bounds;
        bounds.Inflate(-4, -4);

        using var backBrush = new SolidBrush(isDarkMode ? Color.FromArgb(45, 45, 45) : Color.FromArgb(248, 249, 250));
        g.FillRectangle(backBrush, bounds);
        using var pen = new Pen(isDarkMode ? Color.Gray : Color.LightGray);
        g.DrawRectangle(pen, bounds);

        if (item.Icon != null)
            g.DrawImage(item.Icon, bounds.X + 8, bounds.Y + 8, 24, 24);

        Brush titleBrush = AdjustBrushForDarkMode(item.TitleColorBrush, isDarkMode);
        Brush subtitleBrush = AdjustBrushForDarkMode(item.SubtitleColorBrush, isDarkMode);
        Brush contentBrush = AdjustBrushForDarkMode(item.ContentColorBrush, isDarkMode);

        using var titleFont = new Font("微软雅黑", item.TitleFontSize, FontStyle.Bold);
        g.DrawString(item.Title, titleFont, titleBrush, bounds.X + 40, bounds.Y + 6);

        int subtitleHeight = 0;
        if (!string.IsNullOrEmpty(item.Subtitle))
        {
            using var subFont = new Font("微软雅黑", item.SubtitleFontSize);
            var subRect = new Rectangle(bounds.X + 40, bounds.Y + 6 + (int)(item.TitleFontSize * 1.5),
                                        bounds.Width - 48, bounds.Height - 30);
            g.DrawString(item.Subtitle, subFont, subtitleBrush, subRect);
            subtitleHeight = (int)g.MeasureString(item.Subtitle, subFont, subRect.Width).Height;
        }

        if (!string.IsNullOrEmpty(item.Content))
        {
            int contentY = bounds.Y + 6 + (int)(item.TitleFontSize * 1.5) + subtitleHeight + 4;
            var contentRect = new Rectangle(bounds.X + 8, contentY, bounds.Width - 16,
                                            bounds.Height - (contentY - bounds.Y) - 8);
            using var contentFont = new Font("微软雅黑", item.ContentFontSize);
            g.DrawString(item.Content, contentFont, contentBrush, contentRect);
        }
    }

    private void NoticeListBox_MeasureItem(object? sender, MeasureItemEventArgs e)
    {
        if (noticeListBox?.Items[e.Index] is not NoticeItem item)
        {
            e.ItemHeight = 90;
            return;
        }

        using var g = noticeListBox.CreateGraphics();
        int availableWidth = noticeListBox.ClientSize.Width - 30;

        int titleHeight = (int)(item.TitleFontSize * 1.8f);
        int subtitleHeight = 0;
        if (!string.IsNullOrEmpty(item.Subtitle))
        {
            using var subFont = new Font("微软雅黑", item.SubtitleFontSize);
            var subSize = g.MeasureString(item.Subtitle, subFont, availableWidth - 48);
            subtitleHeight = (int)subSize.Height + 4;
        }
        int contentHeight = 0;
        if (!string.IsNullOrEmpty(item.Content))
        {
            using var contentFont = new Font("微软雅黑", item.ContentFontSize);
            var contentSize = g.MeasureString(item.Content, contentFont, availableWidth);
            contentHeight = (int)contentSize.Height + 8;
        }

        e.ItemHeight = 20 + titleHeight + subtitleHeight + contentHeight + 15;
        if (e.ItemHeight < 90) e.ItemHeight = 90;
    }

    private void SetEntryAreaEnabled(bool enabled)
    {
        if (entryPanel != null) entryPanel.Enabled = enabled && !isSeasonStopped;
    }

    private void BtnConfirm_Click(object? sender, EventArgs e)
    {
        if (passwordTextBox == null) return;
        string pwd = passwordTextBox.Text.Trim();

        using var context = new ScoreContext();
        var today = DateTime.Today.DayOfWeek;

        var account = context.Operators
            .AsEnumerable()
            .FirstOrDefault(a => a.AssignedDay == today && BCrypt.Net.BCrypt.Verify(pwd, a.PasswordHash));

        account ??= context.Operators
            .AsEnumerable()
            .FirstOrDefault(a => a.IsMaster && BCrypt.Net.BCrypt.Verify(pwd, a.PasswordHash));

        if (account != null)
        {
            currentOperator = account.Username;
            MessageBox.Show($"欢迎 {account.Username} ，登录成功！", "登录成功");
            passwordTextBox.Enabled = false;
            if (confirmButton != null) confirmButton.Enabled = false;
            if (logoutButton != null) logoutButton.Enabled = true;
            SetEntryAreaEnabled(true);
            AcceptButton = submitButton;   // 改为提交按钮回车
            itemTextBox?.Focus();
        }
        else
        {
            MessageBox.Show("密码错误！");
            passwordTextBox.Clear();
            passwordTextBox.Focus();
        }
    }

    private void BtnLogout_Click(object? sender, EventArgs e)
    {
        currentOperator = null;
        if (passwordTextBox != null) { passwordTextBox.Enabled = true; passwordTextBox.Clear(); }
        if (confirmButton != null) confirmButton.Enabled = true;
        if (logoutButton != null) logoutButton.Enabled = false;
        SetEntryAreaEnabled(false);
        AcceptButton = confirmButton;   // 恢复登录回车
        MessageBox.Show("已退出登录！");
    }

    private async void BtnSubmit_Click(object? sender, EventArgs e)
    {
        if (scoreDataGridView == null || itemTextBox == null || submitButton == null) return;
        scoreDataGridView.EndEdit();

        string item = itemTextBox.Text.Trim();
        if (string.IsNullOrEmpty(item)) { MessageBox.Show("项目不能为空！"); return; }

        var changes = new List<(int group, double score, string name)>();
        foreach (var entry in scoreEntries)
        {
            string name = entry.Name.Trim();
            if (string.IsNullOrEmpty(name)) continue;
            if (!int.TryParse(entry.Group, out int group) || !TryParseScore(entry.Score, out double score))
            {
                MessageBox.Show($"请检查 {name} 的填写。"); return;
            }
            changes.Add((group, score, name));
        }
        if (changes.Count == 0) { MessageBox.Show("至少一条有效记录！"); return; }

        submitButton.Enabled = false;
        submitButton.Text = "写入中……";
        try
        {
            await Task.Run(() => UpdateExcel(changes, item, currentOperator ?? "未知"));
            MessageBox.Show("写入成功！");
            itemTextBox.Clear();
            foreach (var entry in scoreEntries) { entry.Name = ""; entry.Group = "1"; entry.Score = ""; }
        }
        catch (Exception ex) { MessageBox.Show($"写入失败：{ex.Message}"); }
        finally { submitButton.Enabled = true; submitButton.Text = "写入 Excel"; }
    }

    private void BtnAdd_Click(object? sender, EventArgs e) => scoreEntries.Add(new ScoreEntryModel());

    private static bool TryParseScore(string text, out double score)
    {
        string s = text.Trim().Replace(" ", "");
        if (string.IsNullOrEmpty(s)) { score = 0; return false; }
        if (s.StartsWith('/')) s = "+" + s[1..];
        else if (!s.StartsWith('+') && !s.StartsWith('-')) s = "+" + s;
        return double.TryParse(s, out score);
    }

    private static void UpdateExcel(List<(int group, double score, string name)> changes, string item, string operatorName)
    {
        string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ExcelFile);
        using var wb = File.Exists(path) ? new XLWorkbook(path) : new XLWorkbook();
        var wsRec = wb.Worksheets.Contains("RECORD") ? wb.Worksheet("RECORD") : wb.AddWorksheet("RECORD");
        var wsLog = wb.Worksheets.Contains("LOG") ? wb.Worksheet("LOG") : wb.AddWorksheet("LOG");

        if (wsRec.LastRowUsed() == null)
        {
            wsRec.Cell(1, 1).Value = "组别"; wsRec.Cell(1, 2).Value = "积分";
            for (int i = 1; i <= 12; i++) { wsRec.Cell(i + 1, 1).Value = i; wsRec.Cell(i + 1, 2).Value = 0; }
        }

        var map = new Dictionary<int, int>();
        for (int r = 2; r <= (wsRec.LastRowUsed()?.RowNumber() ?? 1); r++)
            if (int.TryParse(wsRec.Cell(r, 1).GetValue<string>(), out int g)) map[g] = r;

        foreach (var (g, sc, _) in changes)
        {
            if (map.TryGetValue(g, out int row))
            {
                double old = wsRec.Cell(row, 2).GetValue<double>();
                wsRec.Cell(row, 2).Value = old + sc;
            }
        }

        static string Fmt(double s) => s >= 0 ? $"+{s:0.##}" : s.ToString("0.##");
        string detail = string.Join("，", changes.Select(x => $"#{x.group:00}{Fmt(x.score)}（{x.name}）")) + $"【{item}】";

        if (wsLog.LastRowUsed() == null) { wsLog.Cell(1, 1).Value = "时间戳"; wsLog.Cell(1, 2).Value = "明细"; wsLog.Cell(1, 3).Value = "操作人"; }
        int newRow = (wsLog.LastRowUsed()?.RowNumber() ?? 1) + 1;
        wsLog.Cell(newRow, 1).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        wsLog.Cell(newRow, 2).Value = detail;
        wsLog.Cell(newRow, 3).Value = operatorName;
        wb.SaveAs(path);
    }
}

// ---------- 防闪烁 ListBox ----------
public class DoubleBufferedListBox : ListBox
{
    public DoubleBufferedListBox()
    {
        DoubleBuffered = true;
    }
}

// ---------- 数据模型 ----------
public class ScoreEntryModel : INotifyPropertyChanged
{
    private string _name = "", _group = "1", _score = "";
    public string Name { get => _name; set { _name = value; OnPropertyChanged(nameof(Name)); } }
    public string Group { get => _group; set { _group = value; OnPropertyChanged(nameof(Group)); } }
    public string Score { get => _score; set { _score = value; OnPropertyChanged(nameof(Score)); } }
    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged(string prop) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
}

public enum NoticeType { Warning, Announcement }

public class NoticeData
{
    public NoticeType Type { get; set; }
    public string Title { get; set; } = "";
    public string Subtitle { get; set; } = "";
    public string Content { get; set; } = "";
    public string TitleColor { get; set; } = "";
    public float TitleFontSize { get; set; } = 14;
    public string SubtitleColor { get; set; } = "";
    public float SubtitleFontSize { get; set; } = 12;
    public string ContentColor { get; set; } = "";
    public float ContentFontSize { get; set; } = 10;
    public bool IsSeasonStop { get; set; }
    public string Image { get; set; } = "";
}

public class NoticeDataWrapper
{
    public List<NoticeData> Notices { get; set; } = new();
    public bool IsSeasonStop { get; set; }
}

public class NoticeItem
{
    public string Title { get; set; } = "";
    public string Subtitle { get; set; } = "";
    public string Content { get; set; } = "";
    public float TitleFontSize { get; set; } = 14;
    public float SubtitleFontSize { get; set; } = 12;
    public float ContentFontSize { get; set; } = 10;
    public string TitleColor { get; set; } = "";
    public string SubtitleColor { get; set; } = "";
    public string ContentColor { get; set; } = "";
    public Image? Icon { get; set; }
    public SolidBrush TitleColorBrush => GetBrush(TitleColor, Color.Black);
    public SolidBrush SubtitleColorBrush => GetBrush(SubtitleColor, Color.Gray);
    public SolidBrush ContentColorBrush => GetBrush(ContentColor, Color.DimGray);
    private static SolidBrush GetBrush(string code, Color def)
    {
        try { return string.IsNullOrEmpty(code) ? new SolidBrush(def) : new SolidBrush(ColorTranslator.FromHtml(code)); }
        catch { return new SolidBrush(def); }
    }
}

// kei.json 映射模型
public class KeyFileModel
{
    public string adminPassword { get; set; } = "";
    public List<AccountEntry> masterAccounts { get; set; } = new();
    public Dictionary<string, AccountEntry> dailyAccounts { get; set; } = new();
}

public class AccountEntry
{
    public string username { get; set; } = "";
    public string password { get; set; } = "";
}