using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;

namespace scorerlauncher;

public partial class Form1 : Form
{
    private const string KeyFile = "kei.json";
    private const string ExcelFile = "score.xlsx";
    private const string NoticeFile = "notice.json";
    private Label? lblLoginTitle;   // 保存登录标题标签的引用

    // 将字段初始化为默认非空值以消除 CS0649 警告
    private string currentOperator = "未知";
    private bool isSeasonStopped;
    private bool isDarkMode;
    private string originalLoginLabelText = "记分员登录";   // 新增字段

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

    private NameValidator? _nameValidator;

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
        InitializeNameValidator();

        SuspendLayout();
        InitializeComponents();
        ResumeLayout(false);

        AcceptButton = confirmButton;
        ApplyColorMode(IsSystemDarkMode());
        LoadNotices();
        SetEntryAreaEnabled(false);
        SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;
    }

    private void InitializeNameValidator()
    {
        try
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string modelPath = Path.Combine(baseDir, "name_model.onnx");
            string vocabPath = Path.Combine(baseDir, "vocab.json");

            if (File.Exists(modelPath) && File.Exists(vocabPath))
            {
                _nameValidator = new NameValidator(modelPath, vocabPath, maxLength: 8,
                    inputName: "input", outputName: "output");
            }
            else
            {
                MessageBox.Show("姓名校验模型文件缺失，将跳过姓名校验。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"姓名校验器初始化失败：{ex.Message}，将跳过校验。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
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

        context.Operators.Add(new OperatorAccount
        {
            Username = "管理员",
            PasswordHash = BCrypt.Net.BCrypt.HashPassword(keyData.adminPassword),
            IsMaster = true,
            IsAdmin = true
        });

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

        lblLoginTitle = new Label
        {
            Text = originalLoginLabelText,
            Font = new Font("微软雅黑", 12, FontStyle.Bold),
            Location = new Point(12, 12),
            AutoSize = true
        };
        loginPanel.Controls.Add(lblLoginTitle);

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

    // 通知列表绘制（完整实现）
    private void NoticeListBox_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || noticeListBox == null) return;
        e.DrawBackground();

        var item = (NoticeItem)noticeListBox.Items[e.Index];
        Rectangle bounds = e.Bounds;
        bool isDark = isDarkMode;

        const int iconSize = 16;
        if (item.Icon != null)
        {
            e.Graphics.DrawImage(item.Icon, bounds.X + 4, bounds.Y + 4, iconSize, iconSize);
        }

        int textLeft = bounds.X + iconSize + 8;
        int topOffset = 2;

        using (Brush titleBrush = item.TitleColorBrush)
        using (Font titleFont = new Font("微软雅黑", item.TitleFontSize, FontStyle.Bold))
        {
            string title = item.Title;
            SizeF titleSize = e.Graphics.MeasureString(title, titleFont);
            e.Graphics.DrawString(title, titleFont, AdjustBrushForDarkMode(titleBrush, isDark), textLeft, bounds.Y + topOffset);
            topOffset += (int)titleSize.Height + 2;
        }

        if (!string.IsNullOrEmpty(item.Subtitle))
        {
            using (Brush subBrush = item.SubtitleColorBrush)
            using (Font subFont = new Font("微软雅黑", item.SubtitleFontSize, FontStyle.Regular))
            {
                string subtitle = item.Subtitle;
                SizeF subSize = e.Graphics.MeasureString(subtitle, subFont);
                e.Graphics.DrawString(subtitle, subFont, AdjustBrushForDarkMode(subBrush, isDark), textLeft, bounds.Y + topOffset);
                topOffset += (int)subSize.Height + 2;
            }
        }

        if (!string.IsNullOrEmpty(item.Content))
        {
            using (Brush contentBrush = item.ContentColorBrush)
            using (Font contentFont = new Font("微软雅黑", item.ContentFontSize, FontStyle.Regular))
            {
                Rectangle contentRect = new Rectangle(textLeft, bounds.Y + topOffset, bounds.Width - textLeft - 4, bounds.Height - topOffset - 4);
                e.Graphics.DrawString(item.Content, contentFont, AdjustBrushForDarkMode(contentBrush, isDark), contentRect);
            }
        }

        if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
        {
            using (Brush selectedBrush = new SolidBrush(Color.FromArgb(80, SystemColors.Highlight)))
            {
                e.Graphics.FillRectangle(selectedBrush, bounds);
            }
        }

        e.DrawFocusRectangle();
    }

    // 通知列表动态高度（完整实现）
    private void NoticeListBox_MeasureItem(object? sender, MeasureItemEventArgs e)
    {
        if (e.Index < 0 || noticeListBox == null) return;
        var item = (NoticeItem)noticeListBox.Items[e.Index];
        using (Graphics g = CreateGraphics())
        {
            float totalHeight = 4;
            using (Font titleFont = new Font("微软雅黑", item.TitleFontSize, FontStyle.Bold))
                totalHeight += g.MeasureString(item.Title ?? "", titleFont).Height + 2;
            if (!string.IsNullOrEmpty(item.Subtitle))
                using (Font subFont = new Font("微软雅黑", item.SubtitleFontSize, FontStyle.Regular))
                    totalHeight += g.MeasureString(item.Subtitle, subFont).Height + 2;
            if (!string.IsNullOrEmpty(item.Content))
                using (Font contentFont = new Font("微软雅黑", item.ContentFontSize, FontStyle.Regular))
                    totalHeight += g.MeasureString(item.Content, contentFont, e.ItemWidth - 40).Height + 2;
            e.ItemHeight = (int)Math.Max(totalHeight, 40);
        }
    }

    private void SetEntryAreaEnabled(bool enabled)
    {
        if (entryPanel != null) entryPanel.Enabled = enabled && !isSeasonStopped;
    }

    // 登录按钮事件
    private async void BtnConfirm_Click(object? sender, EventArgs e)
    {
        if (passwordTextBox == null || confirmButton == null || logoutButton == null)
            return;

        string pwd = passwordTextBox.Text.Trim();
        if (string.IsNullOrEmpty(pwd))
        {
            MessageBox.Show("请输入密码。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        confirmButton.Enabled = false;
        try
        {
            using var context = new ScoreContext();
            var today = DateTime.Now.DayOfWeek;
            var operators = await context.Operators.ToListAsync();

            OperatorAccount? matched = null;
            foreach (var op in operators)
            {
                bool isValidMaster = op.IsMaster;
                bool isValidDaily = !op.IsMaster && op.AssignedDay.HasValue && op.AssignedDay.Value == today;
                if ((isValidMaster || isValidDaily) && BCrypt.Net.BCrypt.Verify(pwd, op.PasswordHash))
                {
                    matched = op;
                    break;
                }
            }

            if (matched == null)
            {
                MessageBox.Show("密码错误或您今日无权限登录。", "登录失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 登录成功，弹出提示
            MessageBox.Show($"欢迎{matched.Username}，登录成功！", "登录成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

            currentOperator = matched.Username;
            // isSeasonStopped 已在 LoadNotices 中赋值

            passwordTextBox.Enabled = false;
            passwordTextBox.Text = "";
            confirmButton.Enabled = false;
            logoutButton.Enabled = true;

            var lblLogin = loginPanel?.Controls.OfType<Label>().FirstOrDefault(l => l.Text == originalLoginLabelText);
            if (lblLogin != null) lblLogin.Text = $"当前登录：{currentOperator}";

            SetEntryAreaEnabled(true);
            if (isSeasonStopped)
            {
                MessageBox.Show("当前赛季已停止，无法进行积分录入！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"登录过程发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (confirmButton.Enabled == false && logoutButton.Enabled == true)
                confirmButton.Enabled = false;
            else
                confirmButton.Enabled = true;
        }
    }

    // 退出按钮事件
    private void BtnLogout_Click(object? sender, EventArgs e)
    {
        if (passwordTextBox == null || confirmButton == null || logoutButton == null || entryPanel == null)
            return;

        currentOperator = "未知";
        passwordTextBox.Enabled = true;
        passwordTextBox.Clear();
        confirmButton.Enabled = true;
        logoutButton.Enabled = false;

        var lblLogin = loginPanel?.Controls.OfType<Label>().FirstOrDefault(l => l.Text.StartsWith("当前登录："));
        if (lblLogin != null) lblLogin.Text = originalLoginLabelText;

        itemTextBox?.Clear();
        scoreEntries.Clear();
        scoreEntries.Add(new ScoreEntryModel());

        SetEntryAreaEnabled(false);
    }

    // 修改后的提交按钮逻辑：支持顿号分隔，日志追加不覆盖
    private async void BtnSubmit_Click(object? sender, EventArgs e)
    {
        if (scoreDataGridView == null || itemTextBox == null || submitButton == null) return;
        scoreDataGridView.EndEdit();

        string item = itemTextBox.Text.Trim();
        if (string.IsNullOrEmpty(item)) { MessageBox.Show("项目不能为空！"); return; }

        var changes = new List<(int group, double score, string name)>();
        foreach (var entry in scoreEntries)
        {
            string name = entry.Name?.Trim() ?? "";
            if (string.IsNullOrEmpty(name)) continue;
            if (!int.TryParse(entry.Group, out int group) || !TryParseScore(entry.Score, out double score))
            {
                MessageBox.Show($"请检查 {name} 的填写。"); return;
            }
            changes.Add((group, score, name));
        }
        if (changes.Count == 0) { MessageBox.Show("至少一条有效记录！"); return; }

        // ---- AI 姓名校验，支持顿号分隔 ----
        if (_nameValidator != null)
        {
            var invalidEntries = new List<(int group, double score, string name)>();
            foreach (var change in changes)
            {
                // 按顿号拆分多个姓名
                var parts = change.name.Split('、', StringSplitOptions.RemoveEmptyEntries);
                bool allValid = true;
                foreach (var part in parts)
                {
                    string singleName = part.Trim();
                    if (string.IsNullOrEmpty(singleName)) continue;
                    float prob = _nameValidator.GetProbability(singleName);
                    System.Diagnostics.Debug.WriteLine($"[AI校验] {singleName} → 有效概率: {prob:F4}");
                    if (!_nameValidator.IsValid(singleName))
                    {
                        allValid = false;
                        break;
                    }
                }
                if (!allValid)
                    invalidEntries.Add(change);
            }

            if (invalidEntries.Any())
            {
                using var skipDialog = new SkipValidationDialog(invalidEntries);
                var result = skipDialog.ShowDialog();

                if (result == DialogResult.Cancel)
                {
                    return;
                }
                else if (result == DialogResult.OK && skipDialog.SkipApproved)
                {
                    // 日志追加（确保同一天多次跳过会累积）
                    LogSkippedValidation(invalidEntries, item, currentOperator ?? "未知");
                }
                else
                {
                    return;
                }
            }
        }

        // 写入 Excel
        submitButton.Enabled = false;
        submitButton.Text = "写入中……";
        try
        {
            await Task.Run(() => UpdateExcel(changes, item, currentOperator ?? "未知"));
            MessageBox.Show("写入成功！");
            itemTextBox.Clear();
            foreach (var entry in scoreEntries) { entry.Name = ""; entry.Group = ""; entry.Score = ""; }
        }
        catch (Exception ex) { MessageBox.Show($"写入失败：{ex.Message}"); }
        finally { submitButton.Enabled = true; submitButton.Text = "写入 Excel"; }
    }

    private void BtnAdd_Click(object? sender, EventArgs e) => scoreEntries.Add(new ScoreEntryModel());

    private static bool TryParseScore(string text, out double score)
    {
        string s = text.Trim().Replace(" ", "");
        if (s.Any(ch => (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z'))) { score = 0; return false; }
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

    // 记录跳过校验日志（追加，同一天多次调用会累加）
    private static readonly object _logLock = new();
    private void LogSkippedValidation(List<(int group, double score, string name)> invalidEntries, string item, string operatorName)
    {
        try
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string logFile = Path.Combine(desktopPath, $"SkipValidation_{DateTime.Now:yyyyMMdd}.log");
            string detail = string.Join("；", invalidEntries.Select(x => $"{x.name}（第{x.group}组，得分{x.score}）"));
            string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 操作员：{operatorName} | 项目：{item} | 跳过的无效姓名：{detail}";

            lock (_logLock)
            {
                using var sw = new StreamWriter(logFile, append: true);
                sw.WriteLine(logEntry);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"记录跳过日志失败：{ex.Message}");
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _nameValidator?.Dispose();
        base.OnFormClosing(e);
    }
}

// ---------- 防闪烁 ListBox ----------
public class DoubleBufferedListBox : ListBox
{
    public DoubleBufferedListBox() { DoubleBuffered = true; }
}

// ---------- 数据模型 ----------
public class ScoreEntryModel : INotifyPropertyChanged
{
    private string _name = "", _group = "", _score = "";
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