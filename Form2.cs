using Microsoft.EntityFrameworkCore;
using Microsoft.Win32;
using System.Drawing;

namespace scorerlauncher;

public partial class Form2 : Form
{
    private const string KeyFile = "kei.json";

    private TextBox? adminPwdTextBox;
    private Button? adminVerifyBtn;
    private Panel? topPanel;          // 新增：顶部面板引用
    private Panel? managePanel;
    private DataGridView? passwordGrid;
    private Button? saveBtn;
    private Button? addBtn;
    private Button? deleteBtn;

    private List<OperatorAccount>? accounts;
    private bool isAuthenticated = false;
    private bool isDarkMode = false;

    private static readonly Dictionary<DayOfWeek, string> WeekdayNames = new()
    {
        [DayOfWeek.Monday] = "周一",
        [DayOfWeek.Tuesday] = "周二",
        [DayOfWeek.Wednesday] = "周三",
        [DayOfWeek.Thursday] = "周四",
        [DayOfWeek.Friday] = "周五",
        [DayOfWeek.Saturday] = "周六",
        [DayOfWeek.Sunday] = "周日",
    };

    public Form2()
    {
        string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ico.ico");
        if (File.Exists(iconPath))
        {
            try { Icon = new Icon(iconPath); } catch { }
        }

        DoubleBuffered = true;
        ApplySystemTheme();

        Text = "密码更改";
        ClientSize = new Size(650, 550);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        AcceptButton = null;

        InitializeComponents();
        SystemEvents.UserPreferenceChanged += (s, e) => BeginInvoke(() => ApplySystemTheme());
    }

    private void ApplySystemTheme()
    {
        isDarkMode = IsSystemDarkMode();
        var bg = isDarkMode ? Color.FromArgb(30, 30, 30) : SystemColors.Control;
        var fg = isDarkMode ? Color.White : SystemColors.ControlText;
        BackColor = bg;
        ForeColor = fg;

        ApplyDarkModeToControls(this, isDarkMode);

        if (passwordGrid != null)
        {
            passwordGrid.BackgroundColor = isDarkMode ? Color.FromArgb(30, 30, 30) : Color.White;
            passwordGrid.DefaultCellStyle.BackColor = isDarkMode ? Color.FromArgb(45, 45, 45) : Color.White;
            passwordGrid.DefaultCellStyle.ForeColor = fg;
            passwordGrid.ColumnHeadersDefaultCellStyle.BackColor = isDarkMode ? Color.FromArgb(60, 60, 60) : SystemColors.Control;
            passwordGrid.ColumnHeadersDefaultCellStyle.ForeColor = fg;
            passwordGrid.EnableHeadersVisualStyles = false;
            passwordGrid.GridColor = isDarkMode ? Color.Gray : Color.LightGray;
            if (passwordGrid.Columns["Permission"] is DataGridViewComboBoxColumn permCol)
            {
                permCol.FlatStyle = FlatStyle.Flat;
                permCol.DefaultCellStyle.BackColor = isDarkMode ? Color.FromArgb(45, 45, 45) : Color.White;
                permCol.DefaultCellStyle.ForeColor = fg;
            }
            passwordGrid.Invalidate();
        }
    }

    private void ApplyDarkModeToControls(Control parent, bool dark)
    {
        if (parent == null) return;

        Color bg = dark ? Color.FromArgb(45, 45, 45) : Color.FromArgb(248, 249, 250);
        Color fg = dark ? Color.White : Color.Black;

        if (parent == deleteBtn)
        {
            parent.ForeColor = Color.White;
            return;
        }

        if (parent is TextBox || parent is ComboBox)
        {
            parent.BackColor = dark ? Color.FromArgb(60, 60, 60) : Color.White;
            parent.ForeColor = fg;
        }
        else if (parent is Button btn)
        {
            btn.BackColor = dark ? Color.FromArgb(74, 158, 255) : Color.FromArgb(74, 158, 255);
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
        }
        else if (parent is ListBox listBox)
        {
            listBox.BackColor = dark ? Color.FromArgb(30, 30, 30) : Color.White;
            listBox.ForeColor = fg;
        }
        else
        {
            parent.BackColor = bg;
            parent.ForeColor = fg;
        }

        foreach (Control child in parent.Controls)
        {
            ApplyDarkModeToControls(child, dark);
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

    private void InitializeComponents()
    {
        // 顶部验证面板
        topPanel = new Panel { Dock = DockStyle.Top, Height = 60, Padding = new Padding(10) };
        var lbl = new Label { Text = "管理员密码：", Location = new Point(10, 18), AutoSize = true };
        adminPwdTextBox = new TextBox { Location = new Point(100, 15), Width = 180, UseSystemPasswordChar = true };
        adminVerifyBtn = new Button
        {
            Text = "确认",
            Location = new Point(290, 13),
            Width = 80,
            BackColor = Color.FromArgb(74, 158, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        adminVerifyBtn.FlatAppearance.BorderSize = 0;
        adminVerifyBtn.Click += AdminVerifyBtn_Click;
        topPanel.Controls.Add(lbl);
        topPanel.Controls.Add(adminPwdTextBox);
        topPanel.Controls.Add(adminVerifyBtn);
        Controls.Add(topPanel);

        AcceptButton = adminVerifyBtn;

        // 管理面板（初始隐藏）
        managePanel = new Panel { Dock = DockStyle.Fill, Visible = false };

        passwordGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersHeight = 24
        };

        var permCol = new DataGridViewComboBoxColumn
        {
            HeaderText = "权限",
            DataPropertyName = "Permission",
            Width = 100,
            DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
            FlatStyle = FlatStyle.Flat,
            ValueType = typeof(string)
        };
        permCol.Items.Add("总密码");
        permCol.Items.AddRange(WeekdayNames.Values.ToArray());
        passwordGrid.Columns.Add(permCol);

        var nameCol = new DataGridViewTextBoxColumn
        {
            HeaderText = "用户名",
            DataPropertyName = "Username",
            Width = 150,
            ReadOnly = false
        };
        passwordGrid.Columns.Add(nameCol);

        var pwdCol = new DataGridViewTextBoxColumn
        {
            HeaderText = "新密码（不填视为不修改）",
            DataPropertyName = "NewPassword",
            Width = 200,
            ReadOnly = false
        };
        passwordGrid.Columns.Add(pwdCol);

        var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 40 };
        addBtn = new Button
        {
            Text = "添加",
            Location = new Point(10, 5),
            Width = 85,
            BackColor = Color.FromArgb(74, 158, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        addBtn.FlatAppearance.BorderSize = 0;
        addBtn.Click += AddBtn_Click;

        deleteBtn = new Button
        {
            Text = "删除",
            Name = "deleteBtn",
            Location = new Point(105, 5),
            Width = 85,
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        deleteBtn.FlatAppearance.BorderSize = 0;
        deleteBtn.Click += DeleteBtn_Click;

        saveBtn = new Button
        {
            Text = "保存",
            Location = new Point(200, 5),
            Width = 85,
            BackColor = Color.FromArgb(74, 158, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        saveBtn.FlatAppearance.BorderSize = 0;
        saveBtn.Click += SaveBtn_Click;

        btnPanel.Controls.Add(addBtn);
        btnPanel.Controls.Add(deleteBtn);
        btnPanel.Controls.Add(saveBtn);

        managePanel.Controls.Add(passwordGrid);
        managePanel.Controls.Add(btnPanel);
        Controls.Add(managePanel);
    }

    private void AdminVerifyBtn_Click(object? sender, EventArgs e)
    {
        if (adminPwdTextBox == null) return;
        string inputPwd = adminPwdTextBox.Text.Trim();

        using var context = new ScoreContext();
        var adminAcc = context.Operators.AsEnumerable()
            .FirstOrDefault(a => a.IsAdmin && BCrypt.Net.BCrypt.Verify(inputPwd, a.PasswordHash));

        if (adminAcc == null)
        {
            MessageBox.Show("密码错误！", "验证失败");
            return;
        }

        isAuthenticated = true;
        adminVerifyBtn!.Enabled = false;
        adminPwdTextBox.Enabled = false;

        // 隐藏顶部登录区域，只保留管理面板
        if (topPanel != null)
            topPanel.Visible = false;

        AcceptButton = saveBtn;
        LoadAccountGrid();
    }

    private void LoadAccountGrid()
    {
        using var context = new ScoreContext();
        accounts = context.Operators
            .Where(a => !a.IsAdmin)
            .OrderBy(a => a.IsMaster ? 0 : 1)
            .ThenBy(a => a.Username)
            .ToList();

        var displayList = accounts.Select(a => new PasswordEditModel
        {
            Permission = a.IsMaster ? "总密码" :
                         (a.AssignedDay != null && WeekdayNames.TryGetValue(a.AssignedDay.Value, out var dayName) ? dayName : "未知"),
            Username = a.Username ?? "",
            NewPassword = "",
            OriginalUsername = a.Username ?? ""
        }).ToList();

        passwordGrid!.DataSource = displayList;
        managePanel!.Visible = true;

        passwordGrid.PerformLayout();
        if (passwordGrid.Rows.Count > 0)
            passwordGrid.FirstDisplayedScrollingRowIndex = 0;

        ApplySystemTheme();
    }

    private void AddBtn_Click(object? sender, EventArgs e)
    {
        if (!isAuthenticated) return;

        using var addForm = new AddAccountForm();
        if (addForm.ShowDialog() != DialogResult.OK) return;

        using var context = new ScoreContext();
        if (context.Operators.Any(a => a.Username == addForm.Username))
        {
            MessageBox.Show("用户名已存在。", "错误");
            return;
        }

        var newAcc = new OperatorAccount
        {
            Username = addForm.Username,
            PasswordHash = BCrypt.Net.BCrypt.HashPassword(addForm.Password),
            IsMaster = addForm.IsMaster,
            AssignedDay = addForm.AssignedDay
        };
        context.Operators.Add(newAcc);
        context.SaveChanges();

        LoadAccountGrid();
    }

    private void DeleteBtn_Click(object? sender, EventArgs e)
    {
        if (!isAuthenticated || passwordGrid == null || passwordGrid.SelectedRows.Count == 0) return;

        var confirm = MessageBox.Show("确定要删除该账户吗？", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (confirm != DialogResult.Yes) return;

        var selectedUsernames = passwordGrid.SelectedRows
            .Cast<DataGridViewRow>()
            .Select(row => (row.DataBoundItem as PasswordEditModel)?.Username ?? "")
            .Where(u => !string.IsNullOrWhiteSpace(u));

        using var context = new ScoreContext();
        foreach (var uname in selectedUsernames)
        {
            var acc = context.Operators.FirstOrDefault(a => a.Username == uname && !a.IsAdmin);
            if (acc != null) context.Operators.Remove(acc);
        }
        context.SaveChanges();

        LoadAccountGrid();
    }

    private void SaveBtn_Click(object? sender, EventArgs e)
    {
        if (!isAuthenticated || accounts == null || passwordGrid == null) return;
        passwordGrid.EndEdit();

        var editModels = passwordGrid.DataSource as List<PasswordEditModel>;
        if (editModels == null) return;

        using var context = new ScoreContext();
        bool changed = false;

        foreach (var em in editModels)
        {
            if (string.IsNullOrWhiteSpace(em.Username)) continue;

            var dbAcc = context.Operators.FirstOrDefault(a => a.Username == em.OriginalUsername && !a.IsAdmin);
            if (dbAcc == null) continue;

            if (dbAcc.Username != em.Username.Trim())
            {
                if (context.Operators.Any(a => a.Username == em.Username.Trim() && a.Id != dbAcc.Id))
                {
                    MessageBox.Show($"用户名 {em.Username} 已被占用，保存中止。", "冲突");
                    return;
                }
                dbAcc.Username = em.Username.Trim();
                changed = true;
            }

            if (!string.IsNullOrWhiteSpace(em.NewPassword))
            {
                dbAcc.PasswordHash = BCrypt.Net.BCrypt.HashPassword(em.NewPassword);
                changed = true;
            }

            if (em.Permission == "总密码")
            {
                if (!dbAcc.IsMaster || dbAcc.AssignedDay != null)
                {
                    dbAcc.IsMaster = true;
                    dbAcc.AssignedDay = null;
                    changed = true;
                }
            }
            else
            {
                var newDay = ParseWeekDay(em.Permission);
                if (newDay != null && (dbAcc.IsMaster || dbAcc.AssignedDay != newDay))
                {
                    dbAcc.IsMaster = false;
                    dbAcc.AssignedDay = newDay;
                    changed = true;
                }
            }
        }

        if (changed)
        {
            context.SaveChanges();
            MessageBox.Show("修改已保存！", "完成");
        }

        LoadAccountGrid();
    }

    private static DayOfWeek? ParseWeekDay(string chineseDay)
    {
        foreach (var kv in WeekdayNames)
        {
            if (kv.Value == chineseDay)
                return kv.Key;
        }
        return null;
    }
}

public class PasswordEditModel
{
    public string Permission { get; set; } = "总密码";
    public string Username { get; set; } = "";
    public string NewPassword { get; set; } = "";
    public string OriginalUsername { get; set; } = "";
}

public class AddAccountForm : Form
{
    public string Username { get; private set; } = "";
    public string Password { get; private set; } = "";
    public bool IsMaster { get; private set; } = true;
    public DayOfWeek? AssignedDay { get; private set; } = null;

    public AddAccountForm()
    {
        Text = "添加账号";
        Width = 350;
        Height = 250;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var lblPermission = new Label { Text = "权限：", Location = new Point(20, 20), AutoSize = true };
        var cmbPermission = new ComboBox
        {
            Location = new Point(80, 18),
            Width = 180,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        cmbPermission.Items.Add("总密码");
        cmbPermission.Items.AddRange(new[] { "周一", "周二", "周三", "周四", "周五", "周六", "周日" });
        cmbPermission.SelectedIndex = 0;

        var lblUser = new Label { Text = "用户名：", Location = new Point(20, 60), AutoSize = true };
        var txtUser = new TextBox { Location = new Point(80, 58), Width = 180 };

        var lblPwd = new Label { Text = "密码：", Location = new Point(20, 100), AutoSize = true };
        var txtPwd = new TextBox { Location = new Point(80, 98), Width = 180, UseSystemPasswordChar = true };

        var btnOk = new Button
        {
            Text = "确定",
            Location = new Point(80, 150),
            Width = 75,
            BackColor = Color.FromArgb(74, 158, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        btnOk.FlatAppearance.BorderSize = 0;
        btnOk.Click += (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(txtUser.Text))
            {
                MessageBox.Show("请输入用户名");
                return;
            }
            if (string.IsNullOrWhiteSpace(txtPwd.Text))
            {
                MessageBox.Show("请输入密码");
                return;
            }

            Username = txtUser.Text.Trim();
            Password = txtPwd.Text.Trim();

            string perm = cmbPermission.SelectedItem?.ToString() ?? "总密码";
            if (perm == "总密码")
            {
                IsMaster = true;
                AssignedDay = null;
            }
            else
            {
                IsMaster = false;
                AssignedDay = perm switch
                {
                    "周一" => DayOfWeek.Monday,
                    "周二" => DayOfWeek.Tuesday,
                    "周三" => DayOfWeek.Wednesday,
                    "周四" => DayOfWeek.Thursday,
                    "周五" => DayOfWeek.Friday,
                    "周六" => DayOfWeek.Saturday,
                    "周日" => DayOfWeek.Sunday,
                    _ => (DayOfWeek?)null
                };
            }

            DialogResult = DialogResult.OK;
            Close();
        };

        var btnCancel = new Button
        {
            Text = "取消",
            Location = new Point(170, 150),
            Width = 75,
            FlatStyle = FlatStyle.Flat
        };
        btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

        Controls.Add(lblPermission);
        Controls.Add(cmbPermission);
        Controls.Add(lblUser);
        Controls.Add(txtUser);
        Controls.Add(lblPwd);
        Controls.Add(txtPwd);
        Controls.Add(btnOk);
        Controls.Add(btnCancel);

        AcceptButton = btnOk;
    }
}