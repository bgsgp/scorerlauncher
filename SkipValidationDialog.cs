using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace scorerlauncher
{
    public class SkipValidationDialog : Form
    {
        private readonly List<(int group, double score, string name)> _invalidEntries;
        private int _clickCount = 0;
        private Label? lblMessage;
        private Button? btnCancel;
        private Timer? resetTimer;

        public bool SkipApproved { get; private set; } = false;

        public SkipValidationDialog(List<(int group, double score, string name)> invalidEntries)
        {
            _invalidEntries = invalidEntries;
            InitializeComponent();
            LoadInvalidList();
            // 整个窗体作为隐藏的点击区域
            this.Click += Form_Click;
        }

        private void InitializeComponent()
        {
            this.Text = "姓名校验失败";
            this.Size = new Size(450, 300);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            // 提示标签
            lblMessage = new Label
            {
                Location = new Point(15, 15),
                Size = new Size(410, 50),
                Text = "以下姓名未通过校验（可能是占位符或无效姓名）：",
                Font = new Font("微软雅黑", 9, FontStyle.Bold),
                ForeColor = Color.DarkRed
            };

            // 无效姓名列表
            var listBox = new ListBox
            {
                Location = new Point(15, 70),
                Size = new Size(410, 120),
                Font = new Font("微软雅黑", 9),
                IntegralHeight = false
            };
            for (int i = 0; i < _invalidEntries.Count; i++)
            {
                var e = _invalidEntries[i];
                listBox.Items.Add($"{i + 1}. {e.name}（组{e.group}，得分{e.score}）");
            }

            // 重置计时器：5秒无点击则重置计数
            resetTimer = new Timer { Interval = 5000 };
            resetTimer.Tick += (s, e) =>
            {
                _clickCount = 0;
                resetTimer.Stop();
            };

            // 取消按钮
            btnCancel = new Button
            {
                Text = "取消写入",
                Location = new Point(350, 200),
                Size = new Size(75, 40),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;

            Controls.Add(lblMessage);
            Controls.Add(listBox);
            Controls.Add(btnCancel);

            this.CancelButton = btnCancel;
        }

        private void LoadInvalidList()
        {
            // 列表已在构造函数中添加
        }

        private void Form_Click(object? sender, EventArgs e)
        {
            resetTimer?.Stop();
            _clickCount++;

            if (_clickCount >= 8)
            {
                SkipApproved = true;
                MessageBox.Show("本次姓名校验已跳过，数据将继续写入。", "跳过成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                resetTimer?.Start();
            }
        }
    }
}