using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace LauncherIntegrated
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new LauncherForm());
        }
    }

    public class LauncherForm : Form
    {
        // 周密码对应表
        private readonly Dictionary<DayOfWeek, string> DayPasswords = new Dictionary<DayOfWeek, string>
        {
            { DayOfWeek.Monday, "qwer1234" },
            { DayOfWeek.Tuesday, "xycloud" },
            { DayOfWeek.Wednesday, "wyq123456" },
            { DayOfWeek.Thursday, "249358" },
            { DayOfWeek.Friday, "wanghr5" },
            { DayOfWeek.Sunday, "1234asdf" }
        };

        // 周密码对应负责人
        private readonly Dictionary<string, string> DayOwners = new Dictionary<string, string>
        {
            { "qwer1234", "王佳森" },
            { "xycloud", "肖赟" },
            { "wyq123456", "王玉祺" },
            { "249358", "陈妍熙" },
            { "wanghr5", "王皓然" },
            { "gzrooster", "龚子" },
            { "1234asdf", "赵欣然" }
        };

        // 通用密码对应负责人
        private readonly Dictionary<string, string> MasterPasswords = new Dictionary<string, string>
        {
            { "gzrooster", "龚子" },
            { "39C5BB", "李艺渊" },
            { "zrj0730123", "曾睿婕" }
        };

        private const string ExcelFile = "score.xlsx";

        private TextBox txtPassword = null!;
        private Button btnConfirm = null!;

        private TextBox txtItem = null!;
        private readonly List<TextBox> txtNames = new List<TextBox>();
        private readonly List<ComboBox> cmbGroups = new List<ComboBox>();
        private readonly List<TextBox> txtScores = new List<TextBox>();
        private Button btnSubmit = null!;

        private Panel loginPanel = null!;
        private Panel inputPanel = null!;

        private string? currentOperator = null;

        public LauncherForm()
        {
            InitializeUi();
            SetInputAreaEnabled(false);
        }

        private void InitializeUi()
        {
            Text = "乞分君";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 900;
            Height = 650;
            MinimumSize = new Size(900, 650);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            BackColor = Color.Gainsboro;

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            string icoPath = Path.Combine(baseDir, "ico.ico");
            if (File.Exists(icoPath))
            {
                try
                {
                    Icon = new Icon(icoPath);
                }
                catch
                {
                }
            }

            string bgPath = Path.Combine(baseDir, "bg.png");
            if (File.Exists(bgPath))
            {
                BackgroundImage = Image.FromFile(bgPath);
                BackgroundImageLayout = ImageLayout.Stretch;
            }

            loginPanel = new Panel
            {
                Size = new Size(420, 150),
                Location = new Point(30, 30),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            Controls.Add(loginPanel);

            Label lblLoginTitle = new Label
            {
                Text = "登录验证",
                Font = new Font("微软雅黑", 15, FontStyle.Bold),
                AutoSize = true,
                BackColor = Color.Transparent,
                Location = new Point(20, 15)
            };
            loginPanel.Controls.Add(lblLoginTitle);

            Label lblPassword = new Label
            {
                Text = "密码",
                Font = new Font("微软雅黑", 11),
                AutoSize = true,
                BackColor = Color.Transparent,
                Location = new Point(20, 62)
            };
            loginPanel.Controls.Add(lblPassword);

            txtPassword = new TextBox
            {
                Width = 220,
                Font = new Font("微软雅黑", 12, FontStyle.Bold),
                UseSystemPasswordChar = true,
                Location = new Point(75, 56),
                BorderStyle = BorderStyle.FixedSingle
            };
            loginPanel.Controls.Add(txtPassword);

            btnConfirm = new Button
            {
                Text = "确定",
                Width = 90,
                Height = 34,
                Location = new Point(310, 54),
                Font = new Font("微软雅黑", 10, FontStyle.Bold)
            };
            btnConfirm.Click += BtnConfirm_Click;
            loginPanel.Controls.Add(btnConfirm);

            Label lblTip = new Label
            {
                Text = "输入今日密码或通用密码后解锁录入区",
                Font = new Font("微软雅黑", 9),
                ForeColor = Color.DimGray,
                AutoSize = true,
                BackColor = Color.Transparent,
                Location = new Point(20, 108)
            };
            loginPanel.Controls.Add(lblTip);

            AcceptButton = btnConfirm;

            inputPanel = new Panel
            {
                Size = new Size(820, 380),
                Location = new Point(30, 210),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            Controls.Add(inputPanel);

            Label lblInputTitle = new Label
            {
                Text = "积分录入",
                Font = new Font("微软雅黑", 15, FontStyle.Bold),
                AutoSize = true,
                BackColor = Color.Transparent,
                Location = new Point(20, 15)
            };
            inputPanel.Controls.Add(lblInputTitle);

            Label lblItem = new Label
            {
                Text = "项目",
                Font = new Font("微软雅黑", 11),
                AutoSize = true,
                BackColor = Color.Transparent,
                Location = new Point(20, 60)
            };
            inputPanel.Controls.Add(lblItem);

            txtItem = new TextBox
            {
                Width = 620,
                Font = new Font("微软雅黑", 11),
                Location = new Point(75, 56),
                BorderStyle = BorderStyle.FixedSingle
            };
            inputPanel.Controls.Add(txtItem);

            string[] headers = { "姓名", "组别", "计分" };
            int[] xPositions = { 20, 250, 400 };

            for (int i = 0; i < headers.Length; i++)
            {
                Label lbl = new Label
                {
                    Text = headers[i],
                    Font = new Font("微软雅黑", 10, FontStyle.Bold),
                    AutoSize = true,
                    BackColor = Color.Transparent,
                    Location = new Point(xPositions[i], 105)
                };
                inputPanel.Controls.Add(lbl);
            }

            for (int i = 0; i < 6; i++)
            {
                int y = 135 + i * 34;

                TextBox nameBox = new TextBox
                {
                    Width = 180,
                    Font = new Font("微软雅黑", 10),
                    Location = new Point(20, y),
                    BorderStyle = BorderStyle.FixedSingle
                };

                ComboBox groupBox = new ComboBox
                {
                    Width = 100,
                    Font = new Font("微软雅黑", 10),
                    Location = new Point(250, y),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                groupBox.Items.AddRange(Enumerable.Range(1, 12).Select(x => x.ToString()).ToArray());

                TextBox scoreBox = new TextBox
                {
                    Width = 100,
                    Font = new Font("微软雅黑", 10),
                    Location = new Point(400, y),
                    BorderStyle = BorderStyle.FixedSingle
                };

                inputPanel.Controls.Add(nameBox);
                inputPanel.Controls.Add(groupBox);
                inputPanel.Controls.Add(scoreBox);

                txtNames.Add(nameBox);
                cmbGroups.Add(groupBox);
                txtScores.Add(scoreBox);
            }

            Label lblDesc = new Label
            {
                Text = "计分支持：+1  -1  /1  1  +0.5",
                Font = new Font("微软雅黑", 9),
                ForeColor = Color.DimGray,
                AutoSize = true,
                BackColor = Color.Transparent,
                Location = new Point(20, 345)
            };
            inputPanel.Controls.Add(lblDesc);

            btnSubmit = new Button
            {
                Text = "写入 Excel",
                Width = 120,
                Height = 38,
                Font = new Font("微软雅黑", 10, FontStyle.Bold),
                Location = new Point(670, 332)
            };
            btnSubmit.Click += BtnSubmit_Click;
            inputPanel.Controls.Add(btnSubmit);
        }

        private void SetInputAreaEnabled(bool enabled)
        {
            txtItem.Enabled = enabled;
            btnSubmit.Enabled = enabled;

            foreach (TextBox tb in txtNames)
                tb.Enabled = enabled;

            foreach (ComboBox cb in cmbGroups)
                cb.Enabled = enabled;

            foreach (TextBox tb in txtScores)
                tb.Enabled = enabled;
        }

        private void BtnConfirm_Click(object? sender, EventArgs e)
        {
            string pwd = txtPassword.Text.Trim();
            string? operatorName = GetOperator(pwd);

            if (!string.IsNullOrEmpty(operatorName))
            {
                currentOperator = operatorName;
                MessageBox.Show($"欢迎 {currentOperator}，可以开始录入了！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                txtPassword.Enabled = false;
                btnConfirm.Enabled = false;
                SetInputAreaEnabled(true);
                txtItem.Focus();
            }
            else
            {
                currentOperator = null;
                MessageBox.Show("密码不正确！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPassword.Clear();
                txtPassword.Focus();
            }
        }

        private string? GetOperator(string pwd)
        {
            DayOfWeek today = DateTime.Today.DayOfWeek;

            if (DayPasswords.TryGetValue(today, out string? todayPwd) && todayPwd == pwd)
            {
                if (DayOwners.TryGetValue(pwd, out string? owner))
                    return owner;
                return "未知";
            }

            if (MasterPasswords.TryGetValue(pwd, out string? masterOwner))
                return masterOwner;

            return null;
        }

        private void BtnSubmit_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(currentOperator))
            {
                MessageBox.Show("请先输入正确密码。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string item = txtItem.Text.Trim();
            if (string.IsNullOrEmpty(item))
            {
                MessageBox.Show("项目不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            List<(int group, double score, string name)> changes = new List<(int group, double score, string name)>();

            for (int i = 0; i < txtNames.Count; i++)
            {
                string name = txtNames[i].Text.Trim();
                string groupText = cmbGroups[i].Text.Trim();
                string scoreText = txtScores[i].Text.Trim();

                if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(groupText) && string.IsNullOrEmpty(scoreText))
                    continue;

                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(groupText) || string.IsNullOrEmpty(scoreText))
                {
                    MessageBox.Show($"第 {i + 1} 行没填完整！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!int.TryParse(groupText, out int group) || group <= 0)
                {
                    MessageBox.Show($"第 {i + 1} 行组别格式不对！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!TryParseScore(scoreText, out double score))
                {
                    MessageBox.Show($"第 {i + 1} 行计分格式不对！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                changes.Add((group, score, name));
            }

            if (changes.Count == 0)
            {
                MessageBox.Show("至少填写一条记录！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                UpdateExcel(changes, item, currentOperator);
                MessageBox.Show("写入成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                txtItem.Clear();
                for (int i = 0; i < txtNames.Count; i++)
                {
                    txtNames[i].Clear();
                    cmbGroups[i].SelectedIndex = -1;
                    txtScores[i].Clear();
                }

                txtItem.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("写入失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool TryParseScore(string text, out double score)
        {
            string s = text.Trim().Replace(" ", "");
            if (string.IsNullOrEmpty(s))
            {
                score = 0;
                return false;
            }

            if (s.StartsWith("/"))
                s = "+" + s.Substring(1);
            else if (!s.StartsWith("+") && !s.StartsWith("-"))
                s = "+" + s;

            return double.TryParse(s, out score);
        }

        private string FormatScore(double score)
        {
            if (Math.Abs(score % 1) < 0.000001)
            {
                int v = (int)score;
                return v >= 0 ? $"+{v}" : v.ToString();
            }

            return score >= 0 ? $"+{score:0.##}" : score.ToString("0.##");
        }

        private void UpdateExcel(List<(int group, double score, string name)> changes, string item, string operatorName)
        {
            string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ExcelFile);

            using XLWorkbook wb = File.Exists(excelPath) ? new XLWorkbook(excelPath) : new XLWorkbook();

            IXLWorksheet wsRecord = wb.Worksheets.Contains("RECORD")
                ? wb.Worksheet("RECORD")
                : wb.AddWorksheet("RECORD");

            IXLWorksheet wsLog = wb.Worksheets.Contains("LOG")
                ? wb.Worksheet("LOG")
                : wb.AddWorksheet("LOG");

            if (wsRecord.LastRowUsed() == null)
            {
                wsRecord.Cell(1, 1).Value = "组别";
                wsRecord.Cell(1, 2).Value = "积分";

                for (int i = 1; i <= 12; i++)
                {
                    wsRecord.Cell(i + 1, 1).Value = i;
                    wsRecord.Cell(i + 1, 2).Value = 0;
                }
            }

            int recordLastRow = wsRecord.LastRowUsed()?.RowNumber() ?? 1;

            Dictionary<int, int> groupRowMap = new Dictionary<int, int>();
            for (int r = 2; r <= recordLastRow; r++)
            {
                string groupText = wsRecord.Cell(r, 1).GetValue<string>();
                if (int.TryParse(groupText, out int g))
                    groupRowMap[g] = r;
            }

            foreach ((int group, double score, string name) in changes)
            {
                if (groupRowMap.TryGetValue(group, out int row))
                {
                    double oldScore = wsRecord.Cell(row, 2).GetValue<double>();
                    wsRecord.Cell(row, 2).Value = oldScore + score;
                }
                else
                {
                    recordLastRow++;
                    wsRecord.Cell(recordLastRow, 1).Value = group;
                    wsRecord.Cell(recordLastRow, 2).Value = score;
                    groupRowMap[group] = recordLastRow;
                }
            }

            string detail = string.Join("，", changes.Select(x => $"#{x.group:00}{FormatScore(x.score)}（{x.name}）")) + $"【{item}】";

            if (wsLog.LastRowUsed() == null)
            {
                wsLog.Cell(1, 1).Value = "时间戳";
                wsLog.Cell(1, 2).Value = "明细";
                wsLog.Cell(1, 3).Value = "操作人";
            }

            int logNewRow = (wsLog.LastRowUsed()?.RowNumber() ?? 1) + 1;
            wsLog.Cell(logNewRow, 1).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            wsLog.Cell(logNewRow, 2).Value = detail;
            wsLog.Cell(logNewRow, 3).Value = operatorName;

            wb.SaveAs(excelPath);
        }
    }
}