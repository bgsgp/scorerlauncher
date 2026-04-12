using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClosedXML.Excel;
using Newtonsoft.Json;
using Drawing = System.Drawing;

namespace scorerlauncher
{
    public partial class MainWindow : Window
    {
        // ==========================================
        // 数据模型（用于通知栏绑定）
        // ==========================================
        private class NoticeItem
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
            public ImageSource? Icon { get; set; }

            public SolidColorBrush TitleColorBrush => GetBrush(TitleColor, Colors.Black);
            public SolidColorBrush SubtitleColorBrush => GetBrush(SubtitleColor, Colors.Gray);
            public SolidColorBrush ContentColorBrush => GetBrush(ContentColor, Colors.DimGray);

            private static SolidColorBrush GetBrush(string colorCode, Color defaultColor)
            {
                try
                {
                    if (string.IsNullOrEmpty(colorCode)) return new SolidColorBrush(defaultColor);
                    var color = (Color)ColorConverter.ConvertFromString(colorCode);
                    return new SolidColorBrush(color);
                }
                catch { return new SolidColorBrush(defaultColor); }
            }
        }

        // 数据模型（用于积分录入绑定）
        private class ScoreEntryModel : INotifyPropertyChanged
        {
            private string _name = "";
            private string _group = "1";
            private string _score = "";

            public string Name { get => _name; set { _name = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Name))); } }
            public string Group { get => _group; set { _group = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Group))); } }
            public string Score { get => _score; set { _score = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Score))); } }
            public List<string> Groups { get; } = Enumerable.Range(1, 12).Select(x => x.ToString()).ToList();

            public event PropertyChangedEventHandler? PropertyChanged;
        }

        // ==========================================
        // 全局变量
        // ==========================================
        private static readonly Dictionary<DayOfWeek, string> DayPasswords = new()
        {
            { DayOfWeek.Monday, "qwer1234" },
            { DayOfWeek.Tuesday, "xycloud" },
            { DayOfWeek.Wednesday, "wyq123456" },
            { DayOfWeek.Thursday, "249358" },
            { DayOfWeek.Friday, "wanghr5" },
            { DayOfWeek.Sunday, "1234asdf" }
        };

        private static readonly Dictionary<string, string> DayOwners = new()
        {
            { "qwer1234", "王佳森" },
            { "xycloud", "肖赟" },
            { "wyq123456", "王玉祺" },
            { "249358", "陈妍熙" },
            { "wanghr5", "王皓然" },
            { "gzrooster", "龚子" },
            { "1234asdf", "赵欣然" }
        };

        private static readonly Dictionary<string, string> MasterPasswords = new()
        {
            { "gzrooster", "龚子" },
            { "39C5BB", "李艺渊" },
            { "zrj0730123", "曾睿婕" },
            { "ztyacjy66", "赵天语" }
        };

        private const string ExcelFile = "score.xlsx";
        private const string NoticeFile = "notice.json";
        private string? currentOperator;
        private bool isDarkMode;
        private bool isSeasonStopped;

        // 控件引用
        private readonly Border? loginArea;
        private readonly Border? entryArea;

        // 使用 ObservableCollection 避免频繁重置 ItemsSource
        private readonly ObservableCollection<ScoreEntryModel> _scoreEntries = new() { new ScoreEntryModel() };

        // 图标缓存
        private static BitmapImage? _cachedWarningIcon;
        private static BitmapImage? _cachedInfoIcon;

        // ==========================================
        // 构造函数 & 初始化
        // ==========================================
        public MainWindow()
        {
            InitializeComponent();

            loginArea = FindName("LoginArea") as Border;
            entryArea = FindName("EntryArea") as Border;

            // 绑定 DataGrid
            ScoreEntries.ItemsSource = _scoreEntries;

            // 加载通知（含覆盖图和图标）
            LoadNotices();

            // 设置初始可用状态
            SetEntryAreaEnabled(false);
            if (loginArea is not null) loginArea.IsEnabled = true;

            // 初始暗色模式为浅色
            isDarkMode = false;
            chkDarkMode.IsChecked = false;
            ApplyDarkMode(false);
        }

        // ==========================================
        // 通知加载（赛季结束显示覆盖图，正常显示通知列表+图标）
        // ==========================================
        private void LoadNotices()
        {
            string noticePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, NoticeFile);
            if (!File.Exists(noticePath))
            {
                MessageBox.Show("通知文件不存在，请检查路径。", "错误");
                return;
            }

            try
            {
                string json = File.ReadAllText(noticePath);
                var data = JsonConvert.DeserializeObject<NoticeDataWrapper>(json);
                if (data is null)
                {
                    MessageBox.Show("通知文件格式不正确。", "错误");
                    return;
                }

                isSeasonStopped = data.IsSeasonStop;

                if (isSeasonStopped)
                {
                    // 赛季结束：显示覆盖图，隐藏通知列表，禁用录入区
                    var seasonStopNotice = data.Notices.FirstOrDefault();
                    if (seasonStopNotice is not null && !string.IsNullOrEmpty(seasonStopNotice.Image))
                    {
                        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, seasonStopNotice.Image);
                        if (File.Exists(imagePath))
                        {
                            OverlayImage.Source = LoadBitmapImage(imagePath);
                        }
                    }
                    OverlayImage.Visibility = Visibility.Visible;
                    NoticeList.Visibility = Visibility.Collapsed;

                    SetEntryAreaEnabled(false);
                    if (loginArea is not null) loginArea.IsEnabled = true;
                }
                else
                {
                    // 正常模式：隐藏覆盖图，显示通知列表
                    OverlayImage.Visibility = Visibility.Collapsed;
                    NoticeList.Visibility = Visibility.Visible;

                    // 构建通知列表（带图标缓存）
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
                        Icon = n.Type == NoticeType.Warning
                            ? GetCachedIcon(Drawing.SystemIcons.Warning, ref _cachedWarningIcon)
                            : GetCachedIcon(Drawing.SystemIcons.Information, ref _cachedInfoIcon)
                    }).ToList();

                    NoticeList.ItemsSource = items;
                    SetEntryAreaEnabled(false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载通知失败：{ex.Message}", "错误");
            }
        }

        private static BitmapImage LoadBitmapImage(string path)
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.UriSource = new Uri(path, UriKind.Absolute);
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.EndInit();
            bitmap.Freeze();
            return bitmap;
        }

        private static BitmapImage GetCachedIcon(Drawing.Icon icon, ref BitmapImage? cache)
        {
            if (cache is not null) return cache;
            using var bitmap = icon.ToBitmap();
            using var ms = new MemoryStream();
            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            ms.Seek(0, SeekOrigin.Begin);
            var img = new BitmapImage();
            img.BeginInit();
            img.StreamSource = ms;
            img.CacheOption = BitmapCacheOption.OnLoad;
            img.EndInit();
            img.Freeze();
            cache = img;
            return img;
        }

        // ==========================================
        // 暗色模式
        // ==========================================
        private void ChkDarkMode_Changed(object sender, RoutedEventArgs e)
        {
            isDarkMode = chkDarkMode.IsChecked == true;
            ApplyDarkMode(isDarkMode);
        }

        private void ApplyDarkMode(bool dark)
        {
            // 修改窗口背景（不透明）
            this.Background = dark ? new SolidColorBrush(Color.FromRgb(30, 30, 30)) : new SolidColorBrush(Color.FromRgb(240, 240, 240));

            // 主卡片背景
            var cardBgColor = dark ? new SolidColorBrush(Color.FromArgb(200, 30, 30, 30))
                                   : new SolidColorBrush(Color.FromArgb(230, 255, 255, 255));

            if (InputPanel is not null) InputPanel.Background = cardBgColor;
            if (NoticePanel is not null) NoticePanel.Background = cardBgColor;

            // 子卡片背景
            var subCardBg = dark ? new SolidColorBrush(Color.FromArgb(230, 45, 45, 45))
                                 : new SolidColorBrush(Color.FromArgb(255, 248, 249, 250));
            if (loginArea is not null) loginArea.Background = subCardBg;
            if (entryArea is not null) entryArea.Background = subCardBg;

            // 文本框背景色
            var textBgColor = dark ? new SolidColorBrush(Color.FromRgb(60, 60, 60)) : Brushes.White;
            if (txtPassword is not null) txtPassword.Background = textBgColor;
            if (txtItem is not null) txtItem.Background = textBgColor;
        }

        // ==========================================
        // 辅助方法
        // ==========================================
        private void SetEntryAreaEnabled(bool enabled)
        {
            if (entryArea is not null)
            {
                entryArea.IsEnabled = enabled;
                if (enabled && isSeasonStopped)
                {
                    entryArea.IsEnabled = false;
                }
            }
        }

        // 单击单元格进入编辑模式
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is DataGridCell cell && !cell.IsEditing && !cell.IsReadOnly)
            {
                cell.IsEditing = true;
            }
        }

        // ==========================================
        // 事件处理
        // ==========================================
        private void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            string pwd = txtPassword.Password.Trim();
            var today = DateTime.Today.DayOfWeek;

            if ((DayPasswords.TryGetValue(today, out var todayPwd) && todayPwd == pwd && DayOwners.TryGetValue(pwd, out currentOperator)) ||
                MasterPasswords.TryGetValue(pwd, out currentOperator))
            {
                MessageBox.Show($"欢迎 {currentOperator}，解锁成功！", "成功");
                txtPassword.IsEnabled = false;
                btnConfirm.IsEnabled = false;
                btnLogout.IsEnabled = true;
                SetEntryAreaEnabled(!isSeasonStopped);
                txtItem.Focus();
            }
            else
            {
                MessageBox.Show("密码不正确！", "错误");
                txtPassword.Clear();
                txtPassword.Focus();
            }
        }

        private void BtnLogout_Click(object sender, RoutedEventArgs e)
        {
            currentOperator = null;
            txtPassword.IsEnabled = true;
            btnConfirm.IsEnabled = true;
            btnLogout.IsEnabled = false;
            SetEntryAreaEnabled(false);
            txtPassword.Clear();
            MessageBox.Show("已退出登录", "提示");
        }

        private async void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            string item = txtItem.Text.Trim();
            if (string.IsNullOrEmpty(item))
            {
                MessageBox.Show("项目不能为空！");
                return;
            }

            var changes = new List<(int group, double score, string name)>();
            foreach (var entry in _scoreEntries)
            {
                string name = entry.Name.Trim();
                if (string.IsNullOrEmpty(name)) continue;

                if (!int.TryParse(entry.Group, out int group) || !TryParseScore(entry.Score, out double score))
                {
                    MessageBox.Show($"请检查 {name} 的填写内容。");
                    return;
                }
                changes.Add((group, score, name));
            }

            if (changes.Count == 0)
            {
                MessageBox.Show("请至少填写一条有效记录！");
                return;
            }

            btnSubmit.IsEnabled = false;
            btnSubmit.Content = "写入中...";
            try
            {
                await Task.Run(() => UpdateExcel(changes, item, currentOperator ?? "未知"));
                MessageBox.Show("写入成功！");
                txtItem.Clear();
                foreach (var entry in _scoreEntries)
                {
                    entry.Name = "";
                    entry.Group = "1";
                    entry.Score = "";
                }
            }
            catch (IOException)
            {
                MessageBox.Show("写入失败：Excel 文件可能正被办公软件打开，请先关闭该工作簿。", "文件占用提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"写入异常: {ex.Message}");
            }
            finally
            {
                btnSubmit.IsEnabled = true;
                btnSubmit.Content = "写入 Excel";
            }
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            _scoreEntries.Add(new ScoreEntryModel());
        }

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
            string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ExcelFile);
            using var wb = File.Exists(excelPath) ? new XLWorkbook(excelPath) : new XLWorkbook();

            var wsRecord = wb.Worksheets.Contains("RECORD") ? wb.Worksheet("RECORD") : wb.AddWorksheet("RECORD");
            var wsLog = wb.Worksheets.Contains("LOG") ? wb.Worksheet("LOG") : wb.AddWorksheet("LOG");

            if (wsRecord.LastRowUsed() is null)
            {
                wsRecord.Cell(1, 1).Value = "组别";
                wsRecord.Cell(1, 2).Value = "积分";
                for (int i = 1; i <= 12; i++)
                {
                    wsRecord.Cell(i + 1, 1).Value = i;
                    wsRecord.Cell(i + 1, 2).Value = 0;
                }
            }

            Dictionary<int, int> groupRowMap = new();
            for (int r = 2; r <= (wsRecord.LastRowUsed()?.RowNumber() ?? 1); r++)
                if (int.TryParse(wsRecord.Cell(r, 1).GetValue<string>(), out int g))
                    groupRowMap[g] = r;

            foreach (var (group, score, name) in changes)
            {
                if (groupRowMap.TryGetValue(group, out int row))
                {
                    var cell = wsRecord.Cell(row, 2);
                    double oldScore = cell.DataType == XLDataType.Number ? cell.GetDouble() : (double.TryParse(cell.GetString(), out double parsed) ? parsed : 0);
                    cell.Value = oldScore + score;
                }
            }

            static string formatScore(double s) => s >= 0 ? $"+{s:0.##}" : s.ToString("0.##");
            string detail = string.Join("，", changes.Select(x => $"#{x.group:00}{formatScore(x.score)}（{x.name}）")) + $"【{item}】";

            if (wsLog.LastRowUsed() is null)
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

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    // ==========================================
    // 通知数据结构
    // ==========================================
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
        public bool IsSeasonStop { get; set; } = false;
        public string Image { get; set; } = "";
    }

    public class NoticeDataWrapper
    {
        public List<NoticeData> Notices { get; set; } = new();
        public bool IsSeasonStop { get; set; } = false;
    }
}