using System;
using System.Collections.Generic;
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
using Microsoft.Win32;
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

            private SolidColorBrush GetBrush(string colorCode, Color defaultColor)
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
        private readonly Dictionary<DayOfWeek, string> DayPasswords = new()
        {
            { DayOfWeek.Monday, "qwer1234" },
            { DayOfWeek.Tuesday, "xycloud" },
            { DayOfWeek.Wednesday, "wyq123456" },
            { DayOfWeek.Thursday, "249358" },
            { DayOfWeek.Friday, "wanghr5" },
            { DayOfWeek.Saturday, "sat666" },
            { DayOfWeek.Sunday, "1234asdf" }
        };

        private readonly Dictionary<string, string> DayOwners = new()
        {
            { "qwer1234", "王佳森" },
            { "xycloud", "肖赟" },
            { "wyq123456", "王玉祺" },
            { "249358", "陈妍熙" },
            { "wanghr5", "王皓然" },
            { "sat666", "轮值人员" },
            { "gzrooster", "龚子" },
            { "1234asdf", "赵欣然" }
        };

        private readonly Dictionary<string, string> MasterPasswords = new()
        {
            { "gzrooster", "龚子" },
            { "39C5BB", "李艺渊" },
            { "zrj0730123", "曾睿婕" }
        };

        private const string ExcelFile = "score.xlsx";
        private const string NoticeFile = "notice.json";
        private string? currentOperator;
        private bool isDarkMode;
        private BitmapImage? originalBackground;
        private BitmapImage? darkBackground;
        private bool isSeasonStopped = false;

        // 控件引用（由 XAML 自动生成，这里仅用于安全访问）
        private Border? loginArea;
        private Border? entryArea;

        // ==========================================
        // 构造函数 & 初始化
        // ==========================================
        public MainWindow()
        {
            InitializeComponent();

            // 获取分离区域的引用
            loginArea = FindName("LoginArea") as Border;
            entryArea = FindName("EntryArea") as Border;

            // 初始化1行数据输入（默认只有1行）
            var entries = new List<ScoreEntryModel>();
            for (int i = 0; i < 1; i++) entries.Add(new ScoreEntryModel());
            ScoreEntries.ItemsSource = entries;

            // 加载背景图
            LoadBackgroundImages();
            UpdateBackground();

            // 加载通知（根据赛季状态设置界面）
            LoadNotices();

            // 设置初始可用状态：登录区始终可用，录入区禁用（待登录）
            SetEntryAreaEnabled(false);
            if (loginArea != null) loginArea.IsEnabled = true;

            // 监听系统主题变更
            SystemEvents.UserPreferenceChanged += OnSystemThemeChanged;
            Closed += (s, e) => SystemEvents.UserPreferenceChanged -= OnSystemThemeChanged;

            // 同步当前系统主题
            SyncWithSystemTheme();
        }

        // ==========================================
        // 背景加载与切换
        // ==========================================
        private void LoadBackgroundImages()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string lightPath = Path.Combine(baseDir, "bg.png");
            string darkPath = Path.Combine(baseDir, "dk.png");
            string overlayPath = Path.Combine(baseDir, "overlay.png");

            if (File.Exists(lightPath)) originalBackground = LoadBitmapImage(lightPath);
            if (File.Exists(darkPath)) darkBackground = LoadBitmapImage(darkPath);
            if (File.Exists(overlayPath)) OverlayImage.Source = LoadBitmapImage(overlayPath);
        }

        private BitmapImage LoadBitmapImage(string path)
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.UriSource = new Uri(path, UriKind.Absolute);
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.EndInit();
            return bitmap;
        }

        private void UpdateBackground()
        {
            if (isDarkMode && darkBackground != null)
                BaseImage.Source = darkBackground;
            else if (!isDarkMode && originalBackground != null)
                BaseImage.Source = originalBackground;
            else
                BaseImage.Source = null;
        }

        // ==========================================
        // 系统主题同步
        // ==========================================
        private void SyncWithSystemTheme()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                if (key?.GetValue("AppsUseLightTheme") is int v)
                {
                    bool systemIsDark = v == 0;
                    if (systemIsDark != isDarkMode)
                    {
                        isDarkMode = systemIsDark;
                        chkDarkMode.IsChecked = isDarkMode;
                        ApplyDarkMode(isDarkMode);
                    }
                }
            }
            catch { }
        }

        private void OnSystemThemeChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.General)
                Dispatcher.Invoke(SyncWithSystemTheme);
        }

        private void ChkDarkMode_Changed(object sender, RoutedEventArgs e)
        {
            isDarkMode = chkDarkMode.IsChecked == true;
            ApplyDarkMode(isDarkMode);
        }

        private void ApplyDarkMode(bool dark)
        {
            UpdateBackground();

            // 主卡片背景（左侧面板和右侧通知区的父容器）
            var cardBgColor = dark ? new SolidColorBrush(Color.FromArgb(200, 30, 30, 30))
                                   : new SolidColorBrush(Color.FromArgb(230, 255, 255, 255));

            // 左侧面板（InputPanel）已在 XAML 中定义，通过 x:Name 可以直接使用
            if (InputPanel != null) InputPanel.Background = cardBgColor;

            // 右侧通知区的父 Border 通过 FindName 查找
            var noticePanel = FindName("NoticePanel") as Border;
            if (noticePanel == null && NoticeList?.Parent is Border parentBorder)
                noticePanel = parentBorder;
            if (noticePanel != null) noticePanel.Background = cardBgColor;

            // 子卡片背景（登录区、录入区）
            var subCardBg = dark ? new SolidColorBrush(Color.FromArgb(230, 45, 45, 45))
                                 : new SolidColorBrush(Color.FromArgb(255, 248, 249, 250));
            if (loginArea != null) loginArea.Background = subCardBg;
            if (entryArea != null) entryArea.Background = subCardBg;

            // 文本输入框背景色
            var textBgColor = dark ? new SolidColorBrush(Color.FromRgb(60, 60, 60)) : Brushes.White;
            if (txtPassword != null) txtPassword.Background = textBgColor;
            if (txtItem != null) txtItem.Background = textBgColor;
        }

        // ==========================================
        // 通知加载（赛季结束仅禁用录入区，登录区保持可用）
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
                if (data == null)
                {
                    MessageBox.Show("通知文件格式不正确。", "错误");
                    return;
                }

                isSeasonStopped = data.IsSeasonStop;

                if (isSeasonStopped)
                {
                    // 赛季结束：显示覆盖图，隐藏通知列表，禁用录入区（登录区保持可用）
                    var seasonStopNotice = data.Notices.FirstOrDefault();
                    if (seasonStopNotice != null && !string.IsNullOrEmpty(seasonStopNotice.Image))
                    {
                        OverlayImage.Source = LoadBitmapImage(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, seasonStopNotice.Image));
                    }
                    OverlayImage.Visibility = Visibility.Visible;
                    NoticeList.Visibility = Visibility.Collapsed;

                    // 录入区禁用
                    SetEntryAreaEnabled(false);
                    // 登录区保持可用
                    if (loginArea != null) loginArea.IsEnabled = true;
                }
                else
                {
                    // 正常模式：隐藏覆盖图，显示通知列表
                    OverlayImage.Visibility = Visibility.Collapsed;
                    NoticeList.Visibility = Visibility.Visible;

                    // 构建通知列表
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
                            ? LoadBitmapImageFromIcon(Drawing.SystemIcons.Warning)
                            : LoadBitmapImageFromIcon(Drawing.SystemIcons.Information)
                    }).ToList();

                    NoticeList.ItemsSource = items;

                    // 录入区根据登录状态决定启用（初始未登录，禁用）
                    SetEntryAreaEnabled(false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载通知失败：{ex.Message}", "错误");
            }
        }

        private BitmapImage LoadBitmapImageFromIcon(Drawing.Icon icon)
        {
            using var bitmap = icon.ToBitmap();
            using var ms = new MemoryStream();
            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            ms.Seek(0, SeekOrigin.Begin);
            var img = new BitmapImage();
            img.BeginInit();
            img.StreamSource = ms;
            img.CacheOption = BitmapCacheOption.OnLoad;
            img.EndInit();
            return img;
        }

        // ==========================================
        // 辅助方法
        // ==========================================
        /// <summary>设置录入区的启用状态</summary>
        private void SetEntryAreaEnabled(bool enabled)
        {
            if (entryArea != null)
            {
                entryArea.IsEnabled = enabled;
                // 如果录入区被禁用，且当前处于赛季结束状态，保持禁用；否则允许启用
                if (enabled && isSeasonStopped)
                {
                    entryArea.IsEnabled = false; // 赛季结束时强制禁用
                }
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
                // 登录后，如果赛季未结束则启用录入区，否则保持禁用
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
            // 退出登录后禁用录入区
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

            var entries = ScoreEntries.ItemsSource as List<ScoreEntryModel>;
            if (entries == null) return;

            var changes = new List<(int group, double score, string name)>();
            for (int i = 0; i < entries.Count; i++)
            {
                var entry = entries[i];
                string name = entry.Name.Trim();
                if (string.IsNullOrEmpty(name)) continue;

                if (!int.TryParse(entry.Group, out int group) || !TryParseScore(entry.Score, out double score))
                {
                    MessageBox.Show($"第 {i + 1} 行填写有误。");
                    return;
                }
                changes.Add((group, score, name));
            }

            if (changes.Count == 0)
            {
                MessageBox.Show("请至少填写一条！");
                return;
            }

            btnSubmit.IsEnabled = false;
            btnSubmit.Content = "写入中...";
            try
            {
                await Task.Run(() => UpdateExcel(changes, item, currentOperator ?? "未知"));
                MessageBox.Show("写入成功！");
                txtItem.Clear();
                foreach (var entry in entries)
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
            if (ScoreEntries.ItemsSource is List<ScoreEntryModel> entries)
            {
                entries.Add(new ScoreEntryModel());
                // 刷新绑定
                ScoreEntries.ItemsSource = null;
                ScoreEntries.ItemsSource = entries;
            }
        }

        private static bool TryParseScore(string text, out double score)
        {
            string s = text.Trim().Replace(" ", "");
            if (string.IsNullOrEmpty(s)) { score = 0; return false; }
            if (s.StartsWith("/")) s = "+" + s[1..];
            else if (!s.StartsWith('+') && !s.StartsWith('-')) s = "+" + s;
            return double.TryParse(s, out score);
        }

        private void UpdateExcel(List<(int group, double score, string name)> changes, string item, string operatorName)
        {
            string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ExcelFile);
            using var wb = File.Exists(excelPath) ? new XLWorkbook(excelPath) : new XLWorkbook();

            var wsRecord = wb.Worksheets.Contains("RECORD") ? wb.Worksheet("RECORD") : wb.AddWorksheet("RECORD");
            var wsLog = wb.Worksheets.Contains("LOG") ? wb.Worksheet("LOG") : wb.AddWorksheet("LOG");

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

            string formatScore(double s) => s >= 0 ? $"+{s:0.##}" : s.ToString("0.##");
            string detail = string.Join("，", changes.Select(x => $"#{x.group:00}{formatScore(x.score)}（{x.name}）")) + $"【{item}】";

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

        // 窗口拖动
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        // 关闭按钮事件
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    // ==========================================
    // 通知数据结构（与 JSON 对应）
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