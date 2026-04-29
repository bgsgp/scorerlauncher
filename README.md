# 🎯 ScoreLauncher（乞分君）

一个基于 **C# WinForms** 开发的班级积分管理工具，支持批量计分、多权限登录、自动日志归档、深色模式与赛季停摆机制。

---

## ✨ 核心功能

- **多权限密码登录**  
  支持管理员主密码与工作日轮值密码，自动记录操作人，随时切换登录/退出状态。  
  *已内置零硬编码安全体系，密码于安装期配置并 BCrypt 哈希入库，安装完毕即删除明文。*

- **批量积分录入**  
  一次性地录入姓名、组别（1-12）、加减分与项目名称，自动生成格式：  
  `#05+1（xxx），#06+1（xxx）【英语早读】`

- **Excel 数据归档**  
  基于 ClosedXML，自动维护两个工作表：
  - `RECORD` — 各组总积分  
  - `LOG` — 时间戳、明细与操作人记录

- **公告系统**  
  支持文字公告、警告通知、图片公告，通过 `notice.json` 自定义字体大小、颜色等。

- **赛季停摆机制**  
  全局暂停计分时，界面显示停摆图片覆盖层并禁用录入区。

- **深色模式**  
  跟随 Windows 系统主题或手动切换，DataGridView、公告区、按钮及子控件均已完整适配。

- **密码管理程序**  
  通过 `-kc` 参数启动独立管理界面，需管理员验证后使用。支持增、删、改所有非管理员账户，权限下拉自由切换，用户名直接编辑，保存自动检查重复。界面无遮挡、滚动条归零，深色模式全覆盖。

---

## 🔗 推荐联动工具：积分备份器

强烈推荐搭配使用独立工具 **[积分备份器](https://github.com/bgsgp/xlsbd)**，为你的积分数据增加一道自动安全保障。

- 📁 **按日期自动备份**：每天可生成带编号的备份副本，永不覆盖，安全归档
- 👁 **全屏一键查 LOG**：备份完成后立刻用 Excel 全屏打开，自动定位到 LOG 表的最新记录行，无需手动翻找
- ⏱ **定时无人值守**：可加入 Windows 任务计划，实现每日定时自动备份，全程无感运行

**推荐工作流**：  
　① 在“乞分君”中完成积分录入  
　② 双击运行“积分备份器.exe”（或等待定时任务自动执行）  
　③ 弹出的 Excel 全屏窗口直接展示最新日志，快速核对

> 📥 独立 exe 开箱即用，无需安装 Python 环境。详见 [xlsbd 仓库](https://github.com/bgsgp/xlsbd)。

---

## 🛠 项目结构

```
scorerlauncher/
├── Form1.cs
├── Form1.Designer.cs
├── score.xlsx
├── notice.json
├── ico.ico
└── seasonstop image files（可选）
```

---

## ⚙️ 技术栈

- **语言**：C#
- **框架**：WinForms（.NET 10.0）
- **Excel**：ClosedXML
- **JSON**：Newtonsoft.Json
- **数据库**：Microsoft.EntityFrameworkCore.Sqlite
- **密码哈希**：BCrypt.Net-Next

---

## 📈 Star History

[![Star History Chart](https://api.star-history.com/svg?repos=bgsgp/scorerlauncher&type=Date&theme=dark)](https://star-history.com/#bgsgp/scorerlauncher&Date)

---

## 🔄 最近更新

- 密码管理界面大幅优化，登录区域验证后自动隐藏，管理面板全窗口填充
- 彻底修复数据首行被列标题遮挡的问题（列标题固定 24px，加载后滚动条归零）
- 深色模式完整覆盖所有子控件（文本框、下拉框、面板、按钮）
- 主界面“深色模式”复选框文字改为“暗夜模式”

---

## 📃 许可证与作者

© 丐帮集团第一院·物理版象棋开发与研究院™

Authors:
- 鬼狗子-Zero、月汐-Zero、清弦-Zero

此项目仅供学习与内部使用，请遵守相关许可条例。
