# ScoreLauncher（乞分君）

一个基于 C# WinForms 开发的图形化计分工具，用于班级积分管理、记分员轮值记录与数据归档。

## Features

### 记分员密码登录

- 工作日轮值密码验证  
- 管理员主密码登录  
- 对应操作人自动记录  
- 登录/退出状态切换

## 积分录入

支持批量录入：

- 姓名  
- 组别（1-12组）  
- 加分 / 扣分  
- 项目名称

格式自动生成：

```text
#05+1（xxx），#06+1（xxx）【英语早读】
```

---

## Excel 数据写入

使用：

- ClosedXML

自动维护：

### RECORD 工作表

用于存储：

- 各组总积分

### LOG 工作表

用于存储：

- 时间戳  
- 记分明细  
- 操作人

---

## 公告系统

支持：

- 公告通知  
- 警告通知  
- JSON配置加载  
- 自定义字体大小  
- 自定义颜色  
- 图片公告

配置文件：

```text
notice.json
```

---

## 赛季停摆机制

支持：

- 全局暂停计分  
- 显示停摆图片覆盖层  
- 禁用录入区

---

## 深色模式

支持：

- 跟随 Windows 系统主题  
- 手动切换深色模式  
- DataGridView 深色适配  
- 公告区深色适配

---

## 项目结构

```text
scorerlauncher/
├── Form1.cs
├── score.xlsx
├── notice.json
├── ico.ico
└── seasonstop image files
```

---

## Tech Stack

- C#
- WinForms
- ClosedXML
- Newtonsoft.Json

---

## 功能特点

- 图形化操作
- 批量计分
- 自动日志归档
- 多权限密码系统
- 公告系统
- 深浅色模式
- 赛季停摆机制

---

## Future Plans

- 成员数据库
- 自动排行榜
- Excel 图表生成
- 局域网同步记分
- 多端联动

---

## License

© Beggars' Group LLC™

Authors:

- 鬼狗子-Zero  
- 月汐-Zero  
- 清弦-Noesis
