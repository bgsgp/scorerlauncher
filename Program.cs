using System;
using System.Windows.Forms;

namespace scorerlauncher;

static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        if (args.Length > 0 && args[0].Equals("-kc", StringComparison.OrdinalIgnoreCase))
            Application.Run(new Form2());   // 密码管理
        else
            Application.Run(new Form1());   // 主界面
    }
}