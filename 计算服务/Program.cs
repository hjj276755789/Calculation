using Calculation.Base;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculation
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        //[STAThreadAttribute]
        [STAThread]
        static void Main()
        {

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //获取当前进程名称
            string currentProcessName = Process.GetCurrentProcess().ProcessName;
            //把该名称的所有进程的列表
            Process[] process = Process.GetProcessesByName(currentProcessName);
            if (process.Length > 1)
            {
                MessageBox.Show("程序已经运行");
                return;
            }
            Application.Run(new Service());

        }
    }
}

