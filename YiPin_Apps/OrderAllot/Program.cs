using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace OrderAllot
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());//订单分配
            //Application.Run(new Form2());//工作完成情况
            //Application.Run(new Form3());
            //Application.Run(new Form4());
            //Application.Run(new Form4Spec());//订单分配(排除重复项)
            //Application.Run(new Form5());
            Application.Run(new Form6());//计算工资
        }
    }
}
