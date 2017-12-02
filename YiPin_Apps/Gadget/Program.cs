﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Gadget
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
            //Application.Run(new _库存盘点());//库存盘点
            Application.Run(new _工资计算());//工资计算
        }
    }
}
