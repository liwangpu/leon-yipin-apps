﻿using System;
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
            //Application.Run(new Form1());
            //Application.Run(new Form2());
            //Application.Run(new Form3());
            Application.Run(new Form4());
            //Application.Run(new Form5());
        }
    }
}
