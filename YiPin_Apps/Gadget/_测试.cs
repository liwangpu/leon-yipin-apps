using System;
using System.Collections.Generic;
using System.ComponentModel;
using CommonLibs;
using LinqToExcel;
using LinqToExcel.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Gadget
{
    public partial class _测试 : Form
    {
        public _测试()
        {
            InitializeComponent();
        }

        private void _测试_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "CSV文件|*.csv";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                //txtUpJiaoHuo.Text = OpenFileDialog1.FileName;

                using (var csv = new ExcelQueryFactory(OpenFileDialog1.FileName))
                {
                    var result = from it in csv.Worksheet<_测试类>()
                                 select it;
                    var a = result.ToList();
                    var t = 1;
                }

            }
        }

        [ExcelTable("测试类")]
        class _测试类
        {
            [ExcelColumn("采购员")]
            public string _采购员 { get; set; }
        }
    }
}
