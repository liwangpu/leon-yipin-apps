using LinqToExcel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CommonLibs;
using LinqToExcel.Attributes;

namespace Gadget
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "CSV 文件|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "CSV 文件";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                //if (Helper.CheckCSVFileName(OpenFileDialog1.FileName))
                //{
                //    txt库存明细.Text = OpenFileDialog1.FileName;
                //}

                using (var excel = new ExcelQueryFactory(OpenFileDialog1.FileName))
                {
                    var sheetNames = excel.GetWorksheetNames().ToList();
                    sheetNames.ForEach(s =>
                    {
                        try
                        {
                            var tmp = from c in excel.Worksheet<Person>(s)
                                      select c;
                            MessageBox.Show(tmp.Count().ToString());
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    });

                }

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }


    public class Person
    {
        [ExcelColumn("姓名")]
        public string Name { get; set; }

        [ExcelColumn("年纪")]
        public string Age { get; set; }
    }
}
