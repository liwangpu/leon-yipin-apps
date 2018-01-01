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
            //var list = new List<_测试类>();
            //for (int i = 0; i < 10; i++)
            //{
            //    var model = new _测试类();
            //    model._采购员 = Guid.NewGuid().ToString();
            //    list.Add(model);
            //}

            //var buffer = XlsxHelper.RtpExcel(list);


            var list = new List<decimal>() 
            {
            1,1,1,1,1,9,2,1,1,1,1,1,4,5,2,1,1,1,1,3,2,2,1,1,1,1,1,1,1,1,11,14
            };

            var outlist = new List<decimal>();

            var helper = new MathHelper();
            var sum = helper.SumKickOutlier(list, out outlist, OutlierRatio.Twice);

            

        }

        [RptTable("测试类")]
        class _测试类
        {
            [RtpColumn("采购员")]
            public string _采购员 { get; set; }
        }
    }
}
