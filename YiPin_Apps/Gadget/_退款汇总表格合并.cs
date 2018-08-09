using CommonLibs;
using Gadget.Libs;
using LinqToExcel;
using LinqToExcel.Attributes;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Gadget
{
    public partial class _退款汇总表格合并 : Form
    {
        public _退款汇总表格合并()
        {
            InitializeComponent();
        }

        private void btn合并_Click(object sender, EventArgs e)
        {
            var pathStr = txtPath.Text;
            if (!string.IsNullOrWhiteSpace(pathStr))
            {
                btn浏览.Enabled = false;
                btn合并.Enabled = false;
                var paths = pathStr.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                var list = new List<_退款数据>();
                var result = new List<_统计结果>();
                #region 读取数据
                var actReadData = new Action(() =>
                {
                    foreach (var path in paths)
                    {
                        using (var excel = new ExcelQueryFactory(path))
                        {
                            var qs = from c in excel.Worksheet<_退款数据>(1)
                                     select c;
                            list.AddRange(qs);
                        }
                    }
                });
                #endregion

                #region 处理数据
                actReadData.BeginInvoke((obj) =>
                {
                    var list卖家简称s = list.Select(x => x._卖家简称).Distinct().ToList();
                    var list物流方式s = list.Select(x => x._物流方式).Distinct().ToList();
                    var list月份s = list.Select(x => x._月份).Distinct().OrderBy(x => x).ToList();
                    foreach (var _卖家简称s in list卖家简称s)
                    {
                        foreach (var _物流方式s in list物流方式s)
                        {
                            foreach (var _月份s in list月份s)
                            {
                                var refers = list.Where(x => x._卖家简称 == _卖家简称s && x._物流方式 == _物流方式s && x._月份 == _月份s).ToList();
                                var mode = new _统计结果();
                                mode._卖家简称 = _卖家简称s;
                                mode._物流方式 = _物流方式s;
                                mode._月份 = _月份s;
                                mode._发货 = refers.Count;
                                mode._退款 = refers.Where(x => x._已退款 == 1).Count();
                                result.Add(mode);
                            }
                        }
                    }
                    Export(result);
                }, null);
                #endregion
            }
        }

        private void btn浏览_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "CSV 文件|*.csv";//设置文件类型
            OpenFileDialog1.Title = "CSV 文件";//设置标题
            OpenFileDialog1.Multiselect = true;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtPath.Text = string.Join(",", OpenFileDialog1.FileNames);
            }
        }

        private void Export(List<_统计结果> list)
        {
            var buffer = new byte[0];
            #region 订单分配
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");


                #region 标题行
                sheet1.Cells[1, 1].Value = "卖家简称";
                sheet1.Cells[1, 2].Value = "物流方式";
                sheet1.Cells[1, 3].Value = "月份";
                sheet1.Cells[1, 4].Value = "发货数量";
                sheet1.Cells[1, 5].Value = "退货数量";
                sheet1.Cells[1, 6].Value = "比例";


                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = list.Count; idx < len; idx++)
                {
                    var curOrder = list[idx];
                    sheet1.Cells[rowIdx, 1].Value = curOrder._卖家简称;
                    sheet1.Cells[rowIdx, 2].Value = curOrder._物流方式;
                    sheet1.Cells[rowIdx, 3].Value = curOrder._月份;
                    sheet1.Cells[rowIdx, 4].Value = curOrder._发货;
                    sheet1.Cells[rowIdx, 5].Value = curOrder._退款;
                    sheet1.Cells[rowIdx, 6].Value = curOrder._比例;

                    rowIdx++;
                }
                #endregion

                #region 全部边框
                {
                    var endRow = sheet1.Dimension.End.Row;
                    var endColumn = sheet1.Dimension.End.Column;
                    using (var rng = sheet1.Cells[1, 1, endRow, endColumn])
                    {
                        rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    }
                }
                #endregion

                sheet1.Cells[sheet1.Dimension.Address].AutoFitColumns();

                buffer = package.GetAsByteArray();
            }
            #endregion

            InvokeMainForm((obj) =>
            {

                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
                saveFile.Title = "导出数据";//设置标题
                saveFile.AddExtension = true;//是否自动增加所辍名
                saveFile.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
                if (saveFile.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
                {
                    var FileName = saveFile.FileName;//得到文件路径   
                    var saveFilName = Path.GetFileNameWithoutExtension(FileName);
                    var savePath = Path.GetDirectoryName(FileName);

                    try
                    {
                        var len = buffer.Length;
                        using (var fs = File.Create(FileName, len))
                        {
                            fs.Write(buffer, 0, len);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "温馨提示");
                    }
                    btn浏览.Enabled = true;
                    btn合并.Enabled = true;
                }
            }, null);
        }

        #region InvokeMainForm 调用主线程
        protected void InvokeMainForm(Action act)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(act);
            }
            else
            {
                act.Invoke();
            }
        }
        protected void InvokeMainForm(Action<object> act, object obj)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(act, obj);
            }
            else
            {
                act.Invoke(obj);
            }
        }
        #endregion

        class _退款数据
        {
            [ExcelColumn("内部便签")]
            public string _内部便签 { get; set; }
            [ExcelColumn("卖家简称")]
            public string _卖家简称 { get; set; }
            [ExcelColumn("交易时间(中国)")]
            public DateTime _交易时间 { get; set; }
            [ExcelColumn("物流方式")]
            public string _物流方式 { get; set; }
            public int _已退款
            {
                get
                {
                    if (!string.IsNullOrWhiteSpace(_内部便签))
                        return _内部便签.IndexOf("退款") > 0 ? 1 : 0;
                    return 0;
                }
            }
            public int _月份
            {
                get
                {
                    return _交易时间.Month;
                }
            }
        }

        class _统计结果
        {
            public string _卖家简称 { get; set; }
            public string _物流方式 { get; set; }
            public int _月份 { get; set; }
            public int _发货 { get; set; }
            public int _退款 { get; set; }
            public decimal _比例
            {
                get
                {
                    if (_发货 <= 0)
                        return 0;
                    return Math.Round(_退款*1m / _发货 , 4);

                }
            }
        }
    }
}
