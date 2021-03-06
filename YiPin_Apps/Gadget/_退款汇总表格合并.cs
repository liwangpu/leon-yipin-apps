﻿using LinqToExcel;
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
    public partial class _退款汇总表格合并 : Form
    {
        public _退款汇总表格合并()
        {
            InitializeComponent();
        }

        #region 上传文件
        private void btn浏览_Click(object sender, EventArgs e)
        {
            txtPath.Text = string.Empty;
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "CSV 文件|*.csv";//设置文件类型
            OpenFileDialog1.Title = "CSV 文件";//设置标题
            OpenFileDialog1.Multiselect = true;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtPath.Text = string.Join(",", OpenFileDialog1.FileNames);
                ShowMsg("待处理数据");
            }
        }
        #endregion

        #region 合并汇总
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
                    ShowMsg("正在读取数据,请稍等");
                    foreach (var path in paths)
                    {
                        using (var excel = new ExcelQueryFactory(path))
                        {
                            var qs = from c in excel.Worksheet<_退款数据>()
                                     select c;
                            list.AddRange(qs);
                        }
                    }
                });
                #endregion

                #region 处理数据
                actReadData.BeginInvoke((obj) =>
                {
                    ShowMsg("正在处理数据,请稍等");
                    foreach (var item in list)
                    {
                        var bExist = false;
                        for (int idx = result.Count - 1; idx >= 0; idx--)
                        {
                            var rsItem = result[idx];
                            if (rsItem._卖家简称 == item._卖家简称 && rsItem._物流方式 == item._物流方式
                            && rsItem._月份 == item._月份 && rsItem._交易国家 == item._交易国家)
                            {
                                rsItem._发货++;
                                if (item._已退款)
                                    rsItem._退款++;
                                bExist = true;
                                break;
                            }
                        }

                        if (!bExist)
                        {
                            var mode = new _统计结果();
                            mode._卖家简称 = item._卖家简称;
                            mode._物流方式 = item._物流方式;
                            mode._月份 = item._月份;
                            mode._交易国家 = item._交易国家;
                            mode._发货 = 1;
                            mode._退款 = item._已退款 ? 1 : 0;
                            result.Add(mode);
                        }
                    }

                    ShowMsg("即将导出表格");
                    Export_合并汇总(result);
                }, null);
                #endregion
            }
        }
        #endregion

        #region 交易数量汇总
        private void btn交易数量汇总_Click(object sender, EventArgs e)
        {
            var pathStr = txtPath.Text;
            if (!string.IsNullOrWhiteSpace(pathStr))
            {
                btn浏览.Enabled = false;
                btn合并.Enabled = false;
                var paths = pathStr.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                var list = new List<_退款数据>();
                var result = new List<_交易数量统计结果>();

                #region 读取数据
                var actReadData = new Action(() =>
                {
                    ShowMsg("正在读取数据,请稍等");
                    foreach (var path in paths)
                    {
                        using (var excel = new ExcelQueryFactory(path))
                        {
                            var qs = from c in excel.Worksheet<_退款数据>()
                                     select c;
                            list.AddRange(qs);
                        }
                    }
                });
                #endregion

                #region 处理数据
                actReadData.BeginInvoke((obj) =>
                {
                    ShowMsg("正在处理数据,请稍等");

                    foreach (var item in list)
                    {
                        var bExist = false;
                        for (int idx = result.Count - 1; idx >= 0; idx--)
                        {
                            var rsItem = result[idx];
                            if (rsItem._卖家简称 == item._卖家简称 && rsItem._交易时间 == item._交易日期)
                            {
                                rsItem._发货数量++;
                                if (item._已退款)
                                    rsItem._退款数量++;
                                bExist = true;
                                break;
                            }
                        }

                        if (!bExist)
                        {
                            var mode = new _交易数量统计结果();
                            mode._卖家简称 = item._卖家简称;
                            mode._交易时间 = item._交易日期;
                            mode._发货数量 = 1;
                            mode._退款数量 = item._已退款 ? 1 : 0;
                            result.Add(mode);
                        }
                    }

                    ShowMsg("即将导出表格");
                    Export_交易数量汇总(result.OrderBy(x => x._交易时间).ToList());
                }, null);
                #endregion
            }
        } 
        #endregion

        #region Export_合并汇总 导出合并汇总
        private void Export_合并汇总(List<_统计结果> list)
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
                sheet1.Cells[1, 3].Value = "交易国家";
                sheet1.Cells[1, 4].Value = "月份";
                sheet1.Cells[1, 5].Value = "发货数量";
                sheet1.Cells[1, 6].Value = "退货数量";
                sheet1.Cells[1, 7].Value = "比例";


                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = list.Count; idx < len; idx++)
                {
                    var curOrder = list[idx];
                    sheet1.Cells[rowIdx, 1].Value = curOrder._卖家简称;
                    sheet1.Cells[rowIdx, 2].Value = curOrder._物流方式;
                    sheet1.Cells[rowIdx, 3].Value = curOrder._交易国家;
                    sheet1.Cells[rowIdx, 4].Value = curOrder._月份;
                    sheet1.Cells[rowIdx, 5].Value = curOrder._发货;
                    sheet1.Cells[rowIdx, 6].Value = curOrder._退款;
                    sheet1.Cells[rowIdx, 7].Value = curOrder._比例;

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
                        ShowMsg("表格导出完毕");
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
        #endregion

        #region Export_交易数量汇总 导出交易数量汇总
        private void Export_交易数量汇总(List<_交易数量统计结果> list)
        {
            var buffer = new byte[0];
            #region 订单分配
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[1, 1].Value = "卖家简称";
                sheet1.Cells[1, 2].Value = "交易时间";
                sheet1.Cells[1, 3].Value = "发货总量";
                sheet1.Cells[1, 4].Value = "退货总量";



                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = list.Count; idx < len; idx++)
                {
                    var curOrder = list[idx];
                    sheet1.Cells[rowIdx, 1].Value = curOrder._卖家简称;
                    sheet1.Cells[rowIdx, 2].Value = curOrder._交易时间;
                    sheet1.Cells[rowIdx, 3].Value = curOrder._发货数量;
                    sheet1.Cells[rowIdx, 4].Value = curOrder._退款数量;

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
                        ShowMsg("表格导出完毕");
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
        #endregion

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

        #region ShowMsg 消息提示
        /// <summary>
        /// 消息提示
        /// </summary>
        /// <param name="strMsg"></param>
        private void ShowMsg(string strMsg)
        {
            if (this.InvokeRequired)
            {
                var act = new Action<string>(ShowMsg);
                this.Invoke(act, strMsg);
            }
            else
            {
                this.lbMsg.Text = strMsg;
            }
        }
        #endregion

        class _退款数据
        {
            private string org卖家简称;
            private string org交易国家;
            private string org物流方式;

            [ExcelColumn("内部便签")]
            public string _内部便签 { get; set; }
            [ExcelColumn("卖家简称")]
            public string _卖家简称
            {
                get
                {
                    return org卖家简称;
                }
                set
                {
                    org卖家简称 = value != null ? value.ToString().Trim() : "";
                }
            }
            [ExcelColumn("交易时间(中国)")]
            public DateTime _交易时间 { get; set; }
            [ExcelColumn("物流方式")]
            public string _物流方式
            {
                get
                {
                    return org物流方式;
                }
                set
                {
                    org物流方式 = value != null ? value.ToString().Trim() : "";
                }
            }
            [ExcelColumn("收货人国家中文")]
            public string _交易国家
            {
                get
                {
                    return org交易国家;
                }
                set
                {
                    org交易国家 = value != null ? value.ToString().Trim() : "";
                }
            }
            public bool _已退款
            {
                get
                {
                    if (!string.IsNullOrWhiteSpace(_内部便签))
                        return _内部便签.IndexOf("已退款") > 0;
                    return false;
                }
            }
            public int _月份
            {
                get
                {
                    return _交易时间.Month;
                }
            }
            public string _交易日期
            {
                get
                {
                    return _交易时间.ToString("yyyy-MM-dd");
                }
            }
        }

        class _统计结果
        {
            public string _卖家简称 { get; set; }
            public string _物流方式 { get; set; }
            public int _月份 { get; set; }
            public string _交易国家 { get; set; }
            public int _发货 { get; set; }
            public int _退款 { get; set; }
            public decimal _比例
            {
                get
                {
                    if (_发货 <= 0)
                        return 0;
                    return Math.Round(_退款 * 1m / _发货, 4);

                }
            }
        }

        class _交易数量统计结果
        {
            public string _卖家简称 { get; set; }
            public string _交易时间 { get; set; }
            public int _发货数量 { get; set; }
            public int _退款数量 { get; set; }
        }
    }
}
