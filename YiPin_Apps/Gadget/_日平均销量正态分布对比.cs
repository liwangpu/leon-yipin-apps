using CommonLibs;
using Gadget.Libs;
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
    public partial class _日平均销量正态分布对比 : Form
    {
        public _日平均销量正态分布对比()
        {
            InitializeComponent();
        }

        private void _日平均销量正态分布对比_Load(object sender, EventArgs e)
        {
            //txt月销量流水.Text = @"C:\Users\Leon\Desktop\近30天销量.csv";
            //btn计算.Enabled = true;

        }

        private void btn上传月销量流水_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt月销量流水, () =>
            {
                btn计算.Enabled = true;
            });
        }

        #region 处理数据
        private void btn计算_Click(object sender, EventArgs e)
        {
            btn计算.Enabled = false;
            var list月销量流水 = new List<_每月流水>();
            var list分析结果 = new List<_分析结果>();
            #region 读取数据
            var actReadData = new Action(() =>
            {
                var strError = string.Empty;
                ShowMsg("开始读取月销量表数据");
                FormHelper.ReadCSVFile(txt月销量流水.Text, ref list月销量流水, ref strError);
            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                ShowMsg("正在处理数据");
                foreach (var item in list月销量流水)
                {
                    var model = new _分析结果();
                    model.SKU = item.SKU;

                    #region 拉依达准则
                    {
                        model._正态分布_期望 = Math.Round(item._月销量流水.Average(), 2);
                        double stdev = CalculateStdDev(item._月销量流水.Select(x => Convert.ToDouble(x)));
                        model._正态分布_标准差 = Convert.ToDecimal(stdev);


                        #region 1倍
                        {
                            var datas = item._月销量流水.Where(x => x > model._正态分布_1倍左区间 && x < model._正态分布_1倍右区间).ToList();
                            var sum = datas.Sum();
                            if (datas.Count > 0)
                                model._正态分布_1倍日销量 = Math.Round(sum / datas.Count, 2);
                        }
                        #endregion

                        #region 2倍
                        {
                            var datas = item._月销量流水.Where(x => x > model._正态分布_2倍左区间 && x < model._正态分布_2倍右区间).ToList();
                            var sum = datas.Sum();
                            if (datas.Count > 0)
                                model._正态分布_2倍日销量 = Math.Round(sum / datas.Count, 2);
                        }
                        #endregion

                        #region 3倍
                        {
                            var datas = item._月销量流水.Where(x => x > model._正态分布_3倍左区间 && x < model._正态分布_3倍右区间).ToList();
                            var sum = datas.Sum();
                            if (datas.Count > 0)
                                model._正态分布_3倍日销量 = Math.Round(sum / datas.Count, 2);
                        }
                        #endregion
                    }
                    #endregion

                    #region 普通算法
                    {
                        model._普通算法_5天销量 = item._月销量流水.Take(5).Sum();
                        model._普通算法_15天销量 = item._月销量流水.Take(15).Sum();
                        model._普通算法_30天销量 = item._月销量流水.Take(30).Sum();
                    }
                    #endregion

                    #region 新算法
                    {
                        model._新算法_5天销量平均值 = PanDatas(item._月销量流水.Take(5).ToList());
                        model._新算法_10天销量平均值 = PanDatas(item._月销量流水.Take(10).ToList());
                        model._新算法_15天销量平均值 = PanDatas(item._月销量流水.Take(15).ToList());
                    }
                    #endregion

                    list分析结果.Add(model);
                }

                ExportExcel(list分析结果);
            }, null);
            #endregion
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

        #region CalculateStdDev 计算标准差
        private static double CalculateStdDev(IEnumerable<double> values)
        {
            double ret = 0;
            if (values.Count() > 0)
            {
                //  计算平均数   
                double avg = values.Average();
                //  计算各数值与平均数的差值的平方，然后求和 
                double sum = values.Sum(d => Math.Pow(d - avg, 2));
                //  除以数量，然后开方
                ret = Math.Round(Math.Sqrt(sum / values.Count()), 2);
            }
            return ret;
        }
        #endregion

        #region PanDatas 去最高最低取平均值
        private static decimal PanDatas(List<decimal> list)
        {
            decimal res = 0;
            if (list.Count > 0)
            {
                var sorts = list.OrderBy(x => x).ToList();
                var data = sorts.Skip(1).Take(sorts.Count - 2);
                var sum = data.Sum();
                res = Math.Round(sum / (list.Count - 2), 2);
            }
            return res;
        }
        #endregion

        #region ExportExcel 导出数据
        private void ExportExcel(List<_分析结果> list)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];

            #region 计算详情
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    var workbox = package.Workbook;

                    #region 详情
                    {
                        var sheet1 = workbox.Worksheets.Add("详情");

                        using (var rng = sheet1.Cells[1, 1, 2, 1])
                        {
                            rng.Value = "SKU";
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }


                        using (var rng = sheet1.Cells[1, 2, 1, 5])
                        {
                            rng.Value = "原先算法";
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        sheet1.Cells[2, 2].Value = "5天销量";
                        sheet1.Cells[2, 3].Value = "15天销量";
                        sheet1.Cells[2, 4].Value = "30天销量";
                        sheet1.Cells[2, 5].Value = "日销量";
                        using (var rng = sheet1.Cells[1, 2, 2, 5])
                        {
                            var colFromHex = System.Drawing.ColorTranslator.FromHtml("#92D050");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        using (var rng = sheet1.Cells[1, 6, 1, 9])
                        {
                            rng.Value = "去最高最低算法";
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        sheet1.Cells[2, 6].Value = "5天销量平均值";
                        sheet1.Cells[2, 7].Value = "10天销量平均值";
                        sheet1.Cells[2, 8].Value = "15天销量平均值";
                        sheet1.Cells[2, 9].Value = "日销量";
                        using (var rng = sheet1.Cells[1, 6, 2, 9])
                        {
                            var colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFF00");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        using (var rng = sheet1.Cells[1, 10, 1, 14])
                        {
                            rng.Value = "拉依达准则";
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        sheet1.Cells[2, 10].Value = "期望μ";
                        sheet1.Cells[2, 11].Value = "标准差σ";
                        sheet1.Cells[2, 12].Value = "1σ日销量";
                        sheet1.Cells[2, 13].Value = "2σ日销量";
                        sheet1.Cells[2, 14].Value = "3σ日销量";
                        using (var rng = sheet1.Cells[1, 10, 2, 14])
                        {
                            var colFromHex = System.Drawing.ColorTranslator.FromHtml("#11A7F5");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        #region 数据行
                        for (int idx = 0, len = list.Count, rowIdx = 3; idx < len; idx++, rowIdx++)
                        {
                            var data = list[idx];
                            sheet1.Cells[rowIdx, 1].Value = data.SKU;

                            sheet1.Cells[rowIdx, 2].Value = data._普通算法_5天销量;
                            sheet1.Cells[rowIdx, 3].Value = data._普通算法_15天销量;
                            sheet1.Cells[rowIdx, 4].Value = data._普通算法_30天销量;
                            sheet1.Cells[rowIdx, 5].Value = data._普通算法_日平均销量;

                            sheet1.Cells[rowIdx, 6].Value = data._新算法_5天销量平均值;
                            sheet1.Cells[rowIdx, 7].Value = data._新算法_10天销量平均值;
                            sheet1.Cells[rowIdx, 8].Value = data._新算法_15天销量平均值;
                            sheet1.Cells[rowIdx, 9].Value = data._新算法_日平均销量;

                            sheet1.Cells[rowIdx, 10].Value = data._正态分布_期望;
                            sheet1.Cells[rowIdx, 11].Value = data._正态分布_标准差;
                            sheet1.Cells[rowIdx, 12].Value = data._正态分布_1倍日销量;
                            sheet1.Cells[rowIdx, 13].Value = data._正态分布_2倍日销量;
                            sheet1.Cells[rowIdx, 14].Value = data._正态分布_3倍日销量;


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

                    }
                    #endregion

                    buffer = package.GetAsByteArray();
                }
            }
            #endregion

            #region 导出
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
                        ShowMsg(ex.Message);
                    }

                    ShowMsg("表格生成完毕");
                    btn计算.Enabled = true;
                }
            }, null);
            #endregion
        } 
        #endregion

        /**************** common class ****************/

        [ExcelTable("每月流水")]
        class _每月流水
        {
            private string _SKU;

            [ExcelColumn("SKU")]
            public string SKU
            {
                get
                {
                    return _SKU;
                }
                set
                {
                    _SKU = value != null ? value.ToString().Trim().ToUpper() : "";
                }
            }

            [ExcelColumn("今天销量")]
            public decimal _今天销量 { get; set; }
            [ExcelColumn("今天往前1天")]
            public decimal _今天往前1天 { get; set; }
            [ExcelColumn("今天往前2天")]
            public decimal _今天往前2天 { get; set; }
            [ExcelColumn("今天往前3天")]
            public decimal _今天往前3天 { get; set; }
            [ExcelColumn("今天往前4天")]
            public decimal _今天往前4天 { get; set; }
            [ExcelColumn("今天往前5天")]
            public decimal _今天往前5天 { get; set; }
            [ExcelColumn("今天往前6天")]
            public decimal _今天往前6天 { get; set; }
            [ExcelColumn("今天往前7天")]
            public decimal _今天往前7天 { get; set; }
            [ExcelColumn("今天往前8天")]
            public decimal _今天往前8天 { get; set; }
            [ExcelColumn("今天往前9天")]
            public decimal _今天往前9天 { get; set; }
            [ExcelColumn("今天往前10天")]
            public decimal _今天往前10天 { get; set; }

            [ExcelColumn("今天往前11天")]
            public decimal _今天往前11天 { get; set; }
            [ExcelColumn("今天往前12天")]
            public decimal _今天往前12天 { get; set; }
            [ExcelColumn("今天往前13天")]
            public decimal _今天往前13天 { get; set; }
            [ExcelColumn("今天往前14天")]
            public decimal _今天往前14天 { get; set; }
            [ExcelColumn("今天往前15天")]
            public decimal _今天往前15天 { get; set; }
            [ExcelColumn("今天往前16天")]
            public decimal _今天往前16天 { get; set; }
            [ExcelColumn("今天往前17天")]
            public decimal _今天往前17天 { get; set; }
            [ExcelColumn("今天往前18天")]
            public decimal _今天往前18天 { get; set; }
            [ExcelColumn("今天往前19天")]
            public decimal _今天往前19天 { get; set; }
            [ExcelColumn("今天往前20天")]
            public decimal _今天往前20天 { get; set; }

            [ExcelColumn("今天往前21天")]
            public decimal _今天往前21天 { get; set; }
            [ExcelColumn("今天往前22天")]
            public decimal _今天往前22天 { get; set; }
            [ExcelColumn("今天往前23天")]
            public decimal _今天往前23天 { get; set; }
            [ExcelColumn("今天往前24天")]
            public decimal _今天往前24天 { get; set; }
            [ExcelColumn("今天往前25天")]
            public decimal _今天往前25天 { get; set; }
            [ExcelColumn("今天往前26天")]
            public decimal _今天往前26天 { get; set; }
            [ExcelColumn("今天往前27天")]
            public decimal _今天往前27天 { get; set; }
            [ExcelColumn("今天往前28天")]
            public decimal _今天往前28天 { get; set; }
            [ExcelColumn("今天往前29天")]
            public decimal _今天往前29天 { get; set; }

            public List<decimal> _月销量流水
            {
                get
                {
                    return new List<decimal>()
                    {
                        _今天销量,
                        _今天往前1天,
                        _今天往前2天,
                        _今天往前3天,
                        _今天往前4天,
                        _今天往前5天,
                        _今天往前6天,
                        _今天往前7天,
                        _今天往前8天,
                        _今天往前9天,
                        _今天往前10天,
                        _今天往前11天,
                        _今天往前12天,
                        _今天往前13天,
                        _今天往前14天,
                        _今天往前15天,
                        _今天往前16天,
                        _今天往前17天,
                        _今天往前18天,
                        _今天往前19天,
                        _今天往前20天,
                        _今天往前21天,
                        _今天往前22天,
                        _今天往前23天,
                        _今天往前24天,
                        _今天往前25天,
                        _今天往前26天,
                        _今天往前27天,
                        _今天往前28天,
                        _今天往前29天
                    };
                }
            }

        }

        class _分析结果
        {
            public string SKU { get; set; }

            public decimal _正态分布_期望 { get; set; }
            public decimal _正态分布_标准差 { get; set; }

            public decimal _正态分布_1倍左区间 { get { return _正态分布_期望 - _正态分布_标准差; } }
            public decimal _正态分布_1倍右区间 { get { return _正态分布_期望 + _正态分布_标准差; } }
            public decimal _正态分布_1倍日销量 { get; set; }

            public decimal _正态分布_2倍左区间 { get { return _正态分布_期望 - 2 * _正态分布_标准差; } }
            public decimal _正态分布_2倍右区间 { get { return _正态分布_期望 + 2 * _正态分布_标准差; } }
            public decimal _正态分布_2倍日销量 { get; set; }

            public decimal _正态分布_3倍左区间 { get { return _正态分布_期望 - 3 * _正态分布_标准差; } }
            public decimal _正态分布_3倍右区间 { get { return _正态分布_期望 + 3 * _正态分布_标准差; } }
            public decimal _正态分布_3倍日销量 { get; set; }

            public decimal _普通算法_5天销量 { get; set; }
            public decimal _普通算法_15天销量 { get; set; }
            public decimal _普通算法_30天销量 { get; set; }
            public decimal _普通算法_日平均销量
            {
                get
                {
                    var n = (_普通算法_5天销量 / 5 + _普通算法_15天销量 / 15 + _普通算法_30天销量 / 30) / 3;
                    return Math.Round(n, 2);
                }
            }

            public decimal _新算法_5天销量平均值 { get; set; }
            public decimal _新算法_10天销量平均值 { get; set; }
            public decimal _新算法_15天销量平均值 { get; set; }
            public decimal _新算法_日平均销量
            {
                get
                {
                    var n = (_新算法_5天销量平均值 + _新算法_10天销量平均值 + _新算法_15天销量平均值) / 3;
                    return Math.Round(n, 2);
                }
            }
        }


    }
}
