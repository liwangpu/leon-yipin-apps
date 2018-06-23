using LinqToExcel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CommonLibs;
using LinqToExcel.Attributes;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Gadget
{
    public partial class _分库盘点 : Form
    {
        public _分库盘点()
        {
            InitializeComponent();
        }

        private void _分库盘点_Load(object sender, EventArgs e)
        {
            //txt库存明细.Text = @"C:\Users\Leon\Desktop\yyy.csv";
            ////txt盘点结果.Text = @"C:\Users\Leon\Desktop\盘点.csv";
            //txt负责人.Text = @"C:\Users\Leon\Desktop\仓库配货人员.csv";



            txtMemo.Text = "备注：仓库人员务必全力配合库存管理人员把库存问题解决,盘点过程务必认真仔细";
        }

        public const string NotReferManager = "未指定负责人";

        /**************** button event ****************/

        #region 上传库存明细
        private void btn库存明细_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "CSV 文件|*.csv";//设置文件类型
            OpenFileDialog1.Title = "CSV 文件";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                if (Helper.CheckCSVFileName(OpenFileDialog1.FileName))
                {
                    txt库存明细.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传盘点结果
        private void btn盘点结果_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Excel 文件|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "Excel 文件";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                if (Helper.CheckCSVFileName(OpenFileDialog1.FileName))
                {
                    txt盘点结果.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传负责人信息
        private void btn负责人_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "CSV 文件|*.csv";//设置文件类型
            OpenFileDialog1.Title = "CSV 文件";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                if (Helper.CheckCSVFileName(OpenFileDialog1.FileName))
                {
                    txt负责人.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 处理数据
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            var list库存明细 = new List<_库存明细Mapping>();
            var list盘点结果 = new List<_盘点表Mapping>();
            var list负责人详细 = new List<_区域负责人Mapping>();
            var list负责人信息 = new List<KeyValuePair<string, string>>();
            var list导出结果 = new List<_分析结果Model>();
            //var i遗留上海天数 = nd遗留上海天数.Value;
            //var b只取盘点对应 = cb只要盘点结果.Checked;

            ShowMsg("开始读取表格信息");

            #region 读取数据
            var actReadData = new Action(() =>
            {
                #region 读取库存明细
                {
                    var strCSVPath = txt库存明细.Text;
                    if (!string.IsNullOrEmpty(strCSVPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCSVPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_库存明细Mapping>()
                                          select c;
                                list库存明细.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 读取盘点表
                {
                    var strExcelPath = txt盘点结果.Text;
                    if (!string.IsNullOrEmpty(strExcelPath))
                    {
                        using (ExcelPackage package = new ExcelPackage(new FileStream(strExcelPath, FileMode.Open)))
                        {
                            var workbox = package.Workbook;
                            var sheetSum = workbox.Worksheets.Count;
                            if (sheetSum > 0)
                            {

                                for (int i = 1; i <= sheetSum; i++)
                                {
                                    var curSheet = workbox.Worksheets[i];
                                    var dataSum = curSheet.Dimension.End.Row;
                                    for (int rowIdx = 4; rowIdx <= dataSum; rowIdx++)
                                    {
                                        var vSku = curSheet.Cells[rowIdx, 2].Text;
                                        var vSum = curSheet.Cells[rowIdx, 5].Text;
                                        if (!string.IsNullOrEmpty(vSum))
                                        {
                                            var model = new _盘点表Mapping();
                                            model.SKU = vSku;
                                            model._数量 = Convert.ToDecimal(vSum);
                                            list盘点结果.Add(model);
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
                #endregion

                #region 读取负责人信息
                {
                    var strCSVPath = txt负责人.Text;
                    if (!string.IsNullOrEmpty(strCSVPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCSVPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_区域负责人Mapping>()
                                          select c;
                                list负责人详细.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 负责人解析
                {
                    list负责人详细.ForEach(item =>
                    {
                        var areas = item._负责区域.Split(new string[] { ",", "，", ".", "。" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        if (areas != null)
                        {
                            areas.ForEach(areaName =>
                            {
                                list负责人信息.Add(new KeyValuePair<string, string>(areaName, item._负责人));
                            });
                        }
                    });
                }
                #endregion

            });
            #endregion

            ShowMsg("开始处理数据");

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                list库存明细.ForEach(dtItem =>
                {
                    //if (dtItem.SKU=="DNFD12D46-S2")
                    //{

                    //}

                    var model = new _分析结果Model();
                    model._SKU = dtItem.SKU;
                    model._商品名称 = dtItem._商品名称;
                    model._库位 = dtItem._库位;
                    model._可用数量 = dtItem._可用数量;
                    model._日平均销量 = dtItem._平均日销量;
                    //model._遗留上海天数 = i遗留上海天数;

                    model._单位 = dtItem._单位;
                    model._库存数量 = dtItem._库存数量;
                    model._占用数量 = dtItem._占用数量;

                    //if (dtItem.SKU == "AMKA1A13-LP")
                    //{

                    //}

                    var ref盘点Item = list盘点结果.Where(x => x.SKU == dtItem.SKU).FirstOrDefault();
                    if (ref盘点Item != null)
                    {
                        model._是否盘点 = true;
                        model._盘点数量 = ref盘点Item._数量;
                    }

                    var manager = list负责人信息.Where(x => x.Key == model._仓库);
                    if (manager != null && manager.Count() > 0)
                    {
                        model._负责人 = manager.First().Value;
                    }
                    else
                    {
                        model._负责人 = NotReferManager;
                    }


                    //if (ref盘点Item == null)
                    //{
                    //    list导出结果.Add(model);
                    //}

                    //if (b只取盘点对应)
                    //{
                    //    if (ref盘点Item != null)
                    //    {
                    //        list导出结果.Add(model);
                    //    }
                    //}
                    //else
                    //{
                    list导出结果.Add(model);
                    //}
                });

                var tmpp = list导出结果.OrderBy(x => x._库位).ToList();
                Export(tmpp.Where(x => x._是否盘点 == false).ToList(), tmpp.Where(x => x._是否盘点 == true).ToList());
            }, null);
            #endregion
        }
        #endregion

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_库存明细Mapping), typeof(_盘点表Mapping));

            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "记事本|*.txt";//设置文件类型
            saveFile.Title = "导出说明文件";//设置标题
            saveFile.AddExtension = true;//是否自动增加所辍名
            saveFile.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (saveFile.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                File.WriteAllText(saveFile.FileName, strDesc);
            }
        }
        #endregion

        /**************** common method ****************/

        #region Export 导出结果表格
        private void Export(List<_分析结果Model> notCulcList, List<_分析结果Model> culcList)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer1 = new byte[0];
            var strMemo = txtMemo.Text;
            #region 未盘点数据表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 汇总表(暂时不需要)
                //{
                //    var sheet1 = workbox.Worksheets.Add("所有");

                //    #region 标题行
                //    sheet1.Cells[1, 1].Value = "SKU";
                //    sheet1.Cells[1, 2].Value = "商品名称";
                //    sheet1.Cells[1, 3].Value = "单位";
                //    sheet1.Cells[1, 4].Value = "库位";
                //    sheet1.Cells[1, 5].Value = "仓库";
                //    sheet1.Cells[1, 6].Value = "货架";
                //    sheet1.Cells[1, 7].Value = "库存数量";
                //    sheet1.Cells[1, 8].Value = "占用数量";
                //    sheet1.Cells[1, 9].Value = "可用数量";
                //    sheet1.Cells[1, 10].Value = "盘点数量";
                //    sheet1.Cells[1, 11].Value = "负责人";
                //    #endregion

                //    #region 数据行
                //    for (int idx = 0, rowIdx = 2, len = notCulcList.Count; idx < len; idx++, rowIdx++)
                //    {
                //        var info = notCulcList[idx];
                //        sheet1.Cells[rowIdx, 1].Value = info._SKU;
                //        sheet1.Cells[rowIdx, 2].Value = info._商品名称;
                //        sheet1.Cells[rowIdx, 3].Value = info._单位;
                //        sheet1.Cells[rowIdx, 4].Value = info._库位;
                //        sheet1.Cells[rowIdx, 5].Value = info._仓库;
                //        sheet1.Cells[rowIdx, 6].Value = info._货架;
                //        sheet1.Cells[rowIdx, 7].Value = info._库存数量;
                //        sheet1.Cells[rowIdx, 8].Value = info._占用数量;
                //        sheet1.Cells[rowIdx, 9].Value = info._可用数量;
                //        //sheet1.Cells[rowIdx, 8].Value = info._盘点数量;
                //        sheet1.Cells[rowIdx, 11].Value = info._负责人;
                //    }
                //    #endregion

                //}
                #endregion

                #region 分表
                {
                    var managerNames = notCulcList.Select(x => x._负责人).Distinct().ToList();
                    managerNames.ForEach(managerName =>
                    {
                        var sheet1 = workbox.Worksheets.Add(!string.IsNullOrEmpty(managerName) ? managerName : NotReferManager);

                        #region 标题行

                        using (var rng = sheet1.Cells[1, 1, 1, 8])
                        {
                            rng.Value = strMemo;
                            rng.Merge = true;
                        }
                        using (var rng = sheet1.Cells[2, 1, 2, 2])
                        {
                            rng.Value = string.Format("盘点人员:{0}", managerName);
                            rng.Merge = true;
                        }

                        using (var rng = sheet1.Cells[2, 3, 2, 8])
                        {
                            rng.Value = string.Format("盘点日期:  {0}", DateTime.Now.ToString("MM-dd"));
                            rng.Merge = true;
                        }

                        sheet1.Cells[3, 1].Value = "库位";
                        sheet1.Cells[3, 2].Value = "SKU码";
                        sheet1.Cells[3, 3].Value = "商品名称";
                        sheet1.Cells[3, 4].Value = "单位";
                      

                        sheet1.Cells[3, 5].Value = "库存数量";
                        sheet1.Cells[3, 6].Value = "占用数量";
                        sheet1.Cells[3, 7].Value = "可用数量";

                        sheet1.Cells[3, 8].Value = "盘点数量";

                        //sheet1.Cells[1, 1].Value = "SKU";
                        //sheet1.Cells[1, 2].Value = "商品名称";
                        //sheet1.Cells[1, 3].Value = "单位";
                        //sheet1.Cells[1, 4].Value = "库位";
                        //sheet1.Cells[1, 5].Value = "仓库";
                        //sheet1.Cells[1, 6].Value = "货架";
                        //sheet1.Cells[1, 7].Value = "库存数量";
                        //sheet1.Cells[1, 8].Value = "占用数量";
                        //sheet1.Cells[1, 9].Value = "可用数量";
                        //sheet1.Cells[1, 10].Value = "盘点数量";

                        #endregion

                        #region 数据行

                        var refList = notCulcList.Where(x => x._负责人 == managerName).ToList();
                        var dataSum = refList.Count;
                        for (int idx = 0, rowIdx = 4, len = dataSum; idx < len; idx++, rowIdx++)
                        {
                            var info = refList[idx];
                            sheet1.Cells[rowIdx, 1].Value = info._库位;
                            sheet1.Cells[rowIdx, 2].Value = info._SKU;
                            sheet1.Cells[rowIdx, 3].Value = info._商品名称;
                            sheet1.Cells[rowIdx, 4].Value = info._单位;

                            sheet1.Cells[rowIdx, 5].Value = info._库存数量;
                            sheet1.Cells[rowIdx, 6].Value = info._占用数量;
                            sheet1.Cells[rowIdx, 7].Value = info._可用数量;

                            sheet1.Row(rowIdx).Height = 13.5;
                        }
                        #endregion

                        #region 样式设定
                        {
                            sheet1.Column(1).Width = 15.14;
                            sheet1.Column(2).Width = 15.14;
                            sheet1.Column(3).Width = 39;
                            sheet1.Column(4).Width = 5;
                            sheet1.Column(8).Width = 9.57;
                            sheet1.Row(1).Height = 35;

                            using (var rng = sheet1.Cells[1, 1, 3, 8])
                            {
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                rng.Style.WrapText = true;
                            }

                            using (var rng = sheet1.Cells[1, 1, 2, 8])
                            {
                                rng.Style.Font.Bold = true;//字体为粗体
                            }

                            using (var rng = sheet1.Cells[3, 1, 3, 8])
                            {
                                Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#D9D9D9");
                                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                            }

                            using (var rng = sheet1.Cells[1, 1, dataSum + 3, 8])
                            {
                                rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            }

                            using (var rng = sheet1.Cells[4, 1, dataSum + 3, 8])
                            {
                                rng.Style.Font.Name = "宋体";//字体
                                rng.Style.Font.Size = 9;//字体大小
                            }

                            sheet1.Column(5).Hidden = true;
                            sheet1.Column(6).Hidden = true;
                            sheet1.Column(7).Hidden = true;
                        }
                        #endregion
                    });
                }
                #endregion

                buffer = package.GetAsByteArray();
            }
            #endregion

            #region 已盘点数据表
            if (culcList.Count > 0)
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    var workbox = package.Workbook;

                    #region 汇总表
                    {
                        var sheet1 = workbox.Worksheets.Add("所有");

                        #region 标题行
                        sheet1.Cells[1, 1].Value = "SKU";
                        sheet1.Cells[1, 2].Value = "商品名称";
                        sheet1.Cells[1, 3].Value = "单位";
                        sheet1.Cells[1, 4].Value = "库位";
                        sheet1.Cells[1, 5].Value = "仓库";
                        sheet1.Cells[1, 6].Value = "货架";
                        sheet1.Cells[1, 7].Value = "库存数量";
                        sheet1.Cells[1, 8].Value = "占用数量";
                        sheet1.Cells[1, 9].Value = "可用数量";
                        sheet1.Cells[1, 10].Value = "盘点数量";
                        sheet1.Cells[1, 11].Value = "负责人";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 2, len = culcList.Count; idx < len; idx++, rowIdx++)
                        {
                            var info = culcList[idx];
                            sheet1.Cells[rowIdx, 1].Value = info._SKU;
                            sheet1.Cells[rowIdx, 2].Value = info._商品名称;
                            sheet1.Cells[rowIdx, 3].Value = info._单位;
                            sheet1.Cells[rowIdx, 4].Value = info._库位;
                            sheet1.Cells[rowIdx, 5].Value = info._仓库;
                            sheet1.Cells[rowIdx, 6].Value = info._货架;
                            sheet1.Cells[rowIdx, 7].Value = info._库存数量;
                            sheet1.Cells[rowIdx, 8].Value = info._占用数量;
                            sheet1.Cells[rowIdx, 9].Value = info._可用数量;
                            sheet1.Cells[rowIdx, 10].Value = info._盘点数量;
                            sheet1.Cells[rowIdx, 11].Value = info._负责人;
                        }
                        #endregion

                    }
                    #endregion

                    #region 分表
                    {
                        var areaNames = culcList.Select(x => x._归属).Distinct().ToList();
                        areaNames.ForEach(areaName =>
                        {
                            var sheet1 = workbox.Worksheets.Add(!string.IsNullOrEmpty(areaName) ? areaName : "未指定区域");

                            #region 标题行
                            sheet1.Cells[1, 1].Value = "SKU";
                            sheet1.Cells[1, 2].Value = "商品名称";
                            sheet1.Cells[1, 3].Value = "单位";
                            sheet1.Cells[1, 4].Value = "库位";
                            sheet1.Cells[1, 5].Value = "仓库";
                            sheet1.Cells[1, 6].Value = "货架";
                            sheet1.Cells[1, 7].Value = "库存数量";
                            sheet1.Cells[1, 8].Value = "占用数量";
                            sheet1.Cells[1, 9].Value = "可用数量";
                            sheet1.Cells[1, 10].Value = "盘点数量";
                            sheet1.Cells[1, 11].Value = "负责人";
                            #endregion

                            #region 数据行
                            var refList = culcList.Where(x => x._归属 == areaName).ToList();
                            for (int idx = 0, rowIdx = 2, len = refList.Count; idx < len; idx++, rowIdx++)
                            {
                                var info = refList[idx];
                                sheet1.Cells[rowIdx, 1].Value = info._SKU;
                                sheet1.Cells[rowIdx, 2].Value = info._商品名称;
                                sheet1.Cells[rowIdx, 3].Value = info._单位;
                                sheet1.Cells[rowIdx, 4].Value = info._库位;
                                sheet1.Cells[rowIdx, 5].Value = info._仓库;
                                sheet1.Cells[rowIdx, 6].Value = info._货架;
                                sheet1.Cells[rowIdx, 7].Value = info._库存数量;
                                sheet1.Cells[rowIdx, 8].Value = info._占用数量;
                                sheet1.Cells[rowIdx, 9].Value = info._可用数量;
                                sheet1.Cells[rowIdx, 10].Value = info._盘点数量;
                                sheet1.Cells[rowIdx, 11].Value = info._负责人;
                            }
                            #endregion
                        });
                    }
                    #endregion

                    buffer1 = package.GetAsByteArray();
                }
            }
            #endregion

            #region 导出结果
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
                    var pureFilName = Path.GetFileNameWithoutExtension(FileName);
                    var tmp = FileName.Split(new string[] { pureFilName }, StringSplitOptions.RemoveEmptyEntries);
                    var notCulcPath = Path.Combine(tmp[0], pureFilName + "(未盘点).xlsx");
                    var culcPath = Path.Combine(tmp[0], pureFilName + "(已盘点).xlsx");

                    var len = buffer.Length;
                    using (var fs = File.Create(notCulcPath, len))
                    {
                        fs.Write(buffer, 0, len);
                    }

                    var len1 = buffer1.Length;
                    if (len1 > 0)
                    {
                        using (var fs = File.Create(culcPath, len1))
                        {
                            fs.Write(buffer1, 0, len1);
                        }
                    }


                    ShowMsg("表格生成完毕");
                }
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

        /**************** common class ****************/

        [ExcelTable("库存明细")]
        class _库存明细Mapping
        {
            private string orgSKU;
            private string org库位;
            private string org商品名称;

            [ExcelColumn("SKU码")]
            public string SKU
            {
                get
                {
                    return orgSKU;
                }
                set
                {
                    orgSKU = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("商品名称")]
            public string _商品名称
            {
                get
                {
                    return org商品名称;
                }
                set
                {
                    org商品名称 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("单位")]
            public string _单位 { get; set; }

            [ExcelColumn("库位")]
            public string _库位
            {
                get
                {
                    return org库位;
                }
                set
                {
                    org库位 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("可用数量")]
            public decimal _可用数量 { get; set; }

            [ExcelColumn("占用数量")]
            public decimal _占用数量 { get; set; }

            [ExcelColumn("库存数量")]
            public decimal _库存数量 { get; set; }

            [ExcelColumn("30天销量")]
            public decimal _30天销量 { get; set; }

            [ExcelColumn("15天销量")]
            public decimal _15天销量 { get; set; }

            [ExcelColumn("5天销量")]
            public decimal _5天销量 { get; set; }

            public decimal _平均日销量
            {
                get
                {
                    return Math.Round((_30天销量 / 30 + _15天销量 / 15 + _5天销量 / 5) / 3, 0);
                }
            }

        }

        [ExcelTable("盘点结果")]
        class _盘点表Mapping
        {
            private string orgSKU;

            [ExcelColumn("SKU")]
            public string SKU
            {
                get
                {
                    return orgSKU;
                }
                set
                {
                    orgSKU = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("盘点数量")]
            public decimal _数量 { get; set; }
        }

        [ExcelTable("区域负责人信息")]
        class _区域负责人Mapping
        {
            private string orgManager;
            private string orgArea;

            [ExcelColumn("盘点人员")]
            public string _负责人
            {
                get
                {
                    return orgManager;
                }
                set
                {
                    orgManager = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("负责区域")]
            public string _负责区域
            {
                get
                {
                    return orgArea;
                }
                set
                {
                    orgArea = value != null ? value.ToString().Trim() : "";
                }
            }
        }

        class _分析结果Model
        {
            public string _SKU { get; set; }
            public string _商品名称 { get; set; }
            public string _单位 { get; set; }
            public string _库位 { get; set; }
            public decimal _可用数量 { get; set; }
            public decimal _盘点数量 { get; set; }
            public decimal _占用数量 { get; set; }
            public decimal _库存数量 { get; set; }
            public decimal _遗留上海天数 { get; set; }
            public decimal _日平均销量 { get; set; }
            public bool _是否盘点 { get; set; }
            public string _负责人 { get; set; }
            public string _仓库
            {
                get
                {
                    var tmp = "";
                    if (!string.IsNullOrEmpty(_库位))
                    {
                        tmp = _库位.Substring(0, 2);
                    }
                    return tmp;
                }
            }
            public string _货架
            {
                get
                {
                    var tmp = "";
                    if (!string.IsNullOrEmpty(_库位))
                    {
                        tmp = _库位.Substring(0, 3);
                    }
                    return tmp;
                }
            }
            public string _归属
            {
                get
                {
                    var tmp = "";
                    if (!string.IsNullOrEmpty(_库位))
                    {
                        tmp = _库位.Substring(0, 1);
                    }
                    return tmp;
                }
            }
            //public decimal _移往昆山数量
            //{
            //    get
            //    {
            //        decimal tmp = 0;
            //        if (_盘点数量 != 0)
            //        {
            //            tmp = _盘点数量 - _留在上海数量;
            //        }
            //        else
            //        {
            //            tmp = _可用数量 - _留在上海数量;
            //        }
            //        return tmp;
            //    }
            //}

            //public decimal _留在上海数量
            //{
            //    get
            //    {
            //        return Math.Round(_遗留上海天数 * _日平均销量, 0);
            //    }
            //}
        }

    }
}
