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
    public partial class _排除侵权_订单分配 : Form
    {
        public _排除侵权_订单分配()
        {
            InitializeComponent();
        }

        private void _排除侵权_订单分配_Load(object sender, EventArgs e)
        {
            //txt销量差明细.Text = @"C:\Users\Leon\Desktop\销量差明细.csv";
            //txt建议备货明细.Text = @"C:\Users\Leon\Desktop\建议备货明细.csv";
            
        }

        /**************** button event ****************/

        #region 上传销量差明细
        private void btn销量差明细_Click(object sender, EventArgs e)
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
                    txt销量差明细.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传建议备货明细
        private void btn建议备货明细_Click(object sender, EventArgs e)
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
                    txt建议备货明细.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 处理按钮事件
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            btnAnalyze.Enabled = false;


            var list备货信息 = new List<_备货信息>();


            #region 读取数据
            var actReadData = new Action(() =>
            {
                ShowMsg("开始读取销量差明细数据");
                #region 读取销量差明细
                {
                    var strCsvPath = txt销量差明细.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_备货信息>()
                                          select c;
                                list备货信息.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                ShowMsg("开始读取建议备货数据");
                #region 读取建议备货
                {
                    var strCsvPath = txt建议备货明细.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_备货信息>()
                                          select c;
                                list备货信息.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion
            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                var list订单分配 = new List<_订单分配>();

                list备货信息.ForEach(cur备货Item =>
                {
                    var model = new _订单分配();
                    model._SKU = cur备货Item.SKU;
                    model._供应商 = cur备货Item._供应商;
                    model._含税单价 = cur备货Item._含税单价;
                    model._Qty = cur备货Item._建议备货数量;
                    model._采购员 = cur备货Item._采购员;
                    model._制单人 = cur备货Item._制单员;

                    if (cur备货Item._建议备货数量>0)
                    {
                        list订单分配.Add(model);
                    }
  
                });

                ExportExcel(list订单分配);
            }, null);
            #endregion
        }
        #endregion

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_备货信息));

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

        #region ExportExcel 导出Excel表格
        /// <summary>
        /// 导出Excel表格
        /// </summary>
        /// <param name="orders"></param>
        private void ExportExcel(List<_订单分配> orders)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer2 = new byte[0];
            var buffer3 = new byte[0];
            var buffer4 = new byte[0];
            var buffer5 = new byte[0];
            var devOrder = new List<_订单分配>();


            #region 订单分配
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[1, 1].Value = "供应商";
                sheet1.Cells[1, 2].Value = "SKU";
                sheet1.Cells[1, 3].Value = "Qty";
                sheet1.Cells[1, 4].Value = "仓库";
                sheet1.Cells[1, 5].Value = "备注";
                sheet1.Cells[1, 6].Value = "合同号";
                sheet1.Cells[1, 7].Value = "采购员";
                sheet1.Cells[1, 8].Value = "含税单价";
                sheet1.Cells[1, 9].Value = "物流费";
                sheet1.Cells[1, 10].Value = "付款方式";
                sheet1.Cells[1, 11].Value = "制单人";
                sheet1.Cells[1, 12].Value = "到货日期";
                sheet1.Cells[1, 13].Value = "1688单号";
                sheet1.Cells[1, 14].Value = "预付款";
                //sheet1.Cells[1, 15].Value = "对应供应商采购总金额";
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = orders.Count; idx < len; idx++)
                {
                    var curOrder = orders[idx];
                    if (Helper.IsBuyer(curOrder._制单人))
                    {

                        sheet1.Cells[rowIdx, 1].Value = curOrder._供应商;
                        sheet1.Cells[rowIdx, 2].Value = curOrder._SKU;
                        sheet1.Cells[rowIdx, 3].Value = curOrder._Qty;
                        sheet1.Cells[rowIdx, 7].Value = curOrder._采购员;
                        sheet1.Cells[rowIdx, 8].Value = curOrder._含税单价;
                        sheet1.Cells[rowIdx, 10].Value = "支付宝";
                        sheet1.Cells[rowIdx, 11].Value = curOrder._制单人;
                        //sheet1.Cells[rowIdx, 15].Value = curOrder._对应供应商采购金额;

                        rowIdx++;
                    }
                    else
                    {
                        devOrder.Add(curOrder);
                    }
                }
                #endregion


                buffer = package.GetAsByteArray();
            }
            #endregion

            #region 工作量单独表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[1, 1].Value = "采购员";
                sheet1.Cells[1, 2].Value = "订单量";
                #endregion

                #region 数据行
                var buyers = new List<string>();
                buyers = orders.Where(x => !string.IsNullOrEmpty(x._采购员)).Select(x => x._采购员).Distinct().ToList();
                for (int idx = 0, len = buyers.Count, rowIdx = 2; idx < len; idx++, rowIdx++)
                {
                    var curBuyerName = buyers[idx];
                    var refOrders = orders.Where(m => m._采购员 == curBuyerName).ToList();
                    var amount = refOrders.Select(m => m._供应商).Distinct().Count();

                    sheet1.Cells[rowIdx, 1].Value = curBuyerName;
                    sheet1.Cells[rowIdx, 2].Value = amount;
                }
                #endregion

                buffer2 = package.GetAsByteArray();
            }
            #endregion

            #region 订单分配(开发单独一张表,其实是从订单分配分出来的)
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[1, 1].Value = "供应商";
                sheet1.Cells[1, 2].Value = "SKU";
                sheet1.Cells[1, 3].Value = "Qty";
                sheet1.Cells[1, 4].Value = "仓库";
                sheet1.Cells[1, 5].Value = "备注";
                sheet1.Cells[1, 6].Value = "合同号";
                sheet1.Cells[1, 7].Value = "采购员";
                sheet1.Cells[1, 8].Value = "含税单价";
                sheet1.Cells[1, 9].Value = "物流费";
                sheet1.Cells[1, 10].Value = "付款方式";
                sheet1.Cells[1, 11].Value = "制单人";
                sheet1.Cells[1, 12].Value = "到货日期";
                sheet1.Cells[1, 13].Value = "1688单号";
                sheet1.Cells[1, 14].Value = "预付款";
                //sheet1.Cells[1, 15].Value = "对应供应商采购总金额";
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = devOrder.Count; idx < len; idx++, rowIdx++)
                {
                    var curOrder = devOrder[idx];
                    sheet1.Cells[rowIdx, 1].Value = curOrder._供应商;
                    sheet1.Cells[rowIdx, 2].Value = curOrder._SKU;
                    sheet1.Cells[rowIdx, 3].Value = curOrder._Qty;
                    sheet1.Cells[rowIdx, 7].Value = curOrder._采购员;
                    sheet1.Cells[rowIdx, 8].Value = curOrder._含税单价;
                    sheet1.Cells[rowIdx, 10].Value = "支付宝";
                    sheet1.Cells[rowIdx, 11].Value = curOrder._制单人;
                    //sheet1.Cells[rowIdx, 15].Value = curOrder._对应供应商采购金额;

                }
                #endregion


                buffer3 = package.GetAsByteArray();
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
                    var FileName2 = Path.Combine(savePath, saveFilName + "工作量.xlsx");
                    var FileName3 = Path.Combine(savePath, saveFilName + "(开发订单).xlsx");
                    var FileName4 = Path.Combine(savePath, saveFilName + "(多余数量).xlsx");
                    //var FileName5 = Path.Combine(savePath, saveFilName + "(判断两个库存).xlsx");

                    try
                    {
                        var len = buffer.Length;
                        using (var fs = File.Create(FileName, len))
                        {
                            fs.Write(buffer, 0, len);
                        }

                        var len2 = buffer2.Length;
                        using (var fs = File.Create(FileName2, len2))
                        {
                            fs.Write(buffer2, 0, len2);
                        }

                        var len3 = buffer3.Length;
                        using (var fs = File.Create(FileName3, len3))
                        {
                            fs.Write(buffer3, 0, len3);
                        }

                        var len4 = buffer4.Length;
                        using (var fs = File.Create(FileName4, len4))
                        {
                            fs.Write(buffer4, 0, len4);
                        }

                    }
                    catch (Exception ex)
                    {
                        ShowMsg(ex.Message);
                    }

                    ShowMsg("表格生成完毕");
                    btnAnalyze.Enabled = true;
                }
            }, null);
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

        [ExcelTable("销量差明细表/建议备货明细表")]
        class _备货信息
        {
            [ExcelColumn("SKU")]
            public string SKU { get; set; }

            [ExcelColumn("供应商")]
            public string _供应商 { get; set; }

            [ExcelColumn("采购员")]
            public string _采购员 { get; set; }

            [ExcelColumn("制单员")]
            public string _制单员 { get; set; }

            [ExcelColumn("含税单价")]
            public decimal _含税单价 { get; set; }

            [ExcelColumn("建议备货数量")]
            public decimal _建议备货数量 { get; set; }
        }

        class _订单分配
        {
            public string _供应商 { get; set; }
            public string _SKU { get; set; }
            public decimal _Qty { get; set; }
            public string _仓库 { get; set; }
            public string _备注 { get; set; }
            public string _合同号 { get; set; }
            public string _采购员 { get; set; }
            public decimal _含税单价 { get; set; }
            public double _物流费 { get; set; }
            public string _付款方式 { get; set; }
            public string _制单人 { get; set; }
            public string _到货日期 { get; set; }
            public string _1688单号 { get; set; }
            public double _预付款 { get; set; }
            public double _对应供应商采购金额 { get; set; }
        }


    }
}
