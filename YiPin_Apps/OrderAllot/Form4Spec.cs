using LinqToExcel;
using OfficeOpenXml;
using OrderAllot.Entities;
using OrderAllot.Maps;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OrderAllot.Libs;


namespace OrderAllot
{
    public partial class Form4Spec : Form
    {
        public Form4Spec()
        {
            InitializeComponent();
        }

        #region 上传默认昆山预警订单
        private void btnUpDfkunsYj_Click(object sender, EventArgs e)
        {
            //上传默认昆山预警订单
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpDfkunsYj.Text = OpenFileDialog1.FileName;
                btnAnalyze.Enabled = true;
            }
        }
        #endregion

        #region 上传昆山采购建议
        private void btnUpKsYj_Click(object sender, EventArgs e)
        {
            //昆山采购建议
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpKsYj.Text = OpenFileDialog1.FileName;
                btnAnalyze.Enabled = true;
            }
        }
        #endregion

        #region 上传昆山所有库存
        private void btnKsKc_Click(object sender, EventArgs e)
        {
            //昆山所有库存
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpKsKc.Text = OpenFileDialog1.FileName;
                btnAnalyze.Enabled = true;
            }
        }
        #endregion

        #region 上传上海所有库存
        private void btnUpSHKc_Click(object sender, EventArgs e)
        {
            //上海所有库存
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpSHKc.Text = OpenFileDialog1.FileName;
            }
        }
        #endregion

        #region 上传临时备货
        private void btnUpTmp_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpTmp.Text = OpenFileDialog1.FileName;
            }
        }
        #endregion

        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            try
            {
                #region 解析并计算
                var diviAmount = Convert.ToDouble(NtxtAmount.Value);
                var dfKunsWarnings = new List<Warning>();
                var KunsWarnings = new List<Warning>();
                var KunsStoreWarnings = new List<Warning>();
                var ShanghStoreWarnings = new List<Warning>();
                var notdfKunsWarnings = new List<Warning>();//上海默认昆山预警,昆山不预警,但是昆山库存够卖两个仓库
                var tmpWarnings = new List<Order>();
                //var dfKunsWarnings = new List<Warning>();
                var orderList = new List<Order>();
                var devList = new List<Order>();//把开发单独分写成一个表格 
                var providers = new List<string>();//供应商唯一队列
                var dfKunsWarningPath = txtUpDfkunsYj.Text;
                var KunsWarningPath = txtUpKsYj.Text;
                var KunsStoreWarningPath = txtUpKsKc.Text;
                var ShanghStoreWarningPath = txtUpSHKc.Text;
                var tmpWarningPath = txtUpTmp.Text;
                var actRead = new Action(() =>
                {
                    ShowMsg("开始读取表格数据");
                    //上海默认昆山预警
                    using (var excel = new ExcelQueryFactory(dfKunsWarningPath))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<Warning>(s)
                                          select c;
                                dfKunsWarnings.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                    }
                    //昆山预警
                    using (var excel = new ExcelQueryFactory(KunsWarningPath))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<Warning>(s)
                                          select c;
                                KunsWarnings.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                    }
                    //昆山仓库
                    using (var excel = new ExcelQueryFactory(KunsStoreWarningPath))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<Warning>(s)
                                          select c;
                                KunsStoreWarnings.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                    }
                    //上海仓库
                    using (var excel = new ExcelQueryFactory(ShanghStoreWarningPath))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<Warning>(s)
                                          select c;
                                ShanghStoreWarnings.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                    }
                    //临时备货
                    using (var excel = new ExcelQueryFactory(tmpWarningPath))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<Order>(s)
                                          select c;
                                tmpWarnings.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                    }
                });

                actRead.BeginInvoke((obj) =>
                {
                    ShowMsg("开始计算表格数据");
                    //1.遍历上海默认昆山预警,把昆山预警里面量表都有的sku相关参数加起来
                    //2.没有的sku进入库存列表进行是否够卖判断,够卖不用采购
                    for (int idx = dfKunsWarnings.Count - 1; idx >= 0; idx--)
                    {
                        var df = dfKunsWarnings[idx];
                        var refKunsItem = KunsWarnings.Where(r => r._SKU == df._SKU).FirstOrDefault();
                        //有对应sku,相关参数加起来
                        if (refKunsItem != null)
                        {
                            df._采购未入库 += refKunsItem._采购未入库;
                            df._建议采购数量 += refKunsItem._建议采购数量;
                            df._可用数量 += refKunsItem._可用数量;
                            df._库存上限 += refKunsItem._库存上限;
                            df._库存下限 += refKunsItem._库存下限;
                            df._缺货及未派单数量 += refKunsItem._缺货及未派单数量;
                        }
                        else
                        {
                            //没有对应sku,判断够不够卖
                            var refKunsStoreItem = KunsStoreWarnings.Where(r => r._SKU == df._SKU).FirstOrDefault();
                            if (refKunsStoreItem != null)
                            {
                                //库存足够,删掉,但是删掉之前加入特殊列表
                                if (df._特殊查看是否够卖 + refKunsStoreItem._特殊查看是否够卖 < 0)
                                {
                                    df._特殊最终库存多余数量 = -(df._特殊查看是否够卖 + refKunsStoreItem._特殊查看是否够卖);
                                    notdfKunsWarnings.Add(df);
                                    dfKunsWarnings.RemoveAt(idx);
                                }
                            }
                        }
                    }
                    //3.将昆山原来的预警进行sku排除后加入上海默认昆山预警进入分配给采购
                    KunsWarnings.ForEach(ks =>
                    {
                        var isHas = dfKunsWarnings.Count(df => df._SKU == ks._SKU) > 0;
                        if (!isHas)
                        {
                            //在昆山预警里面,不是两个预警公共部分,需要到上海库存计算是否够卖
                            var refShanghStoreItem = ShanghStoreWarnings.Where(r => r._SKU == ks._SKU).FirstOrDefault();
                            if (refShanghStoreItem != null)
                            {
                                //库存足够,删掉,但是删掉之前加入特殊列表
                                if (ks._特殊查看是否够卖 + refShanghStoreItem._特殊查看是否够卖 < 0)
                                {
                                    ks._特殊最终库存多余数量 = -(ks._特殊查看是否够卖 + refShanghStoreItem._特殊查看是否够卖);
                                    if (ks._特殊查看是否够卖 + refShanghStoreItem._特殊查看是否够卖 > 0)
                                        dfKunsWarnings.Add(ks);
                                }
                            }


                        }
                    });

                    //供应商唯一取值
                    providers = dfKunsWarnings.Select(p => p._供应商).Where(p => !string.IsNullOrEmpty(p)).Distinct().OrderBy(p => p).ToList();
                    //计算供应商采购金额
                    providers.ForEach(pd =>
                    {
                        var curProviderSku = dfKunsWarnings.Where(w => w._供应商 == pd).ToList();
                        var thisProviderAmount = curProviderSku.Select(c => c._采购金额).Sum();
                        //小于分界,分给合肥
                        if (thisProviderAmount <= diviAmount)
                        {
                            curProviderSku.ForEach(sk =>
                            {
                                var curOrder = new Order();
                                curOrder._供应商 = pd;
                                curOrder._SKU = sk._SKU;
                                curOrder._Qty = sk._建议采购数量;
                                curOrder._采购员 = Helper.ChangeLowerBuyer(sk._采购员);
                                curOrder._含税单价 = sk._商品成本单价;
                                curOrder._制单人 = sk._采购员;
                                curOrder._对应供应商采购金额 = thisProviderAmount;
                                orderList.Add(curOrder);
                            });
                        }
                        else
                        {
                            //大于分界,保存不变
                            curProviderSku.ForEach(sk =>
                            {
                                var curOrder = new Order();
                                curOrder._供应商 = pd;
                                curOrder._SKU = sk._SKU;
                                curOrder._Qty = sk._建议采购数量;
                                curOrder._采购员 = sk._采购员;
                                curOrder._含税单价 = sk._商品成本单价;
                                curOrder._制单人 = curOrder._采购员;
                                curOrder._对应供应商采购金额 = thisProviderAmount;
                                orderList.Add(curOrder);
                            });
                        }
                    });

                    //计算完毕,开始导出数据
                    ExportExcel(orderList, notdfKunsWarnings, tmpWarnings, diviAmount);

                }, null);
                #endregion
            }
            catch (Exception ex)
            {
                ShowMsg(ex.Message);
            }
        }

        #region ExportExcel 导出Excel表格
        /// <summary>
        /// 导出Excel表格
        /// </summary>
        /// <param name="orders"></param>
        private void ExportExcel(List<Order> orders, List<Warning> notBuyWarnings, List<Order> tmp, double diviAmount)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer2 = new byte[0];
            var buffer3 = new byte[0];
            var buffer4 = new byte[0];
            var devOrder = new List<Order>();


            #region 订单分配
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                for (int idx = 0, len = tmp.Count; idx < len; idx++)
                {
                    var curTmp = tmp[idx];
                    curTmp._制单人 = curTmp._采购员;
                    curTmp._Qty = Helper.CalAmount(curTmp._Qty);
                }
                orders.AddRange(tmp);



                //重新计算大小单并转化采购员
                var providers = orders.Select(m => m._供应商).Distinct().ToList();
                providers.ForEach(pr =>
                {
                    var curDiv = orders.Where(p => p._供应商 == pr).Select(p => p._tmp采购总金额).Sum();
                    if (curDiv <= diviAmount)
                    {
                        //小于分界,分给合肥
                        for (int ddx = 0,len=orders.Count; ddx < len; ddx++)
                        {
                            var curOrder = orders[ddx];
                            curOrder._采购员 = Helper.ChangeLowerBuyer(curOrder._采购员);
                        }
                    }
                });

                //var newOrders = orders.OrderBy(mm => mm._供应商).ToList();

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

            #region 订单分配(开发单独一张表,其实是从订单分配分出来的)
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[1, 1].Value = "SKU";
                sheet1.Cells[1, 2].Value = "供应商";
                sheet1.Cells[1, 3].Value = "采购员";
                sheet1.Cells[1, 4].Value = "多余数量";
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = notBuyWarnings.Count; idx < len; idx++, rowIdx++)
                {
                    var curOrder = notBuyWarnings[idx];
                    sheet1.Cells[rowIdx, 1].Value = curOrder._SKU;
                    sheet1.Cells[rowIdx, 2].Value = curOrder._供应商;
                    sheet1.Cells[rowIdx, 3].Value = curOrder._采购员;
                    sheet1.Cells[rowIdx, 4].Value = curOrder._特殊最终库存多余数量;

                }
                #endregion


                buffer4 = package.GetAsByteArray();
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

                    txtExport.Text = FileName;
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
                    btnAnalyze.Enabled = false;
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




    }
}
