using LinqToExcel;
using OfficeOpenXml;
using OrderAllot.Entities;
using OrderAllot.Maps;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OrderAllot
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region Form1_Load
        private void Form1_Load(object sender, EventArgs e)
        {
            NtxtAmount.Value = 100;
        }
        #endregion

        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpload.Text = OpenFileDialog1.FileName;
                btnAnalyze.Enabled = true;
            }
        }

        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            try
            {
                #region 解析并计算
                var diviAmount = Convert.ToDouble(NtxtAmount.Value);
                var warningList = new List<Warning>();
                var orderList = new List<Order>();
                var providers = new List<string>();//供应商唯一队列
                var excelPath = txtUpload.Text;
                var actRead = new Action(() =>
                {
                    ShowMsg("开始读取表格数据");
                    using (var excel = new ExcelQueryFactory(excelPath))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<Warning>(s)
                                          select c;
                                warningList.AddRange(tmp);
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
                    //供应商唯一取值
                    providers = warningList.Select(p => p._供应商).Where(p => !string.IsNullOrEmpty(p)).Distinct().OrderBy(p => p).ToList();
                    //计算供应商采购金额
                    providers.ForEach(pd =>
                    {
                        var curProviderSku = warningList.Where(w => w._供应商 == pd).ToList();
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
                                curOrder._采购员 = ChangeLowerBuyer(sk._采购员);
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
                    ExportExcel(orderList);

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
        private void ExportExcel(List<Order> orders)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer2 = new byte[0];
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
                for (int idx = 0, rowIdx = 2, len = orders.Count; idx < len; idx++, rowIdx++)
                {
                    var curOrder = orders[idx];
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


                buffer = package.GetAsByteArray();
            }

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
                    var FileName2 = Path.Combine(savePath, saveFilName+"工作量.xlsx");


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

        #region ChangeLowerBuyer 采购员转换
        /// <summary>
        /// 采购员转换
        /// </summary>
        /// <param name="orgBuyerName"></param>
        /// <returns></returns>
        private string ChangeLowerBuyer(string orgBuyerName)
        {
            var newBuyerName = orgBuyerName;
            switch (orgBuyerName)
            {
                case "毕玉":
                    newBuyerName = "李曼曼";
                    break;
                case "鲍祝平":
                    newBuyerName = "王素素";
                    break;
                case "黄妍妍":
                    newBuyerName = "曹晨晨";
                    break;
                case "潘明媛":
                    newBuyerName = "侯春喜";
                    break;
                case "章玲玲":
                    newBuyerName = "董文丽";
                    break;
                case "蔡怡雯":
                    newBuyerName = "崔侠梅";
                    break;
                case "邹晓玲":
                    newBuyerName = "苏苗雨";
                    break;
                case "王思雅":
                    newBuyerName = "韦秋菊";
                    break;
                default:
                    break;
            }
            return newBuyerName;
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
