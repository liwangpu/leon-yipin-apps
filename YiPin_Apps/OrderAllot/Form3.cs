using LinqToExcel;
using OfficeOpenXml;
using OrderAllot.Entities;
using OrderAllot.Libs;
using OrderAllot.Maps;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OrderAllot
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

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
                var buyersProviders = new List<BuyersProvider>();
                var providers = new List<string>();//供应商唯一队列
                var outBuyersProviders = new List<BuyersProvider>();

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
                                var bps = from c in excel.Worksheet<BuyersProvider>(s)
                                          select c;
                                buyersProviders.AddRange(bps.ToList().Where(b => !string.IsNullOrEmpty(b._采购) && !string.IsNullOrEmpty(b._供应商)));
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                        providers = buyersProviders.Select(b => b._供应商).Distinct().ToList();
                    }
                });


                actRead.BeginInvoke((obj) =>
                {
                    ShowMsg("开始计算表格数据");

                    var relBuyers = Helper.GetBuyers();
                    providers.ForEach(bp =>
                    {
                        var refProvider = buyersProviders.Where(pp => pp._供应商 == bp).ToList();
                        var outItem = new BuyersProvider();
                        outItem._供应商 = bp;
                        for (int idx = refProvider.Count() - 1; idx >= 0; idx--)
                        {
                            var curItem = refProvider[idx];
                            //没有对应采购
                            if (relBuyers.Where(xx => xx == curItem._采购).Count() == 0)
                            {
                                refProvider.RemoveAt(idx);
                            }
                        }

                        try
                        {
                            var refCount = refProvider.Count();
                            if (refCount > 0)
                            {
                                var refBuyer = refProvider.Select(xx => xx._采购).Distinct();
                                var maxSku = refProvider.Max(mm => mm._SKU数量);
                                outItem._采购 = refProvider.Where(mm => mm._SKU数量 == maxSku).Select(mm => mm._采购).First();
                                outItem._SKU数量 = maxSku;
                                outItem._有几个采购 = refCount;
                                outBuyersProviders.Add(outItem);
                            }
                        }
                        catch (Exception ex)
                        {
                            ShowMsg(ex.Message);
                        }
                    });

                    //计算完毕,开始导出数据
                    ExportExcel(outBuyersProviders);

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
        /// <param name="buyerProvider"></param>
        private void ExportExcel(List<BuyersProvider> buyerProvider)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[1, 1].Value = "供应商";
                sheet1.Cells[1, 2].Value = "采购员";
                sheet1.Cells[1, 3].Value = "SKU数量";
                sheet1.Cells[1, 4].Value = "有几个采购";
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = buyerProvider.Count; idx < len; idx++, rowIdx++)
                {
                    var curBuyerItem = buyerProvider[idx];
                    sheet1.Cells[rowIdx, 1].Value = curBuyerItem._供应商;
                    sheet1.Cells[rowIdx, 2].Value = curBuyerItem._采购;
                    sheet1.Cells[rowIdx, 3].Value = curBuyerItem._SKU数量;
                    sheet1.Cells[rowIdx, 4].Value = curBuyerItem._有几个采购;
                }
                #endregion


                buffer = package.GetAsByteArray();
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
                    string FileName = saveFile.FileName;//得到文件路径   
                    txtExport.Text = FileName;
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
