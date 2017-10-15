using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using YPApps.Entities;
using System.IO;
using YPApps.Libs;

namespace YPApps.Forms
{
    public partial class _订单分配 : FormCore
    {
        public _订单分配()
        {
            InitializeComponent();
        }

        /**************** button event ****************/

        #region 浏览按钮事件
        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpload.Text = OpenFileDialog1.FileName;
                btnAnalyze.Enabled = true;
            }
        }
        #endregion

        #region 处理按钮事件
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            try
            {
                var strOpMsg = string.Empty;
                var dDiviAmount = Convert.ToDouble(NtxtAmount.Value);
                var strImExcelPath = txtUpload.Text;
                var imDataList = new List<IM订单分配>();
                var exDataList_采购 = new List<EX订单分配>();
                var exDataList_开发 = new List<EX订单分配>();
                var exWorkLoad = new List<EX工作量>();
                var providerNames = new List<string>();//供应商唯一队列

                #region 解析表格数据
                var actRead = new Action(() =>
                {
                    ShowMsg("开始读取表格数据");
                    using (var fs = new FileStream(strImExcelPath, FileMode.Open))
                    {
                        imDataList = XlsxHelper.Read<IM订单分配>(fs, IM订单分配.GetMapping(), out strOpMsg);
                    }
                });
                #endregion

                #region 数据处理
                actRead.BeginInvoke((obj) =>
                {
                    ShowMsg("开始计算表格数据");
                    //供应商唯一取值
                    providerNames = imDataList.Select(p => p._供应商).Where(p => !string.IsNullOrEmpty(p)).Distinct().OrderBy(p => p).ToList();
                    //计算供应商采购金额
                    providerNames.ForEach(pd =>
                    {
                        var curProviderSku = imDataList.Where(w => w._供应商 == pd).ToList();
                        var thisProviderAmount = curProviderSku.Select(c => c._采购金额).Sum();
                        //小于分界,分给合肥
                        if (thisProviderAmount <= dDiviAmount)
                        {
                            curProviderSku.ForEach(sk =>
                            {
                                var curOrder = new EX订单分配();
                                curOrder._供应商 = pd;
                                curOrder._SKU = sk._SKU;
                                curOrder._Qty = sk._建议采购数量;
                                curOrder._采购员 = Helper.ChangeLowerBuyer(sk._采购员);
                                curOrder._含税单价 = sk._商品成本单价;
                                curOrder._制单人 = sk._采购员;
                                curOrder._对应供应商采购金额 = thisProviderAmount;
                                if (Helper.IsBuyer(sk._采购员))
                                    exDataList_采购.Add(curOrder);
                                else
                                    exDataList_开发.Add(curOrder);
                            });
                        }
                        else
                        {
                            //大于分界,保存不变
                            curProviderSku.ForEach(sk =>
                            {
                                var curOrder = new EX订单分配();
                                curOrder._供应商 = pd;
                                curOrder._SKU = sk._SKU;
                                curOrder._Qty = sk._建议采购数量;
                                curOrder._采购员 = sk._采购员;
                                curOrder._含税单价 = sk._商品成本单价;
                                curOrder._制单人 = curOrder._采购员;
                                curOrder._对应供应商采购金额 = thisProviderAmount;
                                if (Helper.IsBuyer(sk._采购员))
                                    exDataList_采购.Add(curOrder);
                                else
                                    exDataList_开发.Add(curOrder);
                            });
                        }
                    });



                    #region 导出数据
                    ExportExcel(exDataList_采购, exDataList_开发);
                    #endregion

                }, null);
                #endregion
            }
            catch (Exception ex)
            {
                ShowMsg(ex.Message);
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

        #region ExportExcel 导出表格
        private void ExportExcel(List<EX订单分配> buyerDataList, List<EX订单分配> devDataList)
        {
            var strOpMsg = string.Empty;
            var buyerBuffer = XlsxHelper.SimpleWrite(buyerDataList.Select(x => x.ToDictionary()).ToList(), out strOpMsg);
            var devBuffer = XlsxHelper.SimpleWrite(devDataList.Select(x => x.ToDictionary()).ToList(), out strOpMsg);

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

                    txtExport.Text = FileName;
                    try
                    {
                        XlsxHelper.SaveWorkBook(buyerBuffer, FileName, out strOpMsg);
                        XlsxHelper.SaveWorkBook(devBuffer, FileName3, out strOpMsg);

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

    }
}
