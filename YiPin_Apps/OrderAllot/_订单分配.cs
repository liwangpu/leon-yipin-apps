using OrderAllot.Entities;
using OrderAllot.Libs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace OrderAllot
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

                #region 解析并计算
                var strHeaders = new List<string>();
                strHeaders.Add("供应商");
                strHeaders.Add("SKU码");
                strHeaders.Add("建议采购数量");
                strHeaders.Add("仓库");
                strHeaders.Add("库存上限");
                strHeaders.Add("库存下限");
                strHeaders.Add("可用数量");
                strHeaders.Add("采购未入库");
                strHeaders.Add("缺货及未派单数量");
                strHeaders.Add("采购员");
                strHeaders.Add("商品成本单价");
                var strProperties = new List<string>();
                strProperties.Add("_供应商");
                strProperties.Add("_SKU");
                strProperties.Add("_建议采购数量");
                strProperties.Add("_仓库");
                strProperties.Add("_库存上限");
                strProperties.Add("_库存下限");
                strProperties.Add("_可用数量");
                strProperties.Add("_采购未入库");
                strProperties.Add("_缺货及未派单数量");
                strProperties.Add("_采购员");
                strProperties.Add("_商品成本单价");


                var dDiviAmount = Convert.ToDouble(NtxtAmount.Value);
                var strImExcelPath = txtUpload.Text;
                var imDataList = new List<IM订单分配>();
                var exDataList_采购 = new List<EX订单分配>();
                var exDataList_开发 = new List<EX订单分配>();
                var exWorkLoad = new List<EX工作量>();
                var providerNames = new List<string>();//供应商唯一队列

                var actRead = new Action(() =>
                {
                    ShowMsg("开始读取表格数据");
                    using (var fs = new FileStream(strImExcelPath, FileMode.Open))
                    {
                        imDataList = XlsxHelper.Read<IM订单分配>(fs, strHeaders, strProperties, out strOpMsg);
                    }
                });

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
                                exDataList_采购.Add(curOrder);
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
                                exDataList_采购.Add(curOrder);
                            });
                        }
                    });

                    //二次处理数据
                    var buffer1 = new byte[0];
                    {
                        var sefStrHeaders = new List<string>();
                        var sefStrProperties = new List<string>();
                        //


                    }




                    //计算完毕,开始导出数据
                    ExportExcel(exDataList_采购, exDataList_开发, exWorkLoad);

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


        private void ExportExcel(List<EX订单分配> _采购订单, List<EX订单分配> _开发订单, List<EX工作量> _工作量)
        {
            var buyerBuffer = new byte[0];
            #region 采购员订单
            {
                var strOpMsg = string.Empty;
                var strHeaders = new List<string>();
                strHeaders.Add("供应商");
                strHeaders.Add("SKU");
                strHeaders.Add("Qty");
                strHeaders.Add("仓库");
                strHeaders.Add("备注");
                strHeaders.Add("合同号");
                strHeaders.Add("采购员");
                strHeaders.Add("含税单价");
                strHeaders.Add("物流费");
                strHeaders.Add("付款方式");
                strHeaders.Add("制单人");
                strHeaders.Add("到货日期");
                strHeaders.Add("1688单号");
                strHeaders.Add("预付款");

                var strProperties = new List<string>();
                strProperties.Add("供应商");
                strProperties.Add("_SKU");
                strProperties.Add("_Qty");
                strProperties.Add("_仓库");
                strProperties.Add("_备注");
                strProperties.Add("_合同号");
                strProperties.Add("_采购员");
                strProperties.Add("_含税单价");
                strProperties.Add("_物流费");
                strProperties.Add("_付款方式");
                strProperties.Add("_制单人");
                strProperties.Add("_到货日期");
                strProperties.Add("_1688单号");
                strProperties.Add("_预付款");
                strProperties.Add("_对应供应商采购金额");

                buyerBuffer = XlsxHelper.SimpleWrite<EX订单分配>(_采购订单, strHeaders, strProperties, out strOpMsg);
            }
            #endregion

            var devBuffer = new byte[0];
            #region 开发订单
            {
                var strOpMsg = string.Empty;
                var strHeaders = new List<string>();
                strHeaders.Add("供应商");
                strHeaders.Add("SKU");
                strHeaders.Add("Qty");
                strHeaders.Add("仓库");
                strHeaders.Add("备注");
                strHeaders.Add("合同号");
                strHeaders.Add("采购员");
                strHeaders.Add("含税单价");
                strHeaders.Add("物流费");
                strHeaders.Add("付款方式");
                strHeaders.Add("制单人");
                strHeaders.Add("到货日期");
                strHeaders.Add("1688单号");
                strHeaders.Add("预付款");

                var strProperties = new List<string>();
                strProperties.Add("供应商");
                strProperties.Add("_SKU");
                strProperties.Add("_Qty");
                strProperties.Add("_仓库");
                strProperties.Add("_备注");
                strProperties.Add("_合同号");
                strProperties.Add("_采购员");
                strProperties.Add("_含税单价");
                strProperties.Add("_物流费");
                strProperties.Add("_付款方式");
                strProperties.Add("_制单人");
                strProperties.Add("_到货日期");
                strProperties.Add("_1688单号");
                strProperties.Add("_预付款");
                strProperties.Add("_对应供应商采购金额");

                devBuffer = XlsxHelper.SimpleWrite<EX订单分配>(_开发订单, strHeaders, strProperties, out strOpMsg);
            }
            #endregion

            var workloadBuffer = new byte[0];
            { 
            
            }

        }
    }
}
