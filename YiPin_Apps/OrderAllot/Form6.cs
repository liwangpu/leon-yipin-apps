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
using OfficeOpenXml.Style;

namespace OrderAllot
{
    public partial class Form6 : Form
    {
        /**************** properties ****************/

        #region 档次SKU个数参数

        private decimal v_D1_SKUAmount;
        /// <summary>
        /// 第1档SKU个数
        /// </summary>
        private decimal _D1_SKUAmount
        {
            get
            {
                return v_D1_SKUAmount;
            }
            set
            {
                ntxtD1.Value = value;
                v_D1_SKUAmount = value;
            }
        }

        private decimal v_D2_SKUAmount;
        /// <summary>
        /// 第2档SKU个数
        /// </summary>
        private decimal _D2_SKUAmount
        {
            get
            {
                return v_D2_SKUAmount;
            }
            set
            {
                ntxtD2.Value = value;
                v_D2_SKUAmount = value;
            }
        }

        private decimal v_D3_SKUAmount;
        /// <summary>
        /// 第3档SKU个数
        /// </summary>
        private decimal _D3_SKUAmount
        {
            get
            {
                return v_D3_SKUAmount;
            }
            set
            {
                ntxtD3.Value = value;
                v_D3_SKUAmount = value;
            }
        }

        private decimal v_D4_SKUAmount;
        /// <summary>
        /// 第4档SKU个数
        /// </summary>
        private decimal _D4_SKUAmount
        {
            get
            {
                return v_D4_SKUAmount;
            }
            set
            {
                ntxtD4.Value = value;
                v_D4_SKUAmount = value;
            }
        }

        private decimal v_D5_SKUAmount;
        /// <summary>
        /// 第5档SKU个数
        /// </summary>
        private decimal _D5_SKUAmount
        {
            get
            {
                return v_D5_SKUAmount;
            }
            set
            {
                ntxtD5.Value = value;
                v_D5_SKUAmount = value;
            }
        }
        #endregion

        #region 档次奖励基数参数

        private decimal v_D1_Base;
        /// <summary>
        /// 第1档奖励基数
        /// </summary>
        private decimal _D1_Base
        {
            get
            {
                return v_D1_Base;
            }
            set
            {
                v_D1_Base = value;
            }
        }

        private decimal v_D2_Base;
        /// <summary>
        /// 第2档奖励基数
        /// </summary>
        private decimal _D2_Base
        {
            get
            {
                return v_D2_Base;
            }
            set
            {
                v_D2_Base = value;
            }
        }

        private decimal v_D3_Base;
        /// <summary>
        /// 第3档奖励基数
        /// </summary>
        private decimal _D3_Base
        {
            get
            {
                return v_D3_Base;
            }
            set
            {
                v_D3_Base = value;
            }
        }

        private decimal v_D4_Base;
        /// <summary>
        /// 第4档奖励基数
        /// </summary>
        private decimal _D4_Base
        {
            get
            {
                return v_D4_Base;
            }
            set
            {
                v_D4_Base = value;
            }
        }

        private decimal v_D5_Base;
        /// <summary>
        /// 第5档奖励基数
        /// </summary>
        private decimal _D5_Base
        {
            get
            {
                return v_D5_Base;
            }
            set
            {
                v_D5_Base = value;
            }
        }

        private decimal v_D_Base_Dif;
        /// <summary>
        /// 档次基数差值
        /// </summary>
        private decimal _D_Base_Dif
        {
            get
            {
                //return ntxtBaseDDif.Value;
                return v_D_Base_Dif;
            }
            set
            {
                v_D_Base_Dif = value;
            }
        }
        #endregion

        #region 阶段奖励参数

        #region 第1档各个阶段
        private decimal v_D1_J1;
        private decimal _D1_J1
        {
            get
            {
                return v_D1_J1;
            }
            set
            {
                lbD1_J1.Text = value.ToString();
                v_D1_J1 = value;
            }
        }

        private decimal v_D1_J2;
        private decimal _D1_J2
        {
            get
            {
                return v_D1_J2;
            }
            set
            {
                lbD1_J2.Text = value.ToString();
                v_D1_J2 = value;
            }
        }

        private decimal v_D1_J3;
        private decimal _D1_J3
        {
            get
            {
                return v_D1_J3;
            }
            set
            {
                lbD1_J3.Text = value.ToString();
                v_D1_J3 = value;
            }
        }

        private decimal v_D1_J4;
        private decimal _D1_J4
        {
            get
            {
                return v_D1_J4;
            }
            set
            {
                lbD1_J4.Text = value.ToString();
                v_D1_J4 = value;
            }
        }

        private decimal v_D1_J5;
        private decimal _D1_J5
        {
            get
            {
                return v_D1_J5;
            }
            set
            {
                lbD1_J5.Text = value.ToString();
                v_D1_J5 = value;
            }
        }

        private decimal v_D1_J1_Amount;
        private decimal _D1_J1_Amount
        {
            get
            {
                return v_D1_J1_Amount;
            }
            set
            {
                v_D1_J1_Amount = value;
                lbD1_J1_Amount.Text = value.ToString();
            }
        }

        private decimal v_D1_J2_Amount;
        private decimal _D1_J2_Amount
        {
            get
            {
                return v_D1_J2_Amount;
            }
            set
            {
                v_D1_J2_Amount = value;
                lbD1_J2_Amount.Text = value.ToString();
            }
        }

        private decimal v_D1_J3_Amount;
        private decimal _D1_J3_Amount
        {
            get
            {
                return v_D1_J3_Amount;
            }
            set
            {
                v_D1_J3_Amount = value;
                lbD1_J3_Amount.Text = value.ToString();
            }
        }

        private decimal v_D1_J4_Amount;
        private decimal _D1_J4_Amount
        {
            get
            {
                return v_D1_J4_Amount;
            }
            set
            {
                v_D1_J4_Amount = value;
                lbD1_J4_Amount.Text = value.ToString();
            }
        }

        private decimal v_D1_J5_Amount;
        private decimal _D1_J5_Amount
        {
            get
            {
                return v_D1_J5_Amount;
            }
            set
            {
                v_D1_J5_Amount = value;
                lbD1_J5_Amount.Text = value.ToString();
            }
        }
        #endregion

        #region 第2档各个阶段
        private decimal v_D2_J1;
        private decimal _D2_J1
        {
            get
            {
                return v_D2_J1;
            }
            set
            {
                lbD2_J1.Text = value.ToString();
                v_D2_J1 = value;
            }
        }

        private decimal v_D2_J2;
        private decimal _D2_J2
        {
            get
            {
                return v_D2_J2;
            }
            set
            {
                lbD2_J2.Text = value.ToString();
                v_D2_J2 = value;
            }
        }

        private decimal v_D2_J3;
        private decimal _D2_J3
        {
            get
            {
                return v_D2_J3;
            }
            set
            {
                lbD2_J3.Text = value.ToString();
                v_D2_J3 = value;
            }
        }

        private decimal v_D2_J4;
        private decimal _D2_J4
        {
            get
            {
                return v_D2_J4;
            }
            set
            {
                lbD2_J4.Text = value.ToString();
                v_D2_J4 = value;
            }
        }

        private decimal v_D2_J5;
        private decimal _D2_J5
        {
            get
            {
                return v_D2_J5;
            }
            set
            {
                lbD2_J5.Text = value.ToString();
                v_D2_J5 = value;
            }
        }

        private decimal v_D2_J1_Amount;
        private decimal _D2_J1_Amount
        {
            get
            {
                return v_D2_J1_Amount;
            }
            set
            {
                v_D2_J1_Amount = value;
                lbD2_J1_Amount.Text = value.ToString();
            }
        }

        private decimal v_D2_J2_Amount;
        private decimal _D2_J2_Amount
        {
            get
            {
                return v_D2_J2_Amount;
            }
            set
            {
                v_D2_J2_Amount = value;
                lbD2_J2_Amount.Text = value.ToString();
            }
        }

        private decimal v_D2_J3_Amount;
        private decimal _D2_J3_Amount
        {
            get
            {
                return v_D2_J3_Amount;
            }
            set
            {
                v_D2_J3_Amount = value;
                lbD2_J3_Amount.Text = value.ToString();
            }
        }

        private decimal v_D2_J4_Amount;
        private decimal _D2_J4_Amount
        {
            get
            {
                return v_D2_J4_Amount;
            }
            set
            {
                v_D2_J4_Amount = value;
                lbD2_J4_Amount.Text = value.ToString();
            }
        }

        private decimal v_D2_J5_Amount;
        private decimal _D2_J5_Amount
        {
            get
            {
                return v_D2_J5_Amount;
            }
            set
            {
                v_D2_J5_Amount = value;
                lbD2_J5_Amount.Text = value.ToString();
            }
        }
        #endregion

        #region 第3档各个阶段
        private decimal v_D3_J1;
        private decimal _D3_J1
        {
            get
            {
                return v_D3_J1;
            }
            set
            {
                lbD3_J1.Text = value.ToString();
                v_D3_J1 = value;
            }
        }

        private decimal v_D3_J2;
        private decimal _D3_J2
        {
            get
            {
                return v_D3_J2;
            }
            set
            {
                lbD3_J2.Text = value.ToString();
                v_D3_J2 = value;
            }
        }

        private decimal v_D3_J3;
        private decimal _D3_J3
        {
            get
            {
                return v_D3_J3;
            }
            set
            {
                lbD3_J3.Text = value.ToString();
                v_D3_J3 = value;
            }
        }

        private decimal v_D3_J4;
        private decimal _D3_J4
        {
            get
            {
                return v_D3_J4;
            }
            set
            {
                lbD3_J4.Text = value.ToString();
                v_D3_J4 = value;
            }
        }

        private decimal v_D3_J5;
        private decimal _D3_J5
        {
            get
            {
                return v_D3_J5;
            }
            set
            {
                lbD3_J5.Text = value.ToString();
                v_D3_J5 = value;
            }
        }

        private decimal v_D3_J1_Amount;
        private decimal _D3_J1_Amount
        {
            get
            {
                return v_D3_J1_Amount;
            }
            set
            {
                v_D3_J1_Amount = value;
                lbD3_J1_Amount.Text = value.ToString();
            }
        }

        private decimal v_D3_J2_Amount;
        private decimal _D3_J2_Amount
        {
            get
            {
                return v_D3_J2_Amount;
            }
            set
            {
                v_D3_J2_Amount = value;
                lbD3_J2_Amount.Text = value.ToString();
            }
        }

        private decimal v_D3_J3_Amount;
        private decimal _D3_J3_Amount
        {
            get
            {
                return v_D3_J3_Amount;
            }
            set
            {
                v_D3_J3_Amount = value;
                lbD3_J3_Amount.Text = value.ToString();
            }
        }

        private decimal v_D3_J4_Amount;
        private decimal _D3_J4_Amount
        {
            get
            {
                return v_D3_J4_Amount;
            }
            set
            {
                v_D3_J4_Amount = value;
                lbD3_J4_Amount.Text = value.ToString();
            }
        }

        private decimal v_D3_J5_Amount;
        private decimal _D3_J5_Amount
        {
            get
            {
                return v_D3_J5_Amount;
            }
            set
            {
                v_D3_J5_Amount = value;
                lbD3_J5_Amount.Text = value.ToString();
            }
        }
        #endregion

        #region 第4档各个阶段
        private decimal v_D4_J1;
        private decimal _D4_J1
        {
            get
            {
                return v_D4_J1;
            }
            set
            {
                lbD4_J1.Text = value.ToString();
                v_D4_J1 = value;
            }
        }

        private decimal v_D4_J2;
        private decimal _D4_J2
        {
            get
            {
                return v_D4_J2;
            }
            set
            {
                lbD4_J2.Text = value.ToString();
                v_D4_J2 = value;
            }
        }

        private decimal v_D4_J3;
        private decimal _D4_J3
        {
            get
            {
                return v_D4_J3;
            }
            set
            {
                lbD4_J3.Text = value.ToString();
                v_D4_J3 = value;
            }
        }

        private decimal v_D4_J4;
        private decimal _D4_J4
        {
            get
            {
                return v_D4_J4;
            }
            set
            {
                lbD4_J4.Text = value.ToString();
                v_D4_J4 = value;
            }
        }

        private decimal v_D4_J5;
        private decimal _D4_J5
        {
            get
            {
                return v_D4_J5;
            }
            set
            {
                lbD4_J5.Text = value.ToString();
                v_D4_J5 = value;
            }
        }

        private decimal v_D4_J1_Amount;
        private decimal _D4_J1_Amount
        {
            get
            {
                return v_D4_J1_Amount;
            }
            set
            {
                v_D4_J1_Amount = value;
                lbD4_J1_Amount.Text = value.ToString();
            }
        }

        private decimal v_D4_J2_Amount;
        private decimal _D4_J2_Amount
        {
            get
            {
                return v_D4_J2_Amount;
            }
            set
            {
                v_D4_J2_Amount = value;
                lbD4_J2_Amount.Text = value.ToString();
            }
        }

        private decimal v_D4_J3_Amount;
        private decimal _D4_J3_Amount
        {
            get
            {
                return v_D4_J3_Amount;
            }
            set
            {
                v_D4_J3_Amount = value;
                lbD4_J3_Amount.Text = value.ToString();
            }
        }

        private decimal v_D4_J4_Amount;
        private decimal _D4_J4_Amount
        {
            get
            {
                return v_D4_J4_Amount;
            }
            set
            {
                v_D4_J4_Amount = value;
                lbD4_J4_Amount.Text = value.ToString();
            }
        }

        private decimal v_D4_J5_Amount;
        private decimal _D4_J5_Amount
        {
            get
            {
                return v_D4_J5_Amount;
            }
            set
            {
                v_D4_J5_Amount = value;
                lbD4_J5_Amount.Text = value.ToString();
            }
        }
        #endregion

        #region 第5档各个阶段
        private decimal v_D5_J1;
        private decimal _D5_J1
        {
            get
            {
                return v_D5_J1;
            }
            set
            {
                lbD5_J1.Text = value.ToString();
                v_D5_J1 = value;
            }
        }

        private decimal v_D5_J2;
        private decimal _D5_J2
        {
            get
            {
                return v_D5_J2;
            }
            set
            {
                lbD5_J2.Text = value.ToString();
                v_D5_J2 = value;
            }
        }

        private decimal v_D5_J3;
        private decimal _D5_J3
        {
            get
            {
                return v_D5_J3;
            }
            set
            {
                lbD5_J3.Text = value.ToString();
                v_D5_J3 = value;
            }
        }

        private decimal v_D5_J4;
        private decimal _D5_J4
        {
            get
            {
                return v_D5_J4;
            }
            set
            {
                lbD5_J4.Text = value.ToString();
                v_D5_J4 = value;
            }
        }

        private decimal v_D5_J5;
        private decimal _D5_J5
        {
            get
            {
                return v_D5_J5;
            }
            set
            {
                lbD5_J5.Text = value.ToString();
                v_D5_J5 = value;
            }
        }

        private decimal v_D5_J1_Amount;
        private decimal _D5_J1_Amount
        {
            get
            {
                return v_D5_J1_Amount;
            }
            set
            {
                v_D5_J1_Amount = value;
                lbD5_J1_Amount.Text = value.ToString();
            }
        }

        private decimal v_D5_J2_Amount;
        private decimal _D5_J2_Amount
        {
            get
            {
                return v_D5_J2_Amount;
            }
            set
            {
                v_D5_J2_Amount = value;
                lbD5_J2_Amount.Text = value.ToString();
            }
        }

        private decimal v_D5_J3_Amount;
        private decimal _D5_J3_Amount
        {
            get
            {
                return v_D5_J3_Amount;
            }
            set
            {
                v_D5_J3_Amount = value;
                lbD5_J3_Amount.Text = value.ToString();
            }
        }

        private decimal v_D5_J4_Amount;
        private decimal _D5_J4_Amount
        {
            get
            {
                return v_D5_J4_Amount;
            }
            set
            {
                v_D5_J4_Amount = value;
                lbD5_J4_Amount.Text = value.ToString();
            }
        }

        private decimal v_D5_J5_Amount;
        private decimal _D5_J5_Amount
        {
            get
            {
                return v_D5_J5_Amount;
            }
            set
            {
                v_D5_J5_Amount = value;
                lbD5_J5_Amount.Text = value.ToString();
            }
        }
        #endregion



        private decimal _DJDiff1
        {
            get
            {
                return ntxtJDiff1.Value;
            }
        }
        private decimal _DJDiff2
        {
            get
            {
                return ntxtJDiff2.Value;
            }
        }
        private decimal _DJDiff3
        {
            get
            {
                return ntxtJDiff3.Value;
            }
        }
        private decimal _DJDiff4
        {
            get
            {
                return ntxtJDiff4.Value;
            }
        }
        #endregion

        private decimal _TransferRate;

        private List<_Form6采购流水Model> PurchaseList;

        public Form6()
        {
            InitializeComponent();
        }

        private void Form6_Load(object sender, EventArgs e)
        {
            PurchaseList = new List<_Form6采购流水Model>();
        }

        /**************** 按钮事件 ****************/

        #region 上传数据按钮事件
        private void btnParseSource_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtPurchaseOrg.Text = OpenFileDialog1.FileName;
                PurchaseList.Clear();

                #region 解析数据
                var actRead = new Action(() =>
                      {
                          ShowMsg("开始读取表格数据");

                          if (!string.IsNullOrEmpty(txtPurchaseOrg.Text))
                          {
                              using (var excel = new ExcelQueryFactory(txtPurchaseOrg.Text))
                              {
                                  var list = new List<_Form6采购流水>();
                                  var sheetNames = excel.GetWorksheetNames().ToList();
                                  sheetNames.ForEach(s =>
                                  {
                                      try
                                      {
                                          var tmp = from c in excel.Worksheet<_Form6采购流水>(s)
                                                    select c;
                                          list.AddRange(tmp);
                                      }
                                      catch (Exception ex)
                                      {
                                          ShowMsg(ex.Message);
                                      }
                                  });

                                  //为了解决效率,实体转换
                                  if (list.Count > 0)
                                  {
                                      list.ForEach(m =>
                                      {
                                          var entity = new _Form6采购流水Model();
                                          entity._采购员 = m._采购员;
                                          entity._制单人 = m._制单人;
                                          entity._SKU个数 = !string.IsNullOrEmpty(m._OrgSKU个数) ? Convert.ToInt32(m._OrgSKU个数) : 0;
                                          entity._总金额 = !string.IsNullOrEmpty(m._Org总金额) ? Convert.ToDecimal(m._Org总金额) : 0;
                                          PurchaseList.Add(entity);
                                      });
                                  }
                              }
                          }
                      });
                #endregion

                #region 解析完成后
                actRead.BeginInvoke((obj) =>
                {
                    ShowMsg("数据读取完毕");

                    var enableBtnAct = new Action(() =>
                    {
                        btnCalcu.Enabled = true;
                        btnRefreshConfig.Enabled = true;

                        //if (!cbSelfSKUAmount.Checked)
                        //{
                        //    var _订单内最多SKU个数 = PurchaseList.Max(x => x._SKU个数);
                        //    var _订单内平均SKU个数 = PurchaseList.Sum(x => x._SKU个数) / PurchaseList.Count;



                        //}
                        //var a = 0;
                    });
                    InvokeMainForm(enableBtnAct);

                }, null);
                #endregion

            }
        }
        #endregion

        #region 刷新参数按钮事件
        private void btnRefreshConfig_Click(object sender, EventArgs e)
        {
            RefreshConfig();
            RefreshAward();


        }
        #endregion

        #region 计算奖励按钮事件
        private void btnCalcu_Click(object sender, EventArgs e)
        {
            var result = new List<_订单奖励>();
            var act = new Action(() =>
            {
                ShowMsg("开始计算奖励");
                PurchaseList.ForEach(pur =>
                {
                    if (!string.IsNullOrEmpty(pur._采购员) && !string.IsNullOrEmpty(pur._制单人))
                    {
                        var curBuyer = result.Where(x => x._采购员 == pur._采购员).FirstOrDefault();
                        var ownBuyer = result.Where(x => x._采购员 == pur._制单人).FirstOrDefault();
                        var bAddBuyer = false;
                        var bAddOwnBuyer = false;
                        var bNeedTransfer = false;
                        if (curBuyer == null)
                        {
                            bAddBuyer = true;
                            curBuyer = new _订单奖励();
                            curBuyer._采购员 = pur._采购员;
                        }

                        if (ownBuyer == null)
                        {
                            bAddOwnBuyer = true;
                            ownBuyer = new _订单奖励();
                            ownBuyer._采购员 = pur._制单人;
                        }

                        if (pur._采购员 != pur._制单人)
                        {
                            bNeedTransfer = true;
                        }
                        #region 计算奖励
                        //第一档
                        if (pur._SKU个数 <= _D1_SKUAmount)
                        {
                            #region 档次计算
                            if (pur._总金额 <= _D1_J1_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D1_J1 += _D1_J1;
                                }
                                else
                                {
                                    curBuyer._D1_J1 += _D1_J1 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D1_J1 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D1_J1_Amount && pur._总金额 <= _D1_J2_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D1_J2 += _D1_J2;
                                }
                                else
                                {
                                    curBuyer._D1_J2 += _D1_J2 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D1_J2 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D1_J2_Amount && pur._总金额 <= _D1_J3_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D1_J3 += _D1_J3;
                                }
                                else
                                {
                                    curBuyer._D1_J3 += _D1_J3 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D1_J3 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D1_J3_Amount && pur._总金额 <= _D1_J4_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D1_J4 += _D1_J4;
                                }
                                else
                                {
                                    curBuyer._D1_J4 += _D1_J4 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D1_J4 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D1_J4_Amount && pur._总金额 <= _D1_J5_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D1_J5 += _D1_J5;
                                }
                                else
                                {
                                    curBuyer._D1_J5 += _D1_J5 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D1_J5 * (_TransferRate);
                                }
                            }
                            else { }
                            #endregion
                        }
                        else if (pur._SKU个数 > _D1_SKUAmount && pur._SKU个数 <= _D2_SKUAmount)
                        {
                            #region 档次计算
                            if (pur._总金额 <= _D2_J1_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D2_J1 += _D2_J1;
                                }
                                else
                                {
                                    curBuyer._D2_J1 += _D2_J1 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D2_J1 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D2_J1_Amount && pur._总金额 <= _D2_J2_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D2_J2 += _D2_J2;
                                }
                                else
                                {
                                    curBuyer._D2_J2 += _D2_J2 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D2_J2 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D2_J2_Amount && pur._总金额 <= _D2_J3_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D2_J3 += _D2_J3;
                                }
                                else
                                {
                                    curBuyer._D2_J3 += _D2_J3 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D2_J3 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D2_J3_Amount && pur._总金额 <= _D2_J4_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D2_J4 += _D2_J4;
                                }
                                else
                                {
                                    curBuyer._D2_J4 += _D2_J4 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D2_J4 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D2_J4_Amount && pur._总金额 <= _D2_J5_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D2_J5 += _D2_J5;
                                }
                                else
                                {
                                    curBuyer._D2_J5 += _D2_J5 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D2_J5 * (_TransferRate);
                                }
                            }
                            else { }
                            #endregion
                        }
                        else if (pur._SKU个数 > _D2_SKUAmount && pur._SKU个数 <= _D3_SKUAmount)
                        {
                            #region 档次计算
                            if (pur._总金额 <= _D2_J1_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D2_J1 += _D2_J1;
                                }
                                else
                                {
                                    curBuyer._D2_J1 += _D2_J1 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D2_J1 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D2_J1_Amount && pur._总金额 <= _D2_J2_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D2_J2 += _D2_J2;
                                }
                                else
                                {
                                    curBuyer._D2_J2 += _D2_J2 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D2_J2 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D2_J2_Amount && pur._总金额 <= _D2_J3_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D2_J3 += _D2_J3;
                                }
                                else
                                {
                                    curBuyer._D2_J3 += _D2_J3 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D2_J3 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D2_J3_Amount && pur._总金额 <= _D2_J4_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D2_J4 += _D2_J4;
                                }
                                else
                                {
                                    curBuyer._D2_J4 += _D2_J4 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D2_J4 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D2_J4_Amount && pur._总金额 <= _D2_J5_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D2_J5 += _D2_J5;
                                }
                                else
                                {
                                    curBuyer._D2_J5 += _D2_J5 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D2_J5 * (_TransferRate);
                                }
                            }
                            else { }
                            #endregion
                        }
                        else if (pur._SKU个数 > _D3_SKUAmount && pur._SKU个数 <= _D4_SKUAmount)
                        {
                            #region 档次计算
                            if (pur._总金额 <= _D4_J1_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D4_J1 += _D4_J1;
                                }
                                else
                                {
                                    curBuyer._D4_J1 += _D4_J1 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D4_J1 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D4_J1_Amount && pur._总金额 <= _D4_J2_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D4_J2 += _D4_J2;
                                }
                                else
                                {
                                    curBuyer._D4_J2 += _D4_J2 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D4_J2 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D4_J2_Amount && pur._总金额 <= _D4_J3_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D4_J3 += _D4_J3;
                                }
                                else
                                {
                                    curBuyer._D4_J3 += _D4_J3 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D4_J3 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D4_J3_Amount && pur._总金额 <= _D4_J4_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D4_J4 += _D4_J4;
                                }
                                else
                                {
                                    curBuyer._D4_J4 += _D4_J4 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D4_J4 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D4_J4_Amount && pur._总金额 <= _D4_J5_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D4_J5 += _D4_J5;
                                }
                                else
                                {
                                    curBuyer._D4_J5 += _D4_J5 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D4_J5 * (_TransferRate);
                                }
                            }
                            else { }
                            #endregion
                        }
                        else if (pur._SKU个数 > _D4_SKUAmount && pur._SKU个数 <= _D5_SKUAmount)
                        {
                            #region 档次计算
                            if (pur._总金额 <= _D5_J1_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D5_J1 += _D5_J1;
                                }
                                else
                                {
                                    curBuyer._D5_J1 += _D5_J1 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D5_J1 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D5_J1_Amount && pur._总金额 <= _D5_J2_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D5_J2 += _D5_J2;
                                }
                                else
                                {
                                    curBuyer._D5_J2 += _D5_J2 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D5_J2 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D5_J2_Amount && pur._总金额 <= _D5_J3_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D5_J3 += _D5_J3;
                                }
                                else
                                {
                                    curBuyer._D5_J3 += _D5_J3 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D5_J3 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D5_J3_Amount && pur._总金额 <= _D5_J4_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D5_J4 += _D5_J4;
                                }
                                else
                                {
                                    curBuyer._D5_J4 += _D5_J4 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D5_J4 * (_TransferRate);
                                }
                            }
                            else if (pur._总金额 > _D5_J4_Amount && pur._总金额 <= _D5_J5_Amount)
                            {
                                if (!bNeedTransfer)
                                {
                                    curBuyer._D5_J5 += _D5_J5;
                                }
                                else
                                {
                                    curBuyer._D5_J5 += _D5_J5 * (1 - _TransferRate);
                                    ownBuyer._产品归属奖励 += _D5_J5 * (_TransferRate);
                                }
                            }
                            else { }
                            #endregion
                        }
                        else { }
                        #endregion

                        if (bAddBuyer)
                            result.Add(curBuyer);
                        if (bAddOwnBuyer)
                        {
                            var isExistOwner = result.Where(x => x._采购员 == ownBuyer._采购员).FirstOrDefault();
                            if (isExistOwner != null)
                                isExistOwner._产品归属奖励 += ownBuyer._产品归属奖励;
                            else
                                result.Add(ownBuyer);
                        }
                    }
                });
            });

            act.BeginInvoke((obj) =>
            {
                ShowMsg("奖励计算完毕,准备导出");
                Export(result.OrderByDescending(x => x.IsBuyer).ToList());
            }, null);
        }
        #endregion

        /**************** common ****************/

        #region 刷新配置
        private void RefreshConfig()
        {
            _D_Base_Dif = ntxtBaseDDif.Value;

            _TransferRate = ntxtAward.Value != 0 ? ntxtAward.Value / 100 : 0;

            _D1_SKUAmount = ntxtD1.Value;
            _D2_SKUAmount = ntxtD2.Value;
            _D3_SKUAmount = ntxtD3.Value;
            _D4_SKUAmount = ntxtD4.Value;
            _D5_SKUAmount = ntxtD5.Value;

            #region 刷新档次基数差
            _D1_Base = ntxtBaseD1.Value;
            _D2_Base = ntxtBaseD2.Value;
            _D3_Base = ntxtBaseD3.Value;
            _D4_Base = ntxtBaseD4.Value;
            _D5_Base = ntxtBaseD5.Value;
            #endregion

            #region 第1档各个阶段
            _D1_J3 = _D1_Base;
            _D1_J2 = _D1_J3 - _DJDiff2;
            _D1_J1 = _D1_J2 - _DJDiff1;
            _D1_J4 = _D1_J3 + _DJDiff3;
            _D1_J5 = _D1_J4 + _DJDiff4;
            #endregion

            #region 第2档各个阶段
            _D2_J3 = _D2_Base;
            _D2_J2 = _D2_J3 - _DJDiff2;
            _D2_J1 = _D2_J2 - _DJDiff1;
            _D2_J4 = _D2_J3 + _DJDiff3;
            _D2_J5 = _D2_J4 + _DJDiff4;
            #endregion

            #region 第3档各个阶段
            _D3_J3 = _D3_Base;
            _D3_J2 = _D3_J3 - _DJDiff2;
            _D3_J1 = _D3_J2 - _DJDiff1;
            _D3_J4 = _D3_J3 + _DJDiff3;
            _D3_J5 = _D3_J4 + _DJDiff4;
            #endregion

            #region 第4档各个阶段
            _D4_J3 = _D4_Base;
            _D4_J2 = _D4_J3 - _DJDiff2;
            _D4_J1 = _D4_J2 - _DJDiff1;
            _D4_J4 = _D4_J3 + _DJDiff3;
            _D4_J5 = _D4_J4 + _DJDiff4;
            #endregion

            #region 第5档各个阶段
            _D5_J3 = _D5_Base;
            _D5_J2 = _D5_J3 - _DJDiff2;
            _D5_J1 = _D5_J2 - _DJDiff1;
            _D5_J4 = _D5_J3 + _DJDiff3;
            _D5_J5 = _D5_J4 + _DJDiff4;
            #endregion
        }
        #endregion

        #region 刷新奖励金额配置
        private void RefreshAward()
        {
            if (PurchaseList != null && PurchaseList.Count > 0)
            {
                var grade1Purchases = PurchaseList.Where(x => x._SKU个数 <= _D1_SKUAmount).ToList();
                {
                    //计算平均金额
                    _D1_J3_Amount = GetInteger(grade1Purchases.Sum(x => x._总金额) / grade1Purchases.Count);
                    var tmp = GetInteger(_D1_J3_Amount / 3);
                    _D1_J1_Amount = tmp;
                    _D1_J2_Amount = tmp * 2;
                    var maxAmount = grade1Purchases.Max(x => x._总金额);
                    _D1_J4_Amount = GetInteger((maxAmount - _D1_J3_Amount) / 2) + _D1_J3_Amount;
                    _D1_J5_Amount = maxAmount;
                }

                var grade2Purchases = PurchaseList.Where(x => x._SKU个数 > _D1_SKUAmount && x._SKU个数 <= _D2_SKUAmount).ToList();
                {
                    //计算平均金额
                    _D2_J3_Amount = GetInteger(grade2Purchases.Sum(x => x._总金额) / grade2Purchases.Count);
                    var tmp = GetInteger(_D2_J3_Amount / 3);
                    _D2_J1_Amount = tmp;
                    _D2_J2_Amount = tmp * 2;
                    var maxAmount = grade2Purchases.Max(x => x._总金额);
                    _D2_J4_Amount = GetInteger((maxAmount - _D2_J3_Amount) / 2) + _D2_J3_Amount;
                    _D2_J5_Amount = maxAmount;
                }

                var grade3Purchases = PurchaseList.Where(x => x._SKU个数 > _D2_SKUAmount && x._SKU个数 <= _D3_SKUAmount).ToList();
                {
                    //计算平均金额
                    _D3_J3_Amount = GetInteger(grade3Purchases.Sum(x => x._总金额) / grade3Purchases.Count);
                    var tmp = GetInteger(_D3_J3_Amount / 3);
                    _D3_J1_Amount = tmp;
                    _D3_J2_Amount = tmp * 2;
                    var maxAmount = grade3Purchases.Max(x => x._总金额);
                    _D3_J4_Amount = GetInteger((maxAmount - _D3_J3_Amount) / 2) + _D3_J3_Amount;
                    _D3_J5_Amount = maxAmount;
                }

                var grade4Purchases = PurchaseList.Where(x => x._SKU个数 > _D3_SKUAmount && x._SKU个数 <= _D4_SKUAmount).ToList();
                {
                    //计算平均金额
                    _D4_J3_Amount = GetInteger(grade4Purchases.Sum(x => x._总金额) / grade4Purchases.Count);
                    var tmp = GetInteger(_D4_J3_Amount / 3);
                    _D4_J1_Amount = tmp;
                    _D4_J2_Amount = tmp * 2;
                    var maxAmount = grade4Purchases.Max(x => x._总金额);
                    _D4_J4_Amount = GetInteger((maxAmount - _D4_J3_Amount) / 2) + _D4_J3_Amount;
                    _D4_J5_Amount = maxAmount;
                }

                var grade5Purchases = PurchaseList.Where(x => x._SKU个数 > _D4_SKUAmount && x._SKU个数 <= _D5_SKUAmount).ToList();
                {
                    //计算平均金额
                    _D5_J3_Amount = GetInteger(grade5Purchases.Sum(x => x._总金额) / grade5Purchases.Count);
                    var tmp = GetInteger(_D5_J3_Amount / 3);
                    _D5_J1_Amount = tmp;
                    _D5_J2_Amount = tmp * 2;
                    var maxAmount = grade5Purchases.Max(x => x._总金额);
                    _D5_J4_Amount = GetInteger((maxAmount - _D5_J3_Amount) / 2) + _D5_J3_Amount;
                    _D5_J5_Amount = maxAmount;
                }
            }
        }
        #endregion

        #region 导出结果表格
        private void Export(List<_订单奖励> list)
        {
            var buffer = new byte[0];

            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[2, 1].Value = "采购员";

                sheet1.Cells[2, 2].Value = "第一阶奖励";
                sheet1.Cells[2, 3].Value = "第二阶奖励";
                sheet1.Cells[2, 4].Value = "第三阶奖励";
                sheet1.Cells[2, 5].Value = "第四阶奖励";
                sheet1.Cells[2, 6].Value = "第五阶奖励";

                using (var range = sheet1.Cells[1, 2, 1, 6])
                {
                    range.Merge = true;
                    range.Value = "第一档";
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                }

                sheet1.Cells[2, 7].Value = "第一阶奖励";
                sheet1.Cells[2, 8].Value = "第二阶奖励";
                sheet1.Cells[2, 9].Value = "第三阶奖励";
                sheet1.Cells[2, 10].Value = "第四阶奖励";
                sheet1.Cells[2, 11].Value = "第五阶奖励";
                using (var range = sheet1.Cells[1, 7, 1, 11])
                {
                    range.Merge = true;
                    range.Value = "第二档";
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                }


                sheet1.Cells[2, 12].Value = "第一阶奖励";
                sheet1.Cells[2, 13].Value = "第二阶奖励";
                sheet1.Cells[2, 14].Value = "第三阶奖励";
                sheet1.Cells[2, 15].Value = "第四阶奖励";
                sheet1.Cells[2, 16].Value = "第五阶奖励";
                using (var range = sheet1.Cells[1, 12, 1, 16])
                {
                    range.Merge = true;
                    range.Value = "第三档";
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                }

                sheet1.Cells[2, 17].Value = "第一阶奖励";
                sheet1.Cells[2, 18].Value = "第二阶奖励";
                sheet1.Cells[2, 19].Value = "第三阶奖励";
                sheet1.Cells[2, 20].Value = "第四阶奖励";
                sheet1.Cells[2, 21].Value = "第五阶奖励";
                using (var range = sheet1.Cells[1, 17, 1, 21])
                {
                    range.Merge = true;
                    range.Value = "第四档";
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                }

                sheet1.Cells[2, 22].Value = "第一阶奖励";
                sheet1.Cells[2, 23].Value = "第二阶奖励";
                sheet1.Cells[2, 24].Value = "第三阶奖励";
                sheet1.Cells[2, 25].Value = "第四阶奖励";
                sheet1.Cells[2, 26].Value = "第五阶奖励";
                using (var range = sheet1.Cells[1, 22, 1, 26])
                {
                    range.Merge = true;
                    range.Value = "第五档";
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                }

                sheet1.Cells[2, 27].Value = "制单奖励";
                sheet1.Cells[2, 28].Value = "合计";
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 3, len = list.Count; idx < len; idx++)
                {
                    var curResult = list[idx];

                    sheet1.Cells[rowIdx, 1].Value = curResult._采购员;

                    sheet1.Cells[rowIdx, 2].Value = curResult._D1_J1;
                    sheet1.Cells[rowIdx, 3].Value = curResult._D1_J2;
                    sheet1.Cells[rowIdx, 4].Value = curResult._D1_J3;
                    sheet1.Cells[rowIdx, 5].Value = curResult._D1_J4;
                    sheet1.Cells[rowIdx, 6].Value = curResult._D1_J5;

                    sheet1.Cells[rowIdx, 7].Value = curResult._D2_J1;
                    sheet1.Cells[rowIdx, 8].Value = curResult._D2_J2;
                    sheet1.Cells[rowIdx, 9].Value = curResult._D2_J3;
                    sheet1.Cells[rowIdx, 10].Value = curResult._D2_J4;
                    sheet1.Cells[rowIdx, 11].Value = curResult._D2_J5;

                    sheet1.Cells[rowIdx, 12].Value = curResult._D3_J1;
                    sheet1.Cells[rowIdx, 13].Value = curResult._D3_J2;
                    sheet1.Cells[rowIdx, 14].Value = curResult._D3_J3;
                    sheet1.Cells[rowIdx, 15].Value = curResult._D3_J4;
                    sheet1.Cells[rowIdx, 16].Value = curResult._D3_J5;

                    sheet1.Cells[rowIdx, 17].Value = curResult._D4_J1;
                    sheet1.Cells[rowIdx, 18].Value = curResult._D4_J2;
                    sheet1.Cells[rowIdx, 19].Value = curResult._D4_J3;
                    sheet1.Cells[rowIdx, 20].Value = curResult._D4_J4;
                    sheet1.Cells[rowIdx, 21].Value = curResult._D4_J5;

                    sheet1.Cells[rowIdx, 22].Value = curResult._D5_J1;
                    sheet1.Cells[rowIdx, 23].Value = curResult._D5_J2;
                    sheet1.Cells[rowIdx, 24].Value = curResult._D5_J3;
                    sheet1.Cells[rowIdx, 25].Value = curResult._D5_J4;
                    sheet1.Cells[rowIdx, 26].Value = curResult._D5_J5;

                    sheet1.Cells[rowIdx, 27].Value = curResult._产品归属奖励;
                    sheet1.Cells[rowIdx, 28].Value = curResult._合计;

                    rowIdx++;

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
                    var FileName = saveFile.FileName;//得到文件路径   
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

        #region 保留小数点两位
        private static decimal GetInteger(decimal dou)
        {
            return Math.Round(dou, 2);
        }
        #endregion

    }
}
