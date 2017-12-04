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
    public partial class _工资计算 : Form
    {
        public _工资计算()
        {
            InitializeComponent();
        }

        private void _工资计算_Load(object sender, EventArgs e)
        {
            PurchaseList = new List<_采购流水Model>();

            //txtPurchaseOrg.Text = @"C:\Users\Leon\Desktop\All.xlsx";
            //txt库存周转率.Text = @"C:\Users\Leon\Desktop\采购工资所需报表\库存周转率.xlsx";
            //txt缺货信息.Text = @"C:\Users\Leon\Desktop\采购工资所需报表\缺货率报表.xlsx";
            //txt未入库.Text = @"C:\Users\Leon\Desktop\采购工资所需报表\采购已审核未入库.xlsx";
            //txt议价奖励.Text = @"C:\Users\Leon\Desktop\采购工资所需报表\议价绩效.xlsx";
        }

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

        private List<_采购流水Model> PurchaseList;

        /**************** button event ****************/

        #region 上传采购流水
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
                            var list = new List<_采购流水Mapping>();
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<_采购流水Mapping>(s)
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
                                    var entity = new _采购流水Model();
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
                        //btnCalcuAward.Enabled = true;
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

        #region 刷新订单奖励参数
        private void btnRefreshConfig_Click(object sender, EventArgs e)
        {
            RefreshConfig();
            RefreshAward();
            ShowMsg("参数刷新完毕");
        }
        #endregion

        #region 计算最终工资
        private void btnCalcuFinal_Click(object sender, EventArgs e)
        {
            var _上海采购底薪 = nd上海采购底薪.Value;
            var _合肥采购底薪 = nd合肥采购底薪.Value;

            var _库存周转率Mappings = new List<_库存周转率Mapping>();
            var _缺货率Mappings = new List<_缺货率Mapping>();
            var _未入库Mappings = new List<_未入库Mapping>();
            var _议价奖励appings = new List<_议价奖励Mapping>();

            var _滞销金滞销权重List = new List<_滞销金滞销权重Model>();
            var _缺货率权重奖励List = new List<_缺货率权重Model>();
            var _未入库惩罚List = new List<_未入库Model>();
            var _工资详情List = new List<_工资详情Model>();
            var _订单奖励List = new List<_订单奖励Model>();
            var _议价奖励List = new List<_议价奖励Model>();

            var act读取源数据 = new Action(() =>
            {
                ShowMsg("开始读取库存周转率");

                #region 读取库存周转率
                {
                    var str库存周转率 = txt库存周转率.Text;
                    if (!string.IsNullOrEmpty(str库存周转率))
                    {
                        using (var excel = new ExcelQueryFactory(str库存周转率))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<_库存周转率Mapping>(s)
                                              select c;
                                    _库存周转率Mappings.AddRange(tmp.ToList());
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                }
                #endregion

                ShowMsg("开始读取缺货率");

                #region 读取缺货率
                {
                    var str缺货率 = txt缺货信息.Text;
                    if (!string.IsNullOrEmpty(str缺货率))
                    {
                        using (var excel = new ExcelQueryFactory(str缺货率))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<_缺货率Mapping>(s)
                                              select c;
                                    _缺货率Mappings.AddRange(tmp.ToList());
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                }
                #endregion

                ShowMsg("开始读取议价奖励");

                #region 读取议价奖励
                {
                    var str议价奖励 = txt议价奖励.Text;
                    if (!string.IsNullOrEmpty(str议价奖励))
                    {
                        using (var excel = new ExcelQueryFactory(str议价奖励))
                        {

                            try
                            {
                                var sheetNames = excel.GetWorksheetNames().ToList();
                                sheetNames.ForEach(s =>
                                {
                                    var tmp = from c in excel.Worksheet<_议价奖励Mapping>(s)
                                              select c;
                                    _议价奖励appings.AddRange(tmp.ToList());
                                });
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }

                        }
                    }
                }
                #endregion

                ShowMsg("开始读取未入库");

                #region 读取未入库
                {
                    var str未入库 = txt未入库.Text;
                    if (!string.IsNullOrEmpty(str未入库))
                    {
                        using (var excel = new ExcelQueryFactory(str未入库))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<_未入库Mapping>(s)
                                              select c;
                                    _未入库Mappings.AddRange(tmp.ToList());
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                }
                #endregion


            });

            act读取源数据.BeginInvoke((obj) =>
            {
                ShowMsg("开始计算工作量工资");

                #region 计算订单奖励
                {
                    PurchaseList.ForEach(pur =>
                    {
                        if (!string.IsNullOrEmpty(pur._采购员) && !string.IsNullOrEmpty(pur._制单人))
                        {
                            var curBuyer = _订单奖励List.Where(x => x._采购员 == pur._采购员).FirstOrDefault();
                            var ownBuyer = _订单奖励List.Where(x => x._采购员 == pur._制单人).FirstOrDefault();
                            var bAddBuyer = false;
                            var bAddOwnBuyer = false;
                            var bNeedTransfer = false;
                            if (curBuyer == null)
                            {
                                bAddBuyer = true;
                                curBuyer = new _订单奖励Model();
                                curBuyer._采购员 = pur._采购员;
                            }

                            if (ownBuyer == null)
                            {
                                bAddOwnBuyer = true;
                                ownBuyer = new _订单奖励Model();
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
                                _订单奖励List.Add(curBuyer);
                            if (bAddOwnBuyer)
                            {
                                var isExistOwner = _订单奖励List.Where(x => x._采购员 == ownBuyer._采购员).FirstOrDefault();
                                if (isExistOwner != null)
                                    isExistOwner._产品归属奖励 += ownBuyer._产品归属奖励;
                                else
                                    _订单奖励List.Add(ownBuyer);
                            }
                        }
                    });
                }
                #endregion

                ShowMsg("开始计算滞销金/滞销权重");

                #region 计算库存滞销金/滞销权重
                {
                    if (_库存周转率Mappings.Count > 0)
                    {
                        var _库存周转率采购员List = _库存周转率Mappings.Where(x => !string.IsNullOrEmpty(x._采购人)).Select(x => x._采购人).Distinct().ToList();
                        _库存周转率采购员List.ForEach(buyerName =>
                        {
                            if (Helper.IsBuyer(buyerName))
                            {
                                var ref库存周转率 = _库存周转率Mappings.Where(x => x._采购人 == buyerName).ToList();

                                var cur滞销金滞销权重 = new _滞销金滞销权重Model();
                                cur滞销金滞销权重._采购员 = buyerName;

                                #region 库存周转率>100获取滞销金
                                {
                                    var list = ref库存周转率.Where(x => x._库存周转天数 > 100);
                                    if (list.Count() > 0)
                                    {
                                        cur滞销金滞销权重._滞销_周转天数 = list.Sum(x => x._库存周转天数) / list.Count();
                                        cur滞销金滞销权重._滞销_总金额 = list.Sum(x => x._总金额);
                                        cur滞销金滞销权重._滞销金 = cur滞销金滞销权重._滞销_总金额 / 1000 * 2;
                                    }
                                }
                                #endregion

                                #region 库存周转率<=100获取滞销权重
                                {
                                    var list = ref库存周转率.Where(x => x._库存周转天数 <= 100);
                                    if (list.Count() > 0)
                                    {
                                        cur滞销金滞销权重._权重_周转天数 = list.Sum(x => x._库存周转天数) / list.Count();
                                        cur滞销金滞销权重._权重 = Culc滞销权重(cur滞销金滞销权重._权重_周转天数);
                                        cur滞销金滞销权重._权重_总金额 = list.Sum(x => x._总金额);
                                    }
                                }
                                #endregion

                                _滞销金滞销权重List.Add(cur滞销金滞销权重);
                            }

                        });
                    }
                }
                #endregion

                ShowMsg("开始计算缺货率权重/奖金");

                #region 计算缺货率权重/奖金
                {
                    if (_缺货率Mappings.Count > 0)
                    {
                        var _缺货率采购员List = _缺货率Mappings.Where(x => !string.IsNullOrEmpty(x._采购员)).Select(x => x._采购员).Distinct().ToList();
                        _缺货率采购员List.ForEach(buyerName =>
                        {
                            if (Helper.IsBuyer(buyerName))
                            {
                                var ref缺货率 = _缺货率Mappings.Where(x => x._采购员 == buyerName).ToList();
                                if (ref缺货率.Count > 0)
                                {
                                    var cur缺货率权重 = new _缺货率权重Model();
                                    cur缺货率权重._采购员 = buyerName;

                                    decimal _总交易订单数量 = ref缺货率.Sum(x => x._交易订单数量);
                                    decimal _总缺货订单数量 = ref缺货率.Sum(x => x._缺货订单数量);
                                    decimal _总异常订单数量 = ref缺货率.Sum(x => x._异常订单数量);

                                    if (_总交易订单数量 > 0)
                                    {
                                        var diff = _总缺货订单数量 / _总交易订单数量 * 100;
                                        var _tmp = Culc缺货率权重(diff);
                                        cur缺货率权重._缺货率 = diff;
                                        cur缺货率权重._缺货率权重 = _tmp[0];
                                        cur缺货率权重._缺货率奖励 = _tmp[1];
                                    }

                                    _缺货率权重奖励List.Add(cur缺货率权重);
                                }
                            }

                        });
                    }
                }
                #endregion

                ShowMsg("开始计算议价奖励");

                #region 计算议价奖励
                {
                    if (_议价奖励appings.Count > 0)
                    {
                        var _议价采购员List = _议价奖励appings.Where(x => !string.IsNullOrEmpty(x._采购员)).Select(x => x._采购员).Distinct().ToList();
                        _议价采购员List.ForEach(buyerName =>
                        {
                            if (Helper.IsBuyer(buyerName))
                            {
                                var ref议价 = _议价奖励appings.Where(x => x._采购员 == buyerName).FirstOrDefault();
                                if (ref议价 != null)
                                {
                                    var curModel = new _议价奖励Model();
                                    curModel._采购员 = buyerName;
                                    curModel._议价 = ref议价._议价;
                                    _议价奖励List.Add(curModel);
                                }
                            }
                        });
                    }
                }
                #endregion

                ShowMsg("开始计算未入库处罚");

                #region 计算未入库惩罚
                {
                    if (_未入库Mappings.Count > 0)
                    {
                        var _未入库采购员List = _未入库Mappings.Where(x => !string.IsNullOrEmpty(x._采购员)).Select(x => x._采购员).Distinct().ToList();
                        _未入库采购员List.ForEach(buyerName =>
                        {
                            if (Helper.IsBuyer(buyerName))
                            {
                                var ref未入库 = _未入库Mappings.Where(x => x._采购员 == buyerName).ToList();
                                var curModel = new _未入库Model();
                                curModel._采购员 = buyerName;
                                curModel._未入库金额 = ref未入库.Sum(x => x._未入库金额);
                                _未入库惩罚List.Add(curModel);
                            }
                        });
                    }
                }
                #endregion

                ShowMsg("开始计算最终工资");

                #region 合计最终工资
                {
                    var _所有采购List = Helper.GetBuyers();
                    _所有采购List.ForEach(buyerName =>
                    {
                        //if (buyerName == "鲍祝平")
                        //{

                        //}

                        var curModel = new _工资详情Model();
                        curModel._采购员 = buyerName;

                        #region 计算底薪
                        {
                            if (Helper.IsSpecBuyerType(buyerName, BuyerTypeEnum.ShangHai))
                            {
                                curModel._底薪 = _上海采购底薪;
                                curModel._排序 = (int)BuyerTypeEnum.ShangHai;
                            }
                            else if (Helper.IsSpecBuyerType(buyerName, BuyerTypeEnum.HeFei))
                            {
                                curModel._底薪 = _合肥采购底薪;
                                curModel._排序 = (int)BuyerTypeEnum.HeFei;
                            }
                            else
                            { }
                        }
                        #endregion

                        #region 计算工作量工资(里面已经计算有奖金和滞销金)
                        {
                            decimal _总权重 = 1;
                            var ref工作量奖励 = _订单奖励List.Where(x => x._采购员 == buyerName).FirstOrDefault();
                            if (ref工作量奖励 != null)
                            {
                                var _ref缺货权重 = _缺货率权重奖励List.Where(x => x._采购员 == buyerName).FirstOrDefault();
                                var _ref滞销权重 = _滞销金滞销权重List.Where(x => x._采购员 == buyerName).FirstOrDefault();
                                if (_ref缺货权重 != null)
                                {
                                    _总权重 *= _ref缺货权重._缺货率权重;
                                    curModel._奖金 = _ref缺货权重._缺货率奖励;
                                }
                                if (_ref滞销权重 != null)
                                {
                                    _总权重 *= _ref滞销权重._权重;
                                    curModel._滞销金 = _ref滞销权重._滞销金;
                                }
                                curModel._工作量工资 = ref工作量奖励._合计 * _总权重;

                            }
                        }
                        #endregion

                        #region 计算议价绩效
                        {
                            var ref议价奖励 = _议价奖励List.Where(x => x._采购员 == buyerName).FirstOrDefault();
                            if (ref议价奖励 != null)
                            {
                                curModel._议价绩效 = ref议价奖励._议价奖励;
                            }
                        }
                        #endregion

                        #region 计算未入库惩罚
                        {
                            var ref未入库惩罚 = _未入库惩罚List.Where(x => x._采购员 == buyerName).FirstOrDefault();
                            if (ref未入库惩罚 != null)
                            {
                                curModel._未入库惩罚 = ref未入库惩罚._未入库惩罚;
                            }
                        }
                        #endregion

                        _工资详情List.Add(curModel);
                    });
                }
                #endregion


                Export(_工资详情List.OrderBy(x => x._排序).ToList(), _订单奖励List, _滞销金滞销权重List, _缺货率权重奖励List, _议价奖励List);
            }, null);

        }
        #endregion

        #region 上传库存周转率
        private void btn上传库存周转率_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txt库存周转率.Text = OpenFileDialog1.FileName;

            }
        }
        #endregion

        #region 上传缺货信息
        private void btn上传缺货信息_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txt缺货信息.Text = OpenFileDialog1.FileName;

            }
        }
        #endregion

        #region 上传议价奖励
        private void btn上传议价奖励_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txt议价奖励.Text = OpenFileDialog1.FileName;

            }
        }
        #endregion

        #region 上传未入库
        private void btn上传未入库_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txt未入库.Text = OpenFileDialog1.FileName;

            }
        }
        #endregion

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_采购流水Mapping), typeof(_库存周转率Mapping), typeof(_议价奖励Mapping),
                  typeof(_未入库Mapping), typeof(_缺货率Mapping));

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

        #region Culc缺货率权重 计算缺货率权重
        private decimal[] Culc缺货率权重(decimal que)
        {
            if (que > 0 && que <= 0.15M)
            {
                return new decimal[] { 1.6M, 1000 };
            }
            if (que > 0.15M && que <= 0.3M)
            {
                return new decimal[] { 1.5M, 800 };
            }
            if (que > 0.3M && que <= 0.45M)
            {
                return new decimal[] { 1.4M, 300 };
            }
            if (que > 0.45M && que <= 0.6M)
            {
                return new decimal[] { 1.3M, 0 };
            }
            if (que > 0.6M && que <= 0.75M)
            {
                return new decimal[] { 1.2M, 0 };
            }
            if (que > 0.75M && que <= 0.9M)
            {
                return new decimal[] { 1.1M, 0 };
            }
            if (que > 0.9M && que <= 1.05M)
            {
                return new decimal[] { 1M, 0 };
            }
            if (que > 1.05M && que <= 1.2M)
            {
                return new decimal[] { 0.9M, 0 };
            }
            if (que > 1.2M && que <= 1.35M)
            {
                return new decimal[] { 0.8M, 0 };
            }
            if (que > 1.35M && que <= 1.5M)
            {
                return new decimal[] { 0.7M, 0 };
            }
            if (que > 1.5M && que <= 1.65M)
            {
                return new decimal[] { 0.6M, 0 };
            }
            return new decimal[] { 0.6M, 0 };
        }
        #endregion

        #region Culc滞销权重 计算滞销权重
        private decimal Culc滞销权重(decimal days)
        {
            if (days > 15 && days <= 20)
            {
                return 1;
            }
            if (days > 20 && days <= 25)
            {
                return 0.95M;
            }
            if (days > 25 && days <= 30)
            {
                return 0.9M;
            }
            if (days > 30 && days <= 35)
            {
                return 0.85M;
            }
            if (days > 35 && days <= 40)
            {
                return 0.8M;
            }
            if (days > 40 && days <= 45)
            {
                return 0.75M;
            }
            if (days > 45 && days <= 50)
            {
                return 0.7M;
            }
            if (days > 50 && days <= 60)
            {
                return 0.65M;
            }
            if (days > 60 && days <= 70)
            {
                return 0.6M;
            }
            return 0;
        }
        #endregion

        #region GetInteger 保留小数点两位
        private static decimal GetInteger(decimal dou)
        {
            return Math.Round(dou, 2);
        }
        #endregion

        #region RefreshConfig 刷新配置
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

        #region RefreshAward 刷新奖励金额配置
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

        #region Export 导出结果
        /// <summary>
        /// 导出结果
        /// </summary>
        /// <param name="_工资详情"></param>
        /// <param name="_订单奖励"></param>
        /// <param name="_滞销金权重"></param>
        /// <param name="_缺货率权重"></param>
        /// <param name="_议价奖励"></param>
        private void Export(List<_工资详情Model> _工资详情, List<_订单奖励Model> _订单奖励, List<_滞销金滞销权重Model> _滞销金权重, List<_缺货率权重Model> _缺货率权重, List<_议价奖励Model> _议价奖励)
        {
            ShowMsg("开始生成表格");

            var buffer = new byte[0];

            #region 工作簿
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 最终工资
                {
                    var sheet1 = workbox.Worksheets.Add("最终工资");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "采购员";
                    sheet1.Cells[1, 2].Value = "底薪";
                    sheet1.Cells[1, 3].Value = "制单奖励*总权重";
                    sheet1.Cells[1, 4].Value = "议价绩效";
                    sheet1.Cells[1, 5].Value = "未入库";
                    sheet1.Cells[1, 6].Value = "奖金";
                    sheet1.Cells[1, 7].Value = "滞销金";
                    sheet1.Cells[1, 8].Value = "合计";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = _工资详情.Count; idx < len; idx++)
                    {
                        var curResult = _工资详情[idx];

                        sheet1.Cells[rowIdx, 1].Value = curResult._采购员;
                        sheet1.Cells[rowIdx, 2].Value = curResult._底薪;
                        sheet1.Cells[rowIdx, 3].Value = curResult._工作量工资;
                        sheet1.Cells[rowIdx, 4].Value = curResult._议价绩效;
                        sheet1.Cells[rowIdx, 5].Value = curResult._未入库惩罚;
                        sheet1.Cells[rowIdx, 6].Value = curResult._奖金;
                        sheet1.Cells[rowIdx, 7].Value = curResult._滞销金;
                        sheet1.Cells[rowIdx, 8].Value = curResult._最终工资;

                        rowIdx++;

                    }
                    #endregion
                }
                #endregion

                #region 制单奖励
                {
                    var sheet2 = workbox.Worksheets.Add("制单奖励");

                    #region 标题行
                    sheet2.Cells[2, 1].Value = "采购员";

                    sheet2.Cells[2, 2].Value = "第一阶奖励";
                    sheet2.Cells[2, 3].Value = "第二阶奖励";
                    sheet2.Cells[2, 4].Value = "第三阶奖励";
                    sheet2.Cells[2, 5].Value = "第四阶奖励";
                    sheet2.Cells[2, 6].Value = "第五阶奖励";

                    using (var range = sheet2.Cells[1, 2, 1, 6])
                    {
                        range.Merge = true;
                        range.Value = "第一档";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                    }

                    sheet2.Cells[2, 7].Value = "第一阶奖励";
                    sheet2.Cells[2, 8].Value = "第二阶奖励";
                    sheet2.Cells[2, 9].Value = "第三阶奖励";
                    sheet2.Cells[2, 10].Value = "第四阶奖励";
                    sheet2.Cells[2, 11].Value = "第五阶奖励";
                    using (var range = sheet2.Cells[1, 7, 1, 11])
                    {
                        range.Merge = true;
                        range.Value = "第二档";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                    }


                    sheet2.Cells[2, 12].Value = "第一阶奖励";
                    sheet2.Cells[2, 13].Value = "第二阶奖励";
                    sheet2.Cells[2, 14].Value = "第三阶奖励";
                    sheet2.Cells[2, 15].Value = "第四阶奖励";
                    sheet2.Cells[2, 16].Value = "第五阶奖励";
                    using (var range = sheet2.Cells[1, 12, 1, 16])
                    {
                        range.Merge = true;
                        range.Value = "第三档";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                    }

                    sheet2.Cells[2, 17].Value = "第一阶奖励";
                    sheet2.Cells[2, 18].Value = "第二阶奖励";
                    sheet2.Cells[2, 19].Value = "第三阶奖励";
                    sheet2.Cells[2, 20].Value = "第四阶奖励";
                    sheet2.Cells[2, 21].Value = "第五阶奖励";
                    using (var range = sheet2.Cells[1, 17, 1, 21])
                    {
                        range.Merge = true;
                        range.Value = "第四档";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                    }

                    sheet2.Cells[2, 22].Value = "第一阶奖励";
                    sheet2.Cells[2, 23].Value = "第二阶奖励";
                    sheet2.Cells[2, 24].Value = "第三阶奖励";
                    sheet2.Cells[2, 25].Value = "第四阶奖励";
                    sheet2.Cells[2, 26].Value = "第五阶奖励";
                    using (var range = sheet2.Cells[1, 22, 1, 26])
                    {
                        range.Merge = true;
                        range.Value = "第五档";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                    }

                    sheet2.Cells[2, 27].Value = "制单奖励";
                    sheet2.Cells[2, 28].Value = "合计";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 3, len = _订单奖励.Count; idx < len; idx++)
                    {
                        var curResult = _订单奖励[idx];

                        sheet2.Cells[rowIdx, 1].Value = curResult._采购员;

                        sheet2.Cells[rowIdx, 2].Value = curResult._D1_J1;
                        sheet2.Cells[rowIdx, 3].Value = curResult._D1_J2;
                        sheet2.Cells[rowIdx, 4].Value = curResult._D1_J3;
                        sheet2.Cells[rowIdx, 5].Value = curResult._D1_J4;
                        sheet2.Cells[rowIdx, 6].Value = curResult._D1_J5;

                        sheet2.Cells[rowIdx, 7].Value = curResult._D2_J1;
                        sheet2.Cells[rowIdx, 8].Value = curResult._D2_J2;
                        sheet2.Cells[rowIdx, 9].Value = curResult._D2_J3;
                        sheet2.Cells[rowIdx, 10].Value = curResult._D2_J4;
                        sheet2.Cells[rowIdx, 11].Value = curResult._D2_J5;

                        sheet2.Cells[rowIdx, 12].Value = curResult._D3_J1;
                        sheet2.Cells[rowIdx, 13].Value = curResult._D3_J2;
                        sheet2.Cells[rowIdx, 14].Value = curResult._D3_J3;
                        sheet2.Cells[rowIdx, 15].Value = curResult._D3_J4;
                        sheet2.Cells[rowIdx, 16].Value = curResult._D3_J5;

                        sheet2.Cells[rowIdx, 17].Value = curResult._D4_J1;
                        sheet2.Cells[rowIdx, 18].Value = curResult._D4_J2;
                        sheet2.Cells[rowIdx, 19].Value = curResult._D4_J3;
                        sheet2.Cells[rowIdx, 20].Value = curResult._D4_J4;
                        sheet2.Cells[rowIdx, 21].Value = curResult._D4_J5;

                        sheet2.Cells[rowIdx, 22].Value = curResult._D5_J1;
                        sheet2.Cells[rowIdx, 23].Value = curResult._D5_J2;
                        sheet2.Cells[rowIdx, 24].Value = curResult._D5_J3;
                        sheet2.Cells[rowIdx, 25].Value = curResult._D5_J4;
                        sheet2.Cells[rowIdx, 26].Value = curResult._D5_J5;

                        sheet2.Cells[rowIdx, 27].Value = curResult._产品归属奖励;
                        sheet2.Cells[rowIdx, 28].Value = curResult._合计;

                        rowIdx++;

                    }
                    #endregion
                }
                #endregion

                #region 滞销金滞销权重
                {
                    var sheet3 = workbox.Worksheets.Add("滞销金及权重");

                    #region 标题行
                    using (var range = sheet3.Cells[1, 1, 1, 4])
                    {
                        range.Merge = true;
                        range.Value = "大于100天";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                    }
                    using (var range = sheet3.Cells[1, 5, 1, 7])
                    {
                        range.Merge = true;
                        range.Value = "小于100天";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                    }

                    sheet3.Cells[2, 1].Value = "采购员";
                    sheet3.Cells[2, 2].Value = "总金额";
                    sheet3.Cells[2, 3].Value = "库存周转天数";
                    sheet3.Cells[2, 4].Value = "滞销金";
                    sheet3.Cells[2, 5].Value = "总金额";
                    sheet3.Cells[2, 6].Value = "库存周转天数";
                    sheet3.Cells[2, 7].Value = "滞销权重";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 3, len = _滞销金权重.Count; idx < len; idx++)
                    {
                        var curResult = _滞销金权重[idx];

                        sheet3.Cells[rowIdx, 1].Value = curResult._采购员;
                        sheet3.Cells[rowIdx, 2].Value = curResult._滞销_总金额;
                        sheet3.Cells[rowIdx, 3].Value = curResult._滞销_周转天数;
                        sheet3.Cells[rowIdx, 4].Value = curResult._滞销金;
                        sheet3.Cells[rowIdx, 5].Value = curResult._权重_总金额;
                        sheet3.Cells[rowIdx, 6].Value = curResult._权重_周转天数;
                        sheet3.Cells[rowIdx, 7].Value = curResult._权重;

                        rowIdx++;
                    }
                    #endregion
                }
                #endregion

                #region 缺货率权重
                {
                    var sheet4 = workbox.Worksheets.Add("缺货率权重");

                    #region 标题行
                    sheet4.Cells[1, 1].Value = "采购员";
                    sheet4.Cells[1, 2].Value = "缺货率";
                    sheet4.Cells[1, 3].Value = "缺货率权重";
                    sheet4.Cells[1, 4].Value = "奖金";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = _缺货率权重.Count; idx < len; idx++)
                    {
                        var curResult = _缺货率权重[idx];

                        sheet4.Cells[rowIdx, 1].Value = curResult._采购员;
                        sheet4.Cells[rowIdx, 2].Value = curResult._缺货率;
                        sheet4.Cells[rowIdx, 3].Value = curResult._缺货率权重;
                        sheet4.Cells[rowIdx, 4].Value = curResult._缺货率奖励;

                        rowIdx++;

                    }
                    #endregion
                }
                #endregion

                #region 议价奖励
                {
                    var sheet5 = workbox.Worksheets.Add("议价奖励");

                    #region 标题行
                    sheet5.Cells[1, 1].Value = "采购员";
                    sheet5.Cells[1, 2].Value = "议价";
                    sheet5.Cells[1, 3].Value = "议价奖励";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = _议价奖励.Count; idx < len; idx++)
                    {
                        var curResult = _议价奖励[idx];

                        sheet5.Cells[rowIdx, 1].Value = curResult._采购员;
                        sheet5.Cells[rowIdx, 2].Value = curResult._议价;
                        sheet5.Cells[rowIdx, 3].Value = curResult._议价奖励;
                        rowIdx++;

                    }
                    #endregion
                }
                #endregion

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

        [ExcelTable("采购订单流水表")]
        public class _采购流水Mapping
        {
            private string org采购员;

            [ExcelColumn("采购sku数量")]
            public string _OrgSKU个数 { get; set; }
            [ExcelColumn("总金额")]
            public string _Org总金额 { get; set; }

            [ExcelColumn("采购员")]
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }
            [ExcelColumn("制单人")]
            public string _制单人 { get; set; }
        }

        [ExcelTable("库存周转率报表")]
        public class _库存周转率Mapping
        {
            [ExcelColumn("库存周转天数")]
            public string org库存周转天数 { get; set; }
            [ExcelColumn("采购人")]
            public string org采购人 { get; set; }
            [ExcelColumn("总金额")]
            public string org总金额 { get; set; }

            public string _采购人
            {
                get
                {
                    return org采购人 != null ? org采购人.Trim() : "";
                }
            }
            public decimal _库存周转天数
            {
                get
                {
                    return !string.IsNullOrEmpty(org库存周转天数) ? Convert.ToDecimal(org库存周转天数) : 0;
                }
            }

            public decimal _总金额
            {
                get
                {
                    return !string.IsNullOrEmpty(org总金额) ? Convert.ToDecimal(org总金额) : 0;
                }
            }
        }

        [ExcelTable("议价绩效表")]
        public class _议价奖励Mapping
        {
            private string org采购员;

            [ExcelColumn("议价")]
            public string org议价 { get; set; }

            [ExcelColumn("采购员")]
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }

            public decimal _议价
            {
                get
                {
                    return !string.IsNullOrEmpty(org议价) ? Convert.ToDecimal(org议价) : 0;
                }
            }

        }

        [ExcelTable("审核未入库表")]
        public class _未入库Mapping
        {
            private string org采购员;
            [ExcelColumn("采购员")]
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }
            [ExcelColumn("未入库金额")]
            public string org未入库金额 { get; set; }

            public decimal _未入库金额
            {
                get
                {
                    return !string.IsNullOrEmpty(org未入库金额) ? Convert.ToDecimal(org未入库金额) : 0;
                }
            }

        }

        [ExcelTable("缺货率报表")]
        public class _缺货率Mapping
        {
            private string org采购员;
            [ExcelColumn("采购员")]
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("交易订单数量")]
            public string org交易订单数量 { get; set; }
            [ExcelColumn("异常订单数量")]
            public string org异常订单数量 { get; set; }
            [ExcelColumn("缺货订单数量")]
            public string org缺货订单数量 { get; set; }

            public decimal _交易订单数量
            {
                get
                {
                    return !string.IsNullOrEmpty(org交易订单数量) ? Convert.ToDecimal(org交易订单数量) : 0;
                }
            }

            public decimal _异常订单数量
            {
                get
                {
                    return !string.IsNullOrEmpty(org异常订单数量) ? Convert.ToDecimal(org异常订单数量) : 0;
                }
            }

            public decimal _缺货订单数量
            {
                get
                {
                    return !string.IsNullOrEmpty(org缺货订单数量) ? Convert.ToDecimal(org缺货订单数量) : 0;
                }
            }
        }

        public class _采购流水Model
        {
            private string org采购员;
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }
            public string _制单人 { get; set; }
            public int _SKU个数 { get; set; }
            public decimal _总金额 { get; set; }
        }

        public class _订单奖励Model
        {
            private string org采购员;
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }

            public decimal _D1_J1 { get; set; }
            public decimal _D1_J2 { get; set; }
            public decimal _D1_J3 { get; set; }
            public decimal _D1_J4 { get; set; }
            public decimal _D1_J5 { get; set; }
            public decimal _D1_Sum
            {
                get
                {
                    return _D1_J1 + _D1_J2 + _D1_J3 + _D1_J4 + _D1_J5;
                }
            }


            public decimal _D2_J1 { get; set; }
            public decimal _D2_J2 { get; set; }
            public decimal _D2_J3 { get; set; }
            public decimal _D2_J4 { get; set; }
            public decimal _D2_J5 { get; set; }
            public decimal _D2_Sum
            {
                get
                {
                    return _D2_J1 + _D2_J2 + _D2_J3 + _D2_J4 + _D2_J5;
                }
            }

            public decimal _D3_J1 { get; set; }
            public decimal _D3_J2 { get; set; }
            public decimal _D3_J3 { get; set; }
            public decimal _D3_J4 { get; set; }
            public decimal _D3_J5 { get; set; }
            public decimal _D3_Sum
            {
                get
                {
                    return _D3_J1 + _D3_J2 + _D3_J3 + _D3_J4 + _D3_J5;
                }
            }

            public decimal _D4_J1 { get; set; }
            public decimal _D4_J2 { get; set; }
            public decimal _D4_J3 { get; set; }
            public decimal _D4_J4 { get; set; }
            public decimal _D4_J5 { get; set; }
            public decimal _D4_Sum
            {
                get
                {
                    return _D4_J1 + _D4_J2 + _D4_J3 + _D4_J4 + _D4_J5;
                }
            }

            public decimal _D5_J1 { get; set; }
            public decimal _D5_J2 { get; set; }
            public decimal _D5_J3 { get; set; }
            public decimal _D5_J4 { get; set; }
            public decimal _D5_J5 { get; set; }
            public decimal _D5_Sum
            {
                get
                {
                    return _D5_J1 + _D5_J2 + _D5_J3 + _D5_J4 + _D5_J5;
                }
            }

            public decimal _产品归属奖励 { get; set; }

            public decimal _合计
            {
                get
                {
                    return _D1_Sum + _D2_Sum + _D3_Sum + _D4_Sum + _D5_Sum + _产品归属奖励;
                }
            }

            public bool IsBuyer
            {
                get
                {
                    return Helper.IsBuyer(this._采购员);
                }
            }
        }

        public class _滞销金滞销权重Model
        {
            private string org采购员;
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }
            public decimal _滞销_总金额 { get; set; }
            public decimal _滞销_周转天数 { get; set; }
            public decimal _滞销金 { get; set; }
            public decimal _权重_总金额 { get; set; }
            public decimal _权重_周转天数 { get; set; }
            public decimal _权重 { get; set; }
        }

        public class _缺货率权重Model
        {
            private string org采购员;
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }
            public decimal _缺货率 { get; set; }
            public decimal _缺货率权重 { get; set; }
            public decimal _缺货率奖励 { get; set; }
        }

        public class _议价奖励Model
        {
            private string org采购员;
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }
            public decimal _议价 { get; set; }
            public decimal _议价奖励
            {
                get
                {
                    return _议价 * 0.3M;
                }
            }
        }

        public class _未入库Model
        {
            private string org采购员;
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }
            public decimal _未入库金额 { get; set; }
            public decimal _未入库惩罚
            {
                get
                {
                    return _未入库金额 * 0.3M;
                }
            }
        }

        public class _工资详情Model
        {
            private string org采购员;
            public string _采购员
            {
                get
                {
                    return org采购员;
                }
                set
                {
                    org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }
            public decimal _底薪 { get; set; }
            public decimal _工作量工资 { get; set; }
            public decimal _议价绩效 { get; set; }
            public decimal _未入库惩罚 { get; set; }
            public decimal _奖金 { get; set; }
            public decimal _滞销金 { get; set; }
            public decimal _最终工资
            {
                get
                {
                    return _底薪 + _工作量工资 + _议价绩效 + _奖金 - _未入库惩罚 - _滞销金;
                }
            }
            public int _排序 { get; set; }
        }


    }
}
