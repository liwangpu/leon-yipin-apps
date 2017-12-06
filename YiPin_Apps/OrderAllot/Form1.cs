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
using CommonLibs;

namespace OrderAllot
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();

            //txtUpShangsYj.Text = @"C:\Users\pulw\Desktop\订单分配\上海建议采购.xlsx";
            //txtUpKunsStore.Text = @"C:\Users\pulw\Desktop\订单分配\昆山所有库存.xlsx";

        }

        #region Form1_Load
        private void Form1_Load(object sender, EventArgs e)
        {
            NtxtAmount.Value = 100;
        }
        #endregion

        #region 上传上海库存预警
        private void btnUpShangsYj_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpShangsYj.Text = OpenFileDialog1.FileName;
            }
        }
        #endregion

        #region 上传昆山所有库存
        private void btnUpKunsStore_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpKunsStore.Text = OpenFileDialog1.FileName;
            }
        }
        #endregion

        #region 上传热销
        private void btnUpHot_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpHot.Text = OpenFileDialog1.FileName;
            }
        }
        #endregion

        #region 处理数据
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            try
            {
                var bIn5Day = cb5day.Checked;
                var bIn15Day = cb15day.Checked;
                var bIn30Day = cb30day.Checked;

                #region 解析并计算
                var d订单分配金额 = Convert.ToDouble(NtxtAmount.Value);
                var _Im上海库存预警 = new List<_除热销_Warning>();
                var _Im昆山所有库存 = new List<_除热销_Warning>();
                var _List需要采购的预警 = new List<_除热销_Warning>();
                var _Im热销产品 = new List<_热销产品>();
                var _Ex采购订单分配 = new List<Order>();
                var _Ex开发订单分配 = new List<Order>();//把开发单独分写成一个表格 
                var str上海库存预警ExcelPath = txtUpShangsYj.Text;
                var str昆山所有库存ExcelPath = txtUpKunsStore.Text;
                var str热销产品ExcelPath = txtUpHot.Text;
                #region 读取数据
                var actRead = new Action(() =>
                {
                    ShowMsg("开始读取表格数据");

                    #region 读取上海库存预警
                    if (!string.IsNullOrEmpty(str上海库存预警ExcelPath))
                    {
                        using (var excel = new ExcelQueryFactory(str上海库存预警ExcelPath))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<_除热销_Warning>(s)
                                              select c;
                                    _Im上海库存预警.AddRange(tmp);
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                    #endregion

                    #region 读取昆山所有库存
                    if (!string.IsNullOrEmpty(str昆山所有库存ExcelPath))
                    {
                        using (var excel = new ExcelQueryFactory(str昆山所有库存ExcelPath))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<_除热销_Warning>(s)
                                              select c;
                                    _Im昆山所有库存.AddRange(tmp);
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                    #endregion

                    #region 读取热销产品
                    if (!string.IsNullOrEmpty(str热销产品ExcelPath))
                    {
                        using (var excel = new ExcelQueryFactory(str热销产品ExcelPath))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<_热销产品>(s)
                                              select c;
                                    _Im热销产品.AddRange(tmp);
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                    #endregion
                });
                #endregion

                #region 数据分析
                actRead.BeginInvoke((obj) =>
                {
                    ShowMsg("开始计算表格数据");

                    //判断是否需要采购,如需要加入 _List需要采购的预警
                    #region 判断是否需要采购
                    {
                        _Im上海库存预警.ForEach(cur库存预警Item =>
                        {
                            //if (cur库存预警Item._SKU == "DNFK5K72-BL")
                            //{

                            //}


                            if (!string.IsNullOrEmpty(cur库存预警Item._SKU))
                            {
                                if (cur库存预警Item._建议采购数量 > 0)
                                {
                                    var ref昆山库存Item = _Im昆山所有库存.Where(ss => ss._SKU == cur库存预警Item._SKU).FirstOrDefault();
                                    if (ref昆山库存Item != null)
                                    {
                                        if (ref昆山库存Item._建议采购数量 + cur库存预警Item._建议采购数量 > 0)
                                        {
                                            //需要把两个相加起来
                                            var needOrderItem = new _除热销_Warning();
                                            needOrderItem._SKU = cur库存预警Item._SKU;
                                            needOrderItem._供应商 = cur库存预警Item._供应商;
                                            needOrderItem._采购员 = cur库存预警Item._采购员;
                                            needOrderItem._商品成本单价 = cur库存预警Item._商品成本单价;
                                            needOrderItem._仓库 = cur库存预警Item._仓库;
                                            //要相加的部分
                                            needOrderItem.org采购未入库 = (ref昆山库存Item._采购未入库 + cur库存预警Item._采购未入库).ToString();
                                            needOrderItem._可用数量 = ref昆山库存Item._可用数量 + cur库存预警Item._可用数量;
                                            needOrderItem._库存上限 = ref昆山库存Item._库存上限 + cur库存预警Item._库存上限;
                                            needOrderItem._库存下限 = ref昆山库存Item._库存下限 + cur库存预警Item._库存下限;
                                            needOrderItem.org缺货及未派单数量 = (ref昆山库存Item._缺货及未派单数量 + cur库存预警Item._缺货及未派单数量).ToString();

                                            needOrderItem._30天销量 = ref昆山库存Item._30天销量 + cur库存预警Item._30天销量;
                                            needOrderItem._15天销量 = ref昆山库存Item._15天销量 + cur库存预警Item._15天销量;
                                            needOrderItem._5天销量 = ref昆山库存Item._5天销量 + cur库存预警Item._5天销量;

                                            _List需要采购的预警.Add(needOrderItem);
                                        }
                                    }
                                    else
                                    {
                                        //昆山没有该记录,直接采购
                                        _List需要采购的预警.Add(cur库存预警Item);
                                    }
                                }
                            }




                        });
                    }
                    #endregion

                    //重新计算因为热销产生的建议采购数量过大
                    #region 重新计算因为热销产生的建议采购数量过大
                    if (_Im热销产品.Count > 0)
                    {

                        for (int idx = _List需要采购的预警.Count - 1; idx >= 0; idx--)
                        {
                            var sh = _List需要采购的预警[idx];
                            //if (sh._SKU == "MVPA18B65-FU")
                            //{

                            //}

                            var normal = sh._最终需要采购数量;

                            var refHot = _Im热销产品.Where(x => x._SKU == sh._SKU).FirstOrDefault();
                            if (refHot != null)
                            {
                                //除了热销这两天
                                //var _50天销量总和 = sh._30天销量 + sh._15天销量 + sh._5天销量;
                                //var _排除热销天数销量总和 = (sh._30天销量 - refHot._销量)/30 + sh._15天销量) / 30 + (sh._5天销量 - refHot._销量 * 3);

                                var _30av = (sh._30天销量 - (bIn30Day ? refHot._销量 : 0)) / 30;
                                var _15av = (sh._15天销量 - (bIn15Day ? refHot._销量 : 0)) / 15;
                                var _5av = (sh._5天销量 - (bIn5Day ? refHot._销量 : 0)) / 5;

                                sh._日销量 = (_30av + _15av + _5av) / 3;
                                sh.IsHot = true;
                                if (sh._最终需要采购数量 > normal)
                                {
                                    sh.IsHot = false;
                                }
                            }

                            if (sh._最终需要采购数量 <= 0)
                            {
                                _List需要采购的预警.RemoveAt(idx);
                            }
                        }
                    }
                    #endregion

                    //把最终建议采购的sku 两个仓库的 可用库存+可用库存 是否大于 两个仓库的 库存下限+库存下限
                    //如果大于 那么说明这个sku不需要采购
                    #region 二次判断排除生成另一张表

                    for (int idx = _List需要采购的预警.Count - 1; idx >= 0; idx--)
                    {
                        var curItem = _List需要采购的预警[idx];
                        double _两个仓库可用库存以及采购未入库之和 = 0;//可用库存+可用库存+采购未入库+采购未入库-缺货及未派单
                        double _两个库存下限之和 = 0;//库存下限+库存下限
                        double _缺货以及未派单 = 0;
                        var _昆山仓Item = _Im昆山所有库存.Where(x => x._SKU == curItem._SKU).FirstOrDefault();

                        if (_昆山仓Item != null)
                        {
                            _两个仓库可用库存以及采购未入库之和 += _昆山仓Item._可用数量 + _昆山仓Item._采购未入库;
                            _两个库存下限之和 += _昆山仓Item._库存下限;
                            _缺货以及未派单 += _昆山仓Item._缺货及未派单数量;
                        }

                        if (_两个仓库可用库存以及采购未入库之和 - _缺货以及未派单 >= _两个库存下限之和)
                        {
                            _List需要采购的预警.RemoveAt(idx);
                        }
                    }
                    #endregion

                    ////供应商唯一取值
                    var strProviderNames = _Im上海库存预警.Select(p => p._供应商).Distinct().OrderBy(p => p).ToList();

                    #region 计算划分采购订单
                    strProviderNames.ForEach(strCurProviderName =>
                    {
                        if (!string.IsNullOrEmpty(strCurProviderName))
                        {
                            var refCur供应商预警Items = _List需要采购的预警.Where(ss => ss._供应商 == strCurProviderName).ToList();
                            var refCur供应商采购金额总计 = refCur供应商预警Items.Select(ss => ss._采购金额).Sum();
                            //小于分界,分给合肥
                            if (refCur供应商采购金额总计 <= d订单分配金额)
                            {
                                refCur供应商预警Items.ForEach(cur库存预警Item =>
                                {
                                    var curOrder = new Order();
                                    curOrder._供应商 = strCurProviderName;
                                    curOrder._SKU = cur库存预警Item._SKU;
                                    curOrder._Qty = cur库存预警Item._最终需要采购数量;
                                    curOrder._采购员 = Helper.ChangeLowerBuyer(cur库存预警Item._采购员);
                                    curOrder._含税单价 = cur库存预警Item._商品成本单价;
                                    curOrder._制单人 = cur库存预警Item._采购员;
                                    curOrder._对应供应商采购金额 = refCur供应商采购金额总计;
                                    if (Helper.IsBuyer(cur库存预警Item._采购员))
                                        _Ex采购订单分配.Add(curOrder);
                                    else
                                        _Ex开发订单分配.Add(curOrder);
                                });
                            }
                            else
                            {
                                refCur供应商预警Items.ForEach(cur库存预警Item =>
                                {
                                    var curOrder = new Order();
                                    curOrder._供应商 = strCurProviderName;
                                    curOrder._SKU = cur库存预警Item._SKU;
                                    curOrder._Qty = cur库存预警Item._最终需要采购数量;
                                    curOrder._采购员 = cur库存预警Item._采购员;
                                    curOrder._含税单价 = cur库存预警Item._商品成本单价;
                                    curOrder._制单人 = cur库存预警Item._采购员;
                                    curOrder._对应供应商采购金额 = refCur供应商采购金额总计;
                                    if (Helper.IsBuyer(cur库存预警Item._采购员))
                                        _Ex采购订单分配.Add(curOrder);
                                    else
                                        _Ex开发订单分配.Add(curOrder);
                                });
                            }
                        }
                        else
                        {
                            //空白供应商,可能有不同的采购员,不需要转换
                            var refCur供应商预警Items = _List需要采购的预警.Where(ss => string.IsNullOrEmpty(ss._供应商)).ToList();
                            refCur供应商预警Items.ForEach(cur库存预警Item =>
                            {
                                var curOrder = new Order();
                                curOrder._供应商 = strCurProviderName;
                                curOrder._SKU = cur库存预警Item._SKU;
                                curOrder._Qty = cur库存预警Item._最终需要采购数量;
                                curOrder._采购员 = cur库存预警Item._采购员;
                                curOrder._含税单价 = cur库存预警Item._商品成本单价;
                                curOrder._制单人 = cur库存预警Item._采购员;
                                curOrder._对应供应商采购金额 = 0;
                                if (Helper.IsBuyer(cur库存预警Item._采购员))
                                    _Ex采购订单分配.Add(curOrder);
                                else
                                    _Ex开发订单分配.Add(curOrder);
                            });
                        }
                    });
                    #endregion


                    //计算完毕,开始导出数据
                    ExportExcel(_Ex采购订单分配, _Ex开发订单分配);

                }, null);
                #endregion

                #endregion
            }
            catch (Exception ex)
            {
                ShowMsg(ex.Message);
            }
        }
        #endregion

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_除热销_Warning), typeof(_热销产品));

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


        #region ExportExcel 导出Excel表格
        /// <summary>
        /// 导出Excel表格
        /// </summary>
        /// <param name="List采购订单"></param>
        private void ExportExcel(List<Order> List采购订单, List<Order> List开发订单)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer2 = new byte[0];
            var buffer3 = new byte[0];

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
                for (int idx = 0, rowIdx = 2, len = List采购订单.Count; idx < len; idx++, rowIdx++)
                {
                    var curOrder = List采购订单[idx];
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
            #endregion

            #region 工作量单独表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                var List采购开发订单 = new List<Order>();
                List采购开发订单.AddRange(List采购订单);
                List采购开发订单.AddRange(List开发订单);

                #region 标题行
                sheet1.Cells[1, 1].Value = "采购员";
                sheet1.Cells[1, 2].Value = "订单量";
                #endregion

                #region 数据行
                var buyers = new List<string>();
                buyers = List采购开发订单.Where(x => !string.IsNullOrEmpty(x._采购员)).Select(x => x._采购员).Distinct().ToList();
                for (int idx = 0, len = buyers.Count, rowIdx = 2; idx < len; idx++, rowIdx++)
                {
                    var curBuyerName = buyers[idx];
                    var refOrders = List采购开发订单.Where(m => m._采购员 == curBuyerName).ToList();
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
                for (int idx = 0, rowIdx = 2, len = List开发订单.Count; idx < len; idx++, rowIdx++)
                {
                    var curOrder = List开发订单[idx];
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
