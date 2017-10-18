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

            txtUpDfkunsYj.Text = @"C:\Users\Leon\Desktop\mm\上海-默认发货仓库昆山.xls";//默认昆山预警
            txtUpKsYj.Text = @"C:\Users\Leon\Desktop\mm\昆山建议采购.xls";//昆山库存预警
            txtUpKsKc.Text = @"C:\Users\Leon\Desktop\mm\昆山所有库存.xls";//昆山所有库存
            txtUpSHKc.Text = @"C:\Users\Leon\Desktop\mm\上海所有库存.xls";//上海所有库存
            txtUpTmp.Text = @"C:\Users\Leon\Desktop\mm\备货.xls";//临时备货


            //txtUpDfkunsYj.Text = @"C:\Users\Leon\Desktop\排除重复项\上海-默认昆山仓.xls";//默认昆山预警
            //txtUpKsYj.Text = @"C:\Users\Leon\Desktop\排除重复项\昆山建议采购.xls";//昆山库存预警
            //txtUpKsKc.Text = @"C:\Users\Leon\Desktop\排除重复项\昆山所有库存.xls";//昆山所有库存
            //txtUpSHKc.Text = @"C:\Users\Leon\Desktop\排除重复项\上海所有库存.xls";//上海所有库存
            //txtUpTmp.Text = @"C:\Users\Leon\Desktop\排除重复项\备货.xls";//临时备货

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
                var _d订单金额 = Convert.ToDouble(NtxtAmount.Value);

                var _Im上海默认昆山预警 = new List<Warning>();
                var _Im昆山预警 = new List<Warning>();
                var _Im昆山库存 = new List<Warning>();
                var _Im上海库存 = new List<Warning>();
                var _Im临时备货 = new List<Warning>();
                var _最终需要采购的预警 = new List<Warning>();
                var _Ex库存充足的预警 = new List<Warning>();
                var _Ex采购需要采购订单 = new List<Order>();
                var _Ex开发需要采购订单 = new List<Order>();


                var str上海默认昆山ExcelPath = txtUpDfkunsYj.Text;
                var str昆山预警ExcelPath = txtUpKsYj.Text;
                var str昆山库存ExcelPath = txtUpKsKc.Text;
                var str上海库存ExcelPath = txtUpSHKc.Text;
                var str临时备货ExcelPath = txtUpTmp.Text;


                #region 读取源数据
                var actRead = new Action(() =>
                {
                    ShowMsg("开始读取表格数据");

                    #region 读取上海默认昆山预警
                    if (!string.IsNullOrEmpty(str上海默认昆山ExcelPath))
                    {
                        using (var excel = new ExcelQueryFactory(str上海默认昆山ExcelPath))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<Warning>(s)
                                              select c;
                                    _Im上海默认昆山预警.AddRange(tmp);
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                    #endregion

                    #region 读取昆山预警
                    if (!string.IsNullOrEmpty(str昆山预警ExcelPath))
                    {
                        using (var excel = new ExcelQueryFactory(str昆山预警ExcelPath))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<Warning>(s)
                                              select c;
                                    _Im昆山预警.AddRange(tmp);
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                    #endregion

                    #region 读取昆山仓库
                    if (!string.IsNullOrEmpty(str昆山库存ExcelPath))
                    {
                        using (var excel = new ExcelQueryFactory(str昆山库存ExcelPath))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<Warning>(s)
                                              select c;
                                    _Im昆山库存.AddRange(tmp);
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                    #endregion

                    #region 读取上海仓库
                    if (!string.IsNullOrEmpty(str上海库存ExcelPath))
                    {
                        using (var excel = new ExcelQueryFactory(str上海库存ExcelPath))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<Warning>(s)
                                              select c;
                                    _Im上海库存.AddRange(tmp);
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                    #endregion

                    #region 读取临时备货
                    if (!string.IsNullOrEmpty(str临时备货ExcelPath))
                    {
                        using (var excel = new ExcelQueryFactory(str临时备货ExcelPath))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<Warning>(s)
                                              select c;
                                    _Im临时备货.AddRange(tmp);
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


                #region 数据处理
                actRead.BeginInvoke((obj) =>
                {
                    ShowMsg("开始计算表格数据");

                    var _List两个预警表共有Sku唯一 = new List<string>();//两个表共有的sku,不用判断是否需要采购,直接将需要采购量相加
                    #region 计算两个预警表共有的sku唯一
                    {
                        var _List上海默认昆山预警Sku唯一 = _Im上海默认昆山预警.Where(m => !string.IsNullOrEmpty(m._SKU)).Select(m => m._SKU).Distinct().ToList();
                        var _List昆山预警Sku唯一 = _Im昆山预警.Where(m => !string.IsNullOrEmpty(m._SKU)).Select(m => m._SKU).Distinct().ToList();
                        //提取出两表共有的sku
                        if (_List上海默认昆山预警Sku唯一.Count > _List昆山预警Sku唯一.Count)
                        {
                            _List上海默认昆山预警Sku唯一.ForEach(m =>
                            {
                                if (!string.IsNullOrEmpty(m))
                                {
                                    int isCommonSku = _List昆山预警Sku唯一.Count(ss => ss == m);
                                    if (isCommonSku > 0)
                                        _List两个预警表共有Sku唯一.Add(m);
                                }
                            });
                        }
                        else
                        {
                            _List昆山预警Sku唯一.ForEach(m =>
                            {
                                if (!string.IsNullOrEmpty(m))
                                {
                                    int isCommonSku = _List上海默认昆山预警Sku唯一.Count(ss => ss == m);
                                    if (isCommonSku > 0)
                                        _List两个预警表共有Sku唯一.Add(m);
                                }
                            });
                        }
                    }
                    #endregion


                    //遍历 上海默认昆山预警
                    //1.把共有的sku的 采购建议(库存上限+库存下限...)相加起来
                    //2.把独有的sku 进入昆山所有库存判断是否需要采购
                    #region 遍历 上海默认昆山预警
                    {
                        _Im上海默认昆山预警.ForEach(cur预警Item =>
                        {
                            if (!string.IsNullOrEmpty(cur预警Item._SKU))
                            {
                                //共有的sku 采购建议(库存上限+库存下限...)相加起来
                                if (_List两个预警表共有Sku唯一.Count(ss => ss == cur预警Item._SKU) > 0)
                                {
                                    var ref昆山SkuItem = _Im昆山预警.Where(kk => kk._SKU == cur预警Item._SKU).FirstOrDefault();
                                    if (ref昆山SkuItem != null)
                                    {
                                        var needOrderItem = new Warning();
                                        needOrderItem._SKU = cur预警Item._SKU;
                                        needOrderItem._供应商 = cur预警Item._供应商;
                                        needOrderItem._采购员 = cur预警Item._采购员;
                                        needOrderItem._商品成本单价 = cur预警Item._商品成本单价;
                                        needOrderItem._仓库 = cur预警Item._仓库;
                                        needOrderItem._采购未入库 = ref昆山SkuItem._采购未入库;
                                        //要相加的部分
                                        needOrderItem._采购未入库 = ref昆山SkuItem._采购未入库 + cur预警Item._采购未入库;
                                        needOrderItem._可用数量 = ref昆山SkuItem._可用数量 + cur预警Item._可用数量;
                                        needOrderItem._库存上限 = ref昆山SkuItem._库存上限 + cur预警Item._库存上限;
                                        needOrderItem._库存下限 = ref昆山SkuItem._库存下限 + cur预警Item._库存下限;
                                        needOrderItem._缺货及未派单数量 = ref昆山SkuItem._缺货及未派单数量 + cur预警Item._缺货及未派单数量;
                                        _最终需要采购的预警.Add(needOrderItem);
                                    }
                                }
                                //独有的sku 进入昆山所有库存判断是否需要采购
                                else
                                {
                                    //昆山所有库存没有该sku,是需要采购的,不用判断,直接加入 _最终需要采购的预警
                                    var ref昆山库存SkuItem = _Im昆山库存.Where(cc => cc._SKU == cur预警Item._SKU).FirstOrDefault();
                                    if (ref昆山库存SkuItem != null)
                                    {
                                        if (ref昆山库存SkuItem._建议采购数量 + cur预警Item._建议采购数量 > 0)
                                        {
                                            _最终需要采购的预警.Add(cur预警Item);
                                        }
                                        else
                                        {
                                            _Ex库存充足的预警.Add(cur预警Item);
                                        }
                                    }
                                    else
                                    {
                                        _最终需要采购的预警.Add(cur预警Item);
                                    }
                                }
                            }

                        });
                    }
                    #endregion


                    //遍历 昆山预警
                    //1.把独有的sku 进入上海所有库存判断是否需要采购
                    #region 遍历 昆山预警
                    {
                        _Im昆山预警.ForEach(cur预警Item =>
                        {
                            if (!string.IsNullOrEmpty(cur预警Item._SKU))
                            {
                                //共有的sku已经处理,这里只对独有的sku判断
                                if (_List两个预警表共有Sku唯一.Count(ss => ss == cur预警Item._SKU) == 0)
                                {
                                    var ref上海库存SkuItem = _Im上海库存.Where(cc => cc._SKU == cur预警Item._SKU).FirstOrDefault();
                                    if (ref上海库存SkuItem != null)
                                    {
                                        if (ref上海库存SkuItem._建议采购数量 + cur预警Item._建议采购数量 > 0)
                                        {
                                            _最终需要采购的预警.Add(cur预警Item);
                                        }
                                        else
                                        {
                                            _Ex库存充足的预警.Add(cur预警Item);
                                        }
                                    }
                                    else
                                    {
                                        //如果这个独有的sku没有出现在上海库存,不用判断,直接加入 最终需要采购的预警
                                        _最终需要采购的预警.Add(cur预警Item);
                                    }
                                }
                            }
                        });
                    }
                    #endregion


                    //加入临时备货
                    #region 加入临时备货
                    {
                        //临时备货里面可能会有和 _最终需要采购的预警里面相同的sku,这时候需要合并,否则直接添加
                        var final需要采购sku唯一 = _最终需要采购的预警.Where(m=>!string.IsNullOrEmpty(m._SKU)).Select(ss => ss._SKU).Distinct().ToList();
                        _Im临时备货.ForEach(cur预警Item =>
                        {
                            //共有sku
                            var ref最终需要采购的预警Item = _最终需要采购的预警.Where(ss => ss._SKU == cur预警Item._SKU).FirstOrDefault();
                            if (ref最终需要采购的预警Item != null)
                            {
                                ref最终需要采购的预警Item._缺货及未派单数量 += cur预警Item._缺货及未派单数量;
                                _最终需要采购的预警.Add(ref最终需要采购的预警Item);
                            }
                            else
                            {
                                _最终需要采购的预警.Add(cur预警Item);
                            }
                        });
                    }
                    #endregion

                    var _List供应商唯一 = _最终需要采购的预警.Select(p => p._供应商).Distinct().ToList();

                    //计算采购金额,转换采购
                    #region 计算采购金额,转换采购
                    {
                        _List供应商唯一.ForEach(strCur供应商 =>
                        {
                            var ref供应商预警Items = _最终需要采购的预警.Where(ss => ss._供应商 == strCur供应商).ToList();
                            var ref供应商预警采购金额总计 = ref供应商预警Items.Select(ss => ss._采购金额).Sum();
                            //小于分界,分给合肥
                            if (ref供应商预警采购金额总计 <= _d订单金额)
                            {
                                ref供应商预警Items.ForEach(cur预警Item =>
                                {
                                    var curOrder = new Order();
                                    curOrder._供应商 = strCur供应商;
                                    curOrder._SKU = cur预警Item._SKU;
                                    curOrder._Qty = cur预警Item._建议采购数量;
                                    curOrder._采购员 = Helper.ChangeLowerBuyer(cur预警Item._采购员);
                                    curOrder._含税单价 = cur预警Item._商品成本单价;
                                    curOrder._制单人 = cur预警Item._采购员;
                                    curOrder._对应供应商采购金额 = ref供应商预警采购金额总计;
                                    _Ex采购需要采购订单.Add(curOrder);
                                });
                            }
                            else
                            {
                                ref供应商预警Items.ForEach(cur预警Item =>
                                {
                                    var curOrder = new Order();
                                    curOrder._供应商 = strCur供应商;
                                    curOrder._SKU = cur预警Item._SKU;
                                    curOrder._Qty = cur预警Item._建议采购数量;
                                    curOrder._采购员 = cur预警Item._采购员;
                                    curOrder._含税单价 = cur预警Item._商品成本单价;
                                    curOrder._制单人 = cur预警Item._采购员;
                                    curOrder._对应供应商采购金额 = ref供应商预警采购金额总计;
                                    _Ex采购需要采购订单.Add(curOrder);
                                });
                            }



                        });
                    }
                    #endregion


                    //计算完毕,开始导出数据
                    ExportExcel(_Ex采购需要采购订单, _Ex库存充足的预警);

                }, null);
                #endregion
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
        private void ExportExcel(List<Order> orders, List<Warning> notBuyWarnings)
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
                    sheet1.Cells[rowIdx, 4].Value = -curOrder._建议采购数量;

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
