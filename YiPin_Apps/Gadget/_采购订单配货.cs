using CommonLibs;
using Gadget.Libs;
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
    public partial class _采购订单配货 : Form
    {
        public _采购订单配货()
        {
            InitializeComponent();
        }
        private void _采购订单配货_Load(object sender, EventArgs e)
        {
            txt建议采购.Text = @"C:\Users\Leon\Desktop\a1\昆山建议采购.csv";
            txt下半月流水.Text = @"C:\Users\Leon\Desktop\a1\下半月流水.csv";
        }

        /**************** button event ****************/

        #region 上传建议采购事件
        private void btn建议采购_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt建议采购);
        }
        #endregion

        #region 上传下半月流水事件
        private void btn下半月流水_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt下半月流水);
        }
        #endregion

        #region 处理按钮事件
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            var d销量差 = nup上下半月销量差.Value;
            var list建议采购 = new List<_建议采购>();
            var list下半月流水 = new List<_下半月流水>();
            var list下半月流水详细 = new List<KeyValuePair<string, decimal>>();


            #region 读取数据
            var actReadData = new Action(() =>
            {
                var strError = string.Empty;

                ShowMsg("开始读取建议采购数据");
                FormHelper.ReadCSVFile<_建议采购>(txt建议采购.Text, ref list建议采购, ref strError);
                ShowMsg(strError);

                ShowMsg("开始读下半月销售流水数据");
                FormHelper.ReadCSVFile<_下半月流水>(txt下半月流水.Text, ref list下半月流水, ref strError);
                ShowMsg(strError);

                #region 解析所有销售流水
                {
                    list下半月流水.ForEach(x =>
                    {
                        list下半月流水详细.AddRange(x._订单详情);
                    });
                }
                #endregion
            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                var list销量差降详情 = new List<_销量差降详情>();
                var list异常单详情 = new List<_异常单详情>();
                var list订单分配 = new List<_订单分配>();
                var list处理对比表 = new List<_处理对比表>();

                list建议采购.ForEach(_建议采购Item =>
                {
                    var model = new _订单分配();
                    model._SKU = _建议采购Item.SKU;
                    model._供应商 = _建议采购Item._供应商;
                    model._采购员 = _建议采购Item._采购员;
                    model._制单人 = _建议采购Item._采购员;
                    model._含税单价 = _建议采购Item._商品成本单价;


                    //输入销量差,获取建议采购数量
                    _建议采购Item._销量差 = d销量差;


                    var _d异常单销量和 = 0m;
                    if (_建议采购Item._上下半月销量差 >= d销量差)
                    {
                        var DifModel = new _销量差降详情();
                        DifModel.SKU = _建议采购Item.SKU;
                        DifModel._上半月销量 = _建议采购Item._上半月销量;
                        DifModel._下半月销量 = _建议采购Item._下半月销量;
                        list销量差降详情.Add(DifModel);

                        //库存上升,需要用拉依达法则去掉异常值
                        if (_建议采购Item._下半月销量 > _建议采购Item._上半月销量)
                        {
                            var refDetails = list下半月流水详细.Where(x => x.Key == _建议采购Item.SKU).ToList();

                            #region 大于11个数 使用拉依达法则
                            if (refDetails.Count > 0 && refDetails.Count >= 11)
                            {
                                var LadModel = new _异常单详情();
                                LadModel.SKU = _建议采购Item.SKU;
                                var kickList = new List<decimal>();
                                var stdev = 0m;
                                var lowLimit = 0m;
                                var upLimit = 0m;
                                var tmp = refDetails.Select(x => x.Value).ToList();
                                MathHelper.SumKickOutlier(tmp, out kickList, out stdev, out lowLimit, out upLimit
                                    , OutlierRatio.Triple);


                                LadModel._标准差 = stdev;
                                LadModel._上边界 = upLimit;
                                LadModel._下边界 = lowLimit;
                                LadModel._异常值 = string.Join(",", kickList.Select(x => x.ToString()).ToArray());
                                LadModel._异常单个数 = kickList.Count;
                                LadModel._所有订单 = string.Join(",", tmp.Select(x => x.ToString()).ToArray());
                                LadModel._销售单个数 = refDetails.Count;
                                if (kickList.Count > 0)
                                {
                                    _d异常单销量和 = kickList.Sum();
                                    list异常单详情.Add(LadModel);
                                }


                            }
                            #endregion

                            #region 小于11个数 去掉最高最低取平均值
                            //if (refDetails.Count >= 3 && refDetails.Count < 11)
                            //{

                            //}
                            #endregion

                        }

                        var ComModel = new _处理对比表();
                        ComModel.SKU = _建议采购Item.SKU;
                        ComModel._原建议采购数量 = _建议采购Item._如果按之前的算法建议采购数量;
                        ComModel._处理后建议采购数量 = _建议采购Item._建议采购数量 - _d异常单销量和;

                        if (ComModel._差值 != 0)
                            list处理对比表.Add(ComModel);
                    }

                    model._Qty = _建议采购Item._建议采购数量 - _d异常单销量和;

                    if (model._Qty > 0)
                        list订单分配.Add(model);
                });

                ExportExcel(list订单分配.OrderByDescending(x => x._供应商).ToList()
                    , list销量差降详情.OrderByDescending(x => x._销量差).ToList()
                    , list异常单详情,
                    list处理对比表.OrderByDescending(x => x._差值).ToList());
            }, null);
            #endregion

        }
        #endregion

        #region 导出表格说明事件
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormHelper.GenerateTableDes(typeof(_建议采购), typeof(_下半月流水));
        }
        #endregion

        /**************** common method ****************/

        #region ExportExcel 导出Excel表格
        /// <summary>
        /// 导出Excel表格
        /// </summary>
        /// <param name="orders"></param>
        private void ExportExcel(List<_订单分配> orders, List<_销量差降详情> list销量差, List<_异常单详情> list异常单详情, List<_处理对比表> list处理对比表)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer2 = new byte[0];
            var buffer3 = new byte[0];
            var buffer4 = new byte[0];
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

            #region 详情表
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    var workbox = package.Workbook;


                    #region 销量差详情
                    {
                        var sheet1 = workbox.Worksheets.Add("销量差详情");

                        #region 标题行
                        sheet1.Cells[1, 1].Value = "SKU";
                        sheet1.Cells[1, 2].Value = "上半月销量";
                        sheet1.Cells[1, 3].Value = "下半月销量";
                        sheet1.Cells[1, 4].Value = "销量差";
                        sheet1.Cells[1, 5].Value = "变动";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 2, len = list销量差.Count; idx < len; idx++)
                        {
                            var curOrder = list销量差[idx];
                            sheet1.Cells[rowIdx, 1].Value = curOrder.SKU;
                            sheet1.Cells[rowIdx, 2].Value = curOrder._上半月销量;
                            sheet1.Cells[rowIdx, 3].Value = curOrder._下半月销量;
                            sheet1.Cells[rowIdx, 4].Value = curOrder._销量差;
                            sheet1.Cells[rowIdx, 5].Value = curOrder._变化情况;

                            rowIdx++;
                        }
                        #endregion
                    }
                    #endregion

                    #region 异常单详情
                    {
                        var sheet1 = workbox.Worksheets.Add("异常单详情");

                        #region 标题行
                        sheet1.Cells[1, 1].Value = "SKU";
                        sheet1.Cells[1, 2].Value = "标准差";
                        sheet1.Cells[1, 3].Value = "下边界";
                        sheet1.Cells[1, 4].Value = "上边界";
                        sheet1.Cells[1, 5].Value = "异常值";
                        sheet1.Cells[1, 6].Value = "异常单个数";
                        sheet1.Cells[1, 7].Value = "所有订单";
                        sheet1.Cells[1, 8].Value = "所有订单个数";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 2, len = list异常单详情.Count; idx < len; idx++)
                        {
                            var curOrder = list异常单详情[idx];
                            sheet1.Cells[rowIdx, 1].Value = curOrder.SKU;
                            sheet1.Cells[rowIdx, 2].Value = curOrder._标准差;
                            sheet1.Cells[rowIdx, 3].Value = curOrder._下边界;
                            sheet1.Cells[rowIdx, 4].Value = curOrder._上边界;
                            sheet1.Cells[rowIdx, 5].Value = curOrder._异常值;
                            sheet1.Cells[rowIdx, 6].Value = curOrder._异常单个数;
                            sheet1.Cells[rowIdx, 7].Value = curOrder._所有订单;
                            sheet1.Cells[rowIdx, 8].Value = curOrder._销售单个数;
                            rowIdx++;
                        }
                        #endregion
                    }
                    #endregion

                    #region 处理对比详情
                    {
                        var sheet1 = workbox.Worksheets.Add("处理对比详情");

                        #region 标题行
                        sheet1.Cells[1, 1].Value = "SKU";
                        sheet1.Cells[1, 2].Value = "原建议采购数量";
                        sheet1.Cells[1, 3].Value = "现在建议采购数量";
                        sheet1.Cells[1, 4].Value = "差值";
                        sheet1.Cells[1, 5].Value = "采购情况";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 2, len = list处理对比表.Count; idx < len; idx++)
                        {
                            var curOrder = list处理对比表[idx];
                            sheet1.Cells[rowIdx, 1].Value = curOrder.SKU;
                            sheet1.Cells[rowIdx, 2].Value = curOrder._原建议采购数量;
                            sheet1.Cells[rowIdx, 3].Value = curOrder._处理后建议采购数量;
                            sheet1.Cells[rowIdx, 4].Value = curOrder._差值;
                            sheet1.Cells[rowIdx, 5].Value = curOrder._采购情况;
                            rowIdx++;
                        }
                        #endregion
                    }
                    #endregion

                    buffer4 = package.GetAsByteArray();
                }
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
                    var FileName4 = Path.Combine(savePath, saveFilName + "(详情表).xlsx");
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

        [ExcelTable("各平台近期销量表")]
        class _建议采购
        {
            private string _SKU;

            [ExcelColumn("SKU码")]
            public string SKU
            {
                get
                {
                    return _SKU;
                }
                set
                {
                    _SKU = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("供应商")]
            public string _供应商 { get; set; }

            [ExcelColumn("采购员")]
            public string _采购员 { get; set; }

            [ExcelColumn("业绩归属2")]
            public string _开发 { get; set; }

            [ExcelColumn("可用数量")]
            public decimal _可用数量 { get; set; }

            [ExcelColumn("采购未入库")]
            public decimal _采购未入库 { get; set; }

            [ExcelColumn("缺货及未派单数量")]
            public decimal _缺货及未派单数量 { get; set; }

            [ExcelColumn("商品成本单价")]
            public decimal _商品成本单价 { get; set; }

            [ExcelColumn("30天销量")]
            public decimal _30天销量 { get; set; }

            [ExcelColumn("15天销量")]
            public decimal _15天销量 { get; set; }

            [ExcelColumn("5天销量")]
            public decimal _5天销量 { get; set; }

            [ExcelColumn("预警销售天数")]
            public decimal _预警销售天数 { get; set; }

            [ExcelColumn("采购到货天数")]
            public decimal _采购到货天数 { get; set; }

            public decimal _上半月销量 { get { return _30天销量 - _15天销量; } }

            public decimal _下半月销量 { get { return _15天销量; } }

            public decimal _上下半月销量差 { get { return _下半月销量 - _上半月销量; } }

            public decimal _全月日平均销量
            {
                get
                {
                    var tmp = (_30天销量 / 30 + _15天销量 / 15 + _5天销量 / 5);
                    return tmp != 0 ? tmp / 3 : 0;
                }
            }

            public decimal _下半月日平均销量
            {
                get
                {
                    var tmp = (_15天销量 / 15);
                    return tmp != 0 ? tmp : 0;
                }
            }

            public decimal _销量差 { get; set; }
            public decimal _平均日销量
            {
                get
                {
                    if (_上下半月销量差 >= _销量差)
                    {
                        return _下半月日平均销量;
                    }
                    else
                    {
                        return _全月日平均销量;
                    }
                }
            }

            public decimal _建议采购数量
            {
                get
                {
                    var _库存上限 = _预警销售天数 * _平均日销量;
                    var _库存下限 = _采购到货天数 * _平均日销量;
                    return Convert.ToDecimal(Helper.CalAmount(Convert.ToDouble(_库存上限 + _库存下限 - _可用数量 - _采购未入库 + _缺货及未派单数量)));
                }
            }

            public decimal _如果按之前的算法建议采购数量
            {
                get
                {
                    var _库存上限 = _预警销售天数 * _全月日平均销量;
                    var _库存下限 = _采购到货天数 * _全月日平均销量;
                    return Convert.ToDecimal(Helper.CalAmount(Convert.ToDouble(_库存上限 + _库存下限 - _可用数量 - _采购未入库 + _缺货及未派单数量)));
                }
            }

        }

        [ExcelTable("下半月流水表")]
        class _下半月流水
        {
            private string _Orders;
            private List<KeyValuePair<string, decimal>> _Details = new List<KeyValuePair<string, decimal>>();

            [ExcelColumn("商品信息")]
            public string _商品信息
            {
                set
                {
                    _Orders = value;
                }
            }

            public List<KeyValuePair<string, decimal>> _订单详情
            {
                get
                {
                    if (!string.IsNullOrEmpty(_Orders))
                    {
                        var list = _Orders.Split(';').Where(x => !string.IsNullOrEmpty(x)).ToList();
                        list.ForEach(fusItem =>
                        {
                            if (fusItem == "FEDB3B22*1")
                            {

                            }
                            try
                            {
                                var idx = fusItem.LastIndexOf('*');
                                if (idx != -1)
                                {
                                    var key = fusItem.Substring(0, idx);
                                    var val = fusItem.Substring(idx + 1, fusItem.Length - idx - 1);
                                    _Details.Add(new KeyValuePair<string, decimal>(key, Convert.ToDecimal(val)));
                                }
                            }
                            catch (Exception ex)
                            {

                                var aa = ex.Message;
                            }
                        });
                    }
                    return _Details;
                }
            }

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

        class _销量差降详情
        {
            public string SKU { get; set; }

            public decimal _上半月销量 { get; set; }

            public decimal _下半月销量 { get; set; }

            public decimal _销量差
            {
                get
                {
                    var tmp = _下半月销量 - _上半月销量;
                    return tmp > 0 ? tmp : -tmp;
                }
            }

            public string _变化情况
            {
                get
                {
                    var tmp = _下半月销量 - _上半月销量;
                    return tmp > 0 ? "上升" : "下降";
                }
            }
        }

        class _异常单详情
        {
            public string SKU { get; set; }

            public decimal _标准差 { get; set; }

            public decimal _下边界 { get; set; }

            public decimal _上边界 { get; set; }

            public string _异常值 { get; set; }

            public int _异常单个数 { get; set; }

            public string _所有订单 { get; set; }

            public int _销售单个数 { get; set; }
        }

        class _处理对比表
        {
            public string SKU { get; set; }

            public decimal _原建议采购数量 { get; set; }

            public decimal _处理后建议采购数量 { get; set; }

            public decimal _差值
            {
                get
                {
                    var tmp = _原建议采购数量 - _处理后建议采购数量;
                    return tmp > 0 ? tmp : -tmp;
                }
            }

            public string _采购情况
            {
                get
                {
                    var tmp = _原建议采购数量 - _处理后建议采购数量;
                    return tmp > 0 ? "采购量减少" : "采购量增加";
                }
            }
        }
    }
}
