using CommonLibs;
using Gadget.Libs;
using LinqToExcel.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Gadget
{
    public partial class _批量生成采购订单_参数可调 : Form
    {
        public _批量生成采购订单_参数可调()
        {
            InitializeComponent();
        }

        private void _批量生成采购订单_参数可调_Load(object sender, EventArgs e)
        {
            //txt库存预警原表.Text = @"C:\Users\Leon\Desktop\8月18号\库存预警8月18号.csv";
            //txt库存预警中位数.Text = @"C:\Users\Leon\Desktop\8月18号\库存预警中位数月18号.csv";
            //txt每月流水.Text = @"C:\Users\Leon\Desktop\8月18号\近30天销量8月18号.csv";
            //btn处理数据.Enabled = true;
        }
        /**************** button event ****************/

        #region 上传库存预警原表
        private void btn上传库存预警原表_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt库存预警原表, () =>
            {
                btn处理数据.Enabled = CanCalcu();
            });
        }
        #endregion

        #region 上传库存预警中位数
        private void btn上传库存预警中位数_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt库存预警中位数, () =>
            {
                btn处理数据.Enabled = CanCalcu();
            });
        }
        #endregion

        #region 上传每月流水
        private void btn上传每月流水_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt每月流水, () =>
            {
                btn处理数据.Enabled = CanCalcu();
            });
        }
        #endregion

        private void btn处理数据_Click(object sender, EventArgs e)
        {
            var list库存预警原表 = new List<_库存预警>();
            var list库存预警中位数 = new List<_库存预警>();
            var list每月流水 = new List<_每月流水>();
            var list两表都有的SKUs = new List<string>();
            var list缺货所有SKUs = new List<string>();
            var list处理结果 = new List<_订单分配>();

            int iDay1 = Convert.ToInt32(nupDay1.Value);
            int iDay2 = Convert.ToInt32(nupDay2.Value);
            int iDay3 = Convert.ToInt32(nupDay3.Value);

            #region 读取数据
            var actReadData = new Action(() =>
            {
                var strError = string.Empty;
                ShowMsg("开始读取预警原表数据");
                FormHelper.ReadCSVFile(txt库存预警原表.Text, ref list库存预警原表, ref strError);
                ShowMsg("开始读取预警中位数表数据");
                FormHelper.ReadCSVFile(txt库存预警中位数.Text, ref list库存预警中位数, ref strError);
                ShowMsg("开始读取每月流水数据");
                FormHelper.ReadCSVFile(txt每月流水.Text, ref list每月流水, ref strError);

                ShowMsg("开始过滤每月流水中不需要的数据");
                //过滤掉流水表里面不需要的sku数据,因为该表太大
                var sku_yb = list库存预警原表.Select(x => x.SKU).ToList();
                var sku_zw = list库存预警中位数.Select(x => x.SKU).ToList();

                var q = from t1 in sku_yb
                        join t2 in sku_zw on t1 equals t2
                        select t1;
                list两表都有的SKUs = q.ToList();

                sku_yb.AddRange(sku_zw);
                list缺货所有SKUs = sku_yb.Select(x => x).Distinct().ToList();
                list每月流水 = list每月流水.Where(x => list缺货所有SKUs.Contains(x.SKU)).ToList();

            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                ShowMsg("正在处理数据");
                list缺货所有SKUs.ForEach(sku =>
                {
                    //if (string.IsNullOrWhiteSpace(sku))
                    //{

                    //}
                    var model = new _订单分配();
                    model._SKU = sku;

                    var refer库存预警原表 = list库存预警原表.FirstOrDefault(x => x.SKU == sku);
                    var refer库存预警中位数 = list库存预警中位数.FirstOrDefault(x => x.SKU == sku);
                    var refer流水 = list每月流水.FirstOrDefault();
                    bool bShareSKU = refer库存预警原表 != null && refer库存预警中位数 != null;
                    //公共信息
                    var common = refer库存预警原表 != null ? refer库存预警原表 : refer库存预警中位数;
                    model._预计可用库存 = common._预计可用库存;
                    model._供应商 = common._供应商;
                    model._采购员 = common._采购员;
                    model._含税单价 = common._商品成本单价;
                    model._制单人 = common._采购员;
                    //特有信息
                    model._原表格导出的建议采购数量_原预警表 = refer库存预警原表 != null ? refer库存预警原表._表格导出的原始建议采购 : 0;
                    model._原表格导出的建议采购数量_中位数表 = refer库存预警中位数 != null ? refer库存预警中位数._表格导出的原始建议采购 : 0;
                    //两种算法的结果
                    {
                        var list1 = refer流水._月销量流水.Take(iDay1).ToList();
                        var list2 = refer流水._月销量流水.Take(iDay2).ToList();
                        var list3 = refer流水._月销量流水.Take(iDay3).ToList();

                        model._计算后的建议采购数量_原先算法 = Math.Round(common._原先计算方式(list1, list2, list3), 0);
                        model._计算后的建议采购数量_中位数算法 = Math.Round(common._中位数计算方式(list1, list2, list3), 0);
                    }

                    /*
                     * 建议采购量以库存预警-中位数为主，但是需要修改以下条件：
                     *（1）如果建议采购量 和 预计可用数量 一致或者差不多（建议采购量刚好够补缺货订单），那么建议采购最终数量=                 
                     * （库存预警建议采购量+库存预警中位数建议采购量）/2
                     *（2）当库存预警建议采购量<库存预警中位数建议采购量，以库存预警建议采购量为主
                     *（3）当库存预警中位数 标记建议采购，但是建议采购量为0时，对应的SKU 以库存预警的建议采购量为主
                     * （这个是普源数据的问题，已经联系崔总做修改，待回复）
                     *（4）当库存预警中位数建议采购量<库存预警建议采购量，商品成本单价小于1元的，最终建议采购量=（库存预警建议采购量+库存预警中位数建议采购量）/2
                     *（5）商品单价低于10元的，建议采购量小于5个的，最终建议采购量为 5个。(已经在建议采购做判断了)
                     */

                    //中位数算法为主
                    model._Qty = model._计算后的建议采购数量_中位数算法;

                    if (bShareSKU)
                    {
                        /*
                         * （2）当库存预警建议采购量<库存预警中位数建议采购量，以库存预警建议采购量为主
                         */
                        if (model._计算后的建议采购数量_原先算法 < model._计算后的建议采购数量_中位数算法)
                        {
                            model._Qty = model._计算后的建议采购数量_原先算法;
                        }
                        else
                        {
                            if (common._商品成本单价 < 1)
                            {
                                decimal culc = (model._计算后的建议采购数量_原先算法 + model._计算后的建议采购数量_中位数算法) / 2;
                                model._Qty = Math.Round(culc, 0);
                            }
                        }
                    }

                    list处理结果.Add(model);

                });

                ExportExcel(list处理结果);
            }, null);
            #endregion

        }

        #region 导出说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormHelper.GenerateTableDes(typeof(_库存预警), typeof(_每月流水));
        }
        #endregion

        #region ExportExcel 导出Excel表格
        /// <summary>
        /// 导出Excel表格
        /// </summary>
        /// <param name="orders"></param>
        private void ExportExcel(List<_订单分配> orders)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer2 = new byte[0];
            var buffer3 = new byte[0];
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

                sheet1.Cells[1, 15].Value = "预计可用库存";
                sheet1.Cells[1, 16].Value = "表格导出建议采购(原预警)";
                sheet1.Cells[1, 17].Value = "表格导出建议采购(中位数预警)";
                sheet1.Cells[1, 18].Value = "原先算法建议采购";
                sheet1.Cells[1, 19].Value = "中位数算法建议采购";

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
                        if (curOrder._是否紧急单)
                            sheet1.Cells[rowIdx, 5].Value = "紧急";
                        sheet1.Cells[rowIdx, 7].Value = curOrder._采购员;
                        sheet1.Cells[rowIdx, 8].Value = curOrder._含税单价;
                        sheet1.Cells[rowIdx, 10].Value = "支付宝";
                        sheet1.Cells[rowIdx, 11].Value = curOrder._制单人;
                        //sheet1.Cells[rowIdx, 15].Value = curOrder._对应供应商采购金额;

                        sheet1.Cells[rowIdx, 15].Value = curOrder._预计可用库存;
                        sheet1.Cells[rowIdx, 16].Value = curOrder._原表格导出的建议采购数量_原预警表;
                        sheet1.Cells[rowIdx, 17].Value = curOrder._原表格导出的建议采购数量_中位数表;
                        sheet1.Cells[rowIdx, 18].Value = curOrder._计算后的建议采购数量_原先算法;
                        sheet1.Cells[rowIdx, 19].Value = curOrder._计算后的建议采购数量_中位数算法;
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


                sheet1.Cells[1, 15].Value = "预计可用库存";
                sheet1.Cells[1, 16].Value = "表格导出建议采购(原预警)";
                sheet1.Cells[1, 17].Value = "表格导出建议采购(中位数预警)";
                sheet1.Cells[1, 18].Value = "原先算法建议采购";
                sheet1.Cells[1, 19].Value = "中位数算法建议采购";
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = devOrder.Count; idx < len; idx++, rowIdx++)
                {
                    var curOrder = devOrder[idx];
                    sheet1.Cells[rowIdx, 1].Value = curOrder._供应商;
                    sheet1.Cells[rowIdx, 2].Value = curOrder._SKU;
                    sheet1.Cells[rowIdx, 3].Value = curOrder._Qty;
                    if (curOrder._是否紧急单)
                        sheet1.Cells[rowIdx, 5].Value = "紧急";
                    sheet1.Cells[rowIdx, 7].Value = curOrder._采购员;
                    sheet1.Cells[rowIdx, 8].Value = curOrder._含税单价;
                    sheet1.Cells[rowIdx, 10].Value = "支付宝";
                    sheet1.Cells[rowIdx, 11].Value = curOrder._制单人;
                    //sheet1.Cells[rowIdx, 15].Value = curOrder._对应供应商采购金额;
                    sheet1.Cells[rowIdx, 15].Value = curOrder._预计可用库存;
                    sheet1.Cells[rowIdx, 16].Value = curOrder._原表格导出的建议采购数量_原预警表;
                    sheet1.Cells[rowIdx, 17].Value = curOrder._原表格导出的建议采购数量_中位数表;
                    sheet1.Cells[rowIdx, 18].Value = curOrder._计算后的建议采购数量_原先算法;
                    sheet1.Cells[rowIdx, 19].Value = curOrder._计算后的建议采购数量_中位数算法;
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

        /**************** common method ****************/

        #region CanCalcu 判断是否可以开始计算数据
        private bool CanCalcu()
        {
            var b上传预警原表 = !string.IsNullOrWhiteSpace(txt库存预警原表.Text);
            var b上传预警中位数 = !string.IsNullOrWhiteSpace(txt库存预警中位数.Text);
            var b上传每月流水 = !string.IsNullOrWhiteSpace(txt每月流水.Text);
            return b上传预警原表 && b上传预警中位数 && b上传每月流水;
        }
        #endregion

        #region Calcu中位数 获取中位数
        /// <summary>
        /// 获取中位数
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        public static decimal Calcu中位数(List<decimal> datas)
        {
            var count = datas.Count;
            if (count > 0)
            {
                var isEven = count % 2 == 0;//偶数
                var sorts = datas.OrderBy(x => x).ToList();
                if (isEven)
                {
                    var n1 = sorts[count / 2];
                    var n2 = sorts[count / 2 - 1];
                    return (n1 + n2) / 2;
                }
                else
                {
                    return sorts[(count - 1) / 2];
                }
            }
            return 0m;
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

        [ExcelTable("库存预警")]
        class _库存预警
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

            [ExcelColumn("缺货占用数量")]
            public decimal _缺货占用数量 { get; set; }

            [ExcelColumn("商品成本单价")]
            public decimal _商品成本单价 { get; set; }

            [ExcelColumn("预警销售天数")]
            public decimal _预警销售天数 { get; set; }

            [ExcelColumn("采购到货天数")]
            public decimal _采购到货天数 { get; set; }

            [ExcelColumn("预计可用库存")]
            public decimal _预计可用库存 { get; set; }

            [ExcelColumn("建议采购数量")]
            public decimal _表格导出的原始建议采购 { get; set; }

            /**************** virtual ****************/

            public virtual bool _是否紧急单
            {
                get
                {
                    return _可用数量 - _缺货及未派单数量 + _缺货占用数量 <= 0;
                }
            }

            public decimal _原先计算方式(List<decimal> list1, List<decimal> list2, List<decimal> list3)
            {
                int count1 = list1.Count;
                int count2 = list2.Count;
                int count3 = list3.Count;
                decimal s1 = list1.Sum();
                decimal s2 = list2.Sum();
                decimal s3 = list3.Sum();

                var _日平均销量 = (s1 / count1 + s2 / count2 + s3 / count3) / 3;
                var _库存上限 = _预警销售天数 * _日平均销量;
                var _库存下限 = _采购到货天数 * _日平均销量;

                return _库存上限 + _库存下限 - _可用数量 - _采购未入库 + _缺货及未派单数量;
            }

            public decimal _中位数计算方式(List<decimal> list1, List<decimal> list2, List<decimal> list3)
            {
                var d1 = Calcu中位数(list1);
                var d2 = Calcu中位数(list2);
                var d3 = Calcu中位数(list3);

                var _日平均销量 = (d1 + d2 + d3) / 3;
                var _库存上限 = _预警销售天数 * _日平均销量;
                var _库存下限 = _采购到货天数 * _日平均销量;

                return _库存上限 + _库存下限 - _可用数量 - _采购未入库 + _缺货及未派单数量;
            }

        }

        [ExcelTable("每月流水")]
        class _每月流水
        {
            private string _SKU;

            [ExcelColumn("SKU")]
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

            [ExcelColumn("今天销量")]
            public decimal _今天销量 { get; set; }
            [ExcelColumn("今天往前1天")]
            public decimal _今天往前1天 { get; set; }
            [ExcelColumn("今天往前2天")]
            public decimal _今天往前2天 { get; set; }
            [ExcelColumn("今天往前3天")]
            public decimal _今天往前3天 { get; set; }
            [ExcelColumn("今天往前4天")]
            public decimal _今天往前4天 { get; set; }
            [ExcelColumn("今天往前5天")]
            public decimal _今天往前5天 { get; set; }
            [ExcelColumn("今天往前6天")]
            public decimal _今天往前6天 { get; set; }
            [ExcelColumn("今天往前7天")]
            public decimal _今天往前7天 { get; set; }
            [ExcelColumn("今天往前8天")]
            public decimal _今天往前8天 { get; set; }
            [ExcelColumn("今天往前9天")]
            public decimal _今天往前9天 { get; set; }
            [ExcelColumn("今天往前10天")]
            public decimal _今天往前10天 { get; set; }

            [ExcelColumn("今天往前11天")]
            public decimal _今天往前11天 { get; set; }
            [ExcelColumn("今天往前12天")]
            public decimal _今天往前12天 { get; set; }
            [ExcelColumn("今天往前13天")]
            public decimal _今天往前13天 { get; set; }
            [ExcelColumn("今天往前14天")]
            public decimal _今天往前14天 { get; set; }
            [ExcelColumn("今天往前15天")]
            public decimal _今天往前15天 { get; set; }
            [ExcelColumn("今天往前16天")]
            public decimal _今天往前16天 { get; set; }
            [ExcelColumn("今天往前17天")]
            public decimal _今天往前17天 { get; set; }
            [ExcelColumn("今天往前18天")]
            public decimal _今天往前18天 { get; set; }
            [ExcelColumn("今天往前19天")]
            public decimal _今天往前19天 { get; set; }
            [ExcelColumn("今天往前20天")]
            public decimal _今天往前20天 { get; set; }

            [ExcelColumn("今天往前21天")]
            public decimal _今天往前21天 { get; set; }
            [ExcelColumn("今天往前22天")]
            public decimal _今天往前22天 { get; set; }
            [ExcelColumn("今天往前23天")]
            public decimal _今天往前23天 { get; set; }
            [ExcelColumn("今天往前24天")]
            public decimal _今天往前24天 { get; set; }
            [ExcelColumn("今天往前25天")]
            public decimal _今天往前25天 { get; set; }
            [ExcelColumn("今天往前26天")]
            public decimal _今天往前26天 { get; set; }
            [ExcelColumn("今天往前27天")]
            public decimal _今天往前27天 { get; set; }
            [ExcelColumn("今天往前28天")]
            public decimal _今天往前28天 { get; set; }
            [ExcelColumn("今天往前29天")]
            public decimal _今天往前29天 { get; set; }

            public List<decimal> _月销量流水
            {
                get
                {
                    return new List<decimal>()
                    {
                        _今天销量,
                        _今天往前1天,
                        _今天往前2天,
                        _今天往前3天,
                        _今天往前4天,
                        _今天往前5天,
                        _今天往前6天,
                        _今天往前7天,
                        _今天往前8天,
                        _今天往前9天,
                        _今天往前10天,
                        _今天往前11天,
                        _今天往前12天,
                        _今天往前13天,
                        _今天往前14天,
                        _今天往前15天,
                        _今天往前16天,
                        _今天往前17天,
                        _今天往前18天,
                        _今天往前19天,
                        _今天往前20天,
                        _今天往前21天,
                        _今天往前22天,
                        _今天往前23天,
                        _今天往前24天,
                        _今天往前25天,
                        _今天往前26天,
                        _今天往前27天,
                        _今天往前28天,
                        _今天往前29天
                    };
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
            public decimal _物流费 { get; set; }
            public string _付款方式 { get; set; }
            public string _制单人 { get; set; }
            public string _到货日期 { get; set; }
            public string _1688单号 { get; set; }
            public decimal _预付款 { get; set; }
            public decimal _对应供应商采购金额 { get; set; }
            public decimal _预计可用库存 { get; set; }
            public bool _是否紧急单 { get; set; }

            public decimal _原表格导出的建议采购数量_原预警表 { get; set; }
            public decimal _原表格导出的建议采购数量_中位数表 { get; set; }
            public decimal _计算后的建议采购数量_原先算法 { get; set; }
            public decimal _计算后的建议采购数量_中位数算法 { get; set; }



        }


    }
}
