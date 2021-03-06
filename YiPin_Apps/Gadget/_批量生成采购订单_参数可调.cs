﻿using CommonLibs;
using Gadget.Libs;
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
    public partial class _批量生成采购订单_参数可调 : Form
    {
        public _批量生成采购订单_参数可调()
        {
            InitializeComponent();
        }

        private void _批量生成采购订单_参数可调_Load(object sender, EventArgs e)
        {
            //txt库存预警原表.Text = @"C:\Users\Leon\Desktop\CSV\库存预警建议采购0822调参数前.csv";
            //txt库存预警中位数.Text = @"C:\Users\Leon\Desktop\CSV\库存预警中位数8月22号晚上19点半调参数前.csv";
            //txt每月流水.Text = @"C:\Users\Leon\Desktop\CSV\近30天销量8月22号19点半.csv";
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
            var list计算详情 = new List<_计算详情>();
            var list异常SKUs = new List<string>();

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
                foreach (var sku in list缺货所有SKUs)
                {
                    //if (sku== "AMKA1A03-WR")
                    //{

                    //}
                    var _计算详情Model = new _计算详情();
                    var model = new _订单分配();
                    model._SKU = sku;
                    _计算详情Model._SKU = sku;

                    var refer库存预警原表 = list库存预警原表.FirstOrDefault(x => x.SKU == sku);
                    var refer库存预警中位数 = list库存预警中位数.FirstOrDefault(x => x.SKU == sku);
                    var refer流水 = list每月流水.FirstOrDefault(x => x.SKU == sku);

                    if (refer流水 == null)
                    {
                        list异常SKUs.Add(sku);
                        continue;
                    }

                    bool bShareSKU = refer库存预警原表 != null && refer库存预警中位数 != null;
                    //公共信息
                    var common = refer库存预警原表 != null ? refer库存预警原表 : refer库存预警中位数;
                    model._预计可用库存 = common._预计可用库存;
                    model._供应商 = common._供应商;
                    model._采购员 = common._采购员;
                    model._含税单价 = common._商品成本单价;
                    model._制单人 = common._采购员;

                    //详情基本信息
                    _计算详情Model._采购员 = common._采购员;
                    _计算详情Model._含税单价 = common._商品成本单价;
                    _计算详情Model._可用数量 = common._可用数量;
                    _计算详情Model._采购未入库 = common._采购未入库;
                    _计算详情Model._缺货及未派单数量 = common._缺货及未派单数量;
                    _计算详情Model._预计可用库存 = common._预计可用库存;
                    _计算详情Model._预警销售天数 = common._预警销售天数;
                    _计算详情Model._采购到货天数 = common._采购到货天数;
                    _计算详情Model._缺货占用数量 = common._缺货占用数量;

                    //特有信息
                    model._原表格导出的建议采购数量_原预警表 = refer库存预警原表 != null ? refer库存预警原表._表格导出的原始建议采购 : 0;
                    model._原表格导出的建议采购数量_中位数表 = refer库存预警中位数 != null ? refer库存预警中位数._表格导出的原始建议采购 : 0;

                    #region 两种算法的结果
                    {
                        var list1 = refer流水._月销量流水.Take(iDay1).ToList();
                        var list2 = refer流水._月销量流水.Take(iDay2).ToList();
                        var list3 = refer流水._月销量流水.Take(iDay3).ToList();

                        #region 记录销量详情
                        {
                            _计算详情Model._参数1详情 = string.Join(",", list1);
                            _计算详情Model._参数2详情 = string.Join(",", list2);
                            _计算详情Model._参数3详情 = string.Join(",", list3);
                        }
                        #endregion

                        #region 原先算法
                        {
                            var detail = new _计算过程();
                            common._原先计算方式(list1, list2, list3, ref detail);
                            model._计算后的建议采购数量_原先算法 = detail._建议采购;

                            _计算详情Model._库存上限_原先算法 = detail._库存上限;
                            _计算详情Model._库存下限_原先算法 = detail._库存下限;
                            _计算详情Model._日平均销量_原先算法 = detail._日平均销量;
                            _计算详情Model._计算后的建议采购数量_原先算法 = detail._建议采购;

                            _计算详情Model._参数1_销量总和 = detail._参数1_销量总和;
                            _计算详情Model._参数2_销量总和 = detail._参数2_销量总和;
                            _计算详情Model._参数3_销量总和 = detail._参数3_销量总和;
                        }
                        #endregion

                        #region 中位数算法
                        {
                            var detail = new _计算过程();
                            common._中位数计算方式(list1, list2, list3, ref detail);
                            model._计算后的建议采购数量_中位数算法 = detail._建议采购;

                            _计算详情Model._库存上限_中位数 = detail._库存上限;
                            _计算详情Model._库存下限_中位数 = detail._库存下限;
                            _计算详情Model._日平均销量_中位数 = detail._日平均销量;
                            _计算详情Model._计算后的建议采购数量_中位数 = detail._建议采购;

                            _计算详情Model._参数1_中位数 = detail._参数1_中位数;
                            _计算详情Model._参数2_中位数 = detail._参数2_中位数;
                            _计算详情Model._参数3_中位数 = detail._参数3_中位数;
                        }
                        #endregion

                        _计算详情Model._原表格导出的建议采购数量_原预警表 = model._原表格导出的建议采购数量_原预警表;
                        _计算详情Model._原表格导出的建议采购数量_中位数表 = model._原表格导出的建议采购数量_中位数表;

                    }
                    #endregion

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

                    _计算详情Model.Qty = model._Qty;

                    list处理结果.Add(model);
                    list计算详情.Add(_计算详情Model);
                }


                ExportExcel(list处理结果, list计算详情, list异常SKUs, iDay1, iDay2, iDay3);
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
        /// <param name="list"></param>
        private void ExportExcel(List<_订单分配> list, List<_计算详情> details, List<string> errorSKUs, int n1, int n2, int n3)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer2 = new byte[0];
            var buffer3 = new byte[0];
            var buffer4 = new byte[0];
            var devList = new List<_订单分配>();


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

                //sheet1.Cells[1, 15].Value = "预计可用库存";
                //sheet1.Cells[1, 16].Value = "表格导出建议采购(原预警)";
                //sheet1.Cells[1, 17].Value = "表格导出建议采购(中位数预警)";
                //sheet1.Cells[1, 18].Value = "原先算法建议采购";
                //sheet1.Cells[1, 19].Value = "中位数算法建议采购";

                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = list.Count; idx < len; idx++)
                {
                    var curOrder = list[idx];
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

                        //sheet1.Cells[rowIdx, 15].Value = curOrder._预计可用库存;
                        //sheet1.Cells[rowIdx, 16].Value = curOrder._原表格导出的建议采购数量_原预警表;
                        //sheet1.Cells[rowIdx, 17].Value = curOrder._原表格导出的建议采购数量_中位数表;
                        //sheet1.Cells[rowIdx, 18].Value = curOrder._计算后的建议采购数量_原先算法;
                        //sheet1.Cells[rowIdx, 19].Value = curOrder._计算后的建议采购数量_中位数算法;
                        rowIdx++;
                    }
                    else
                    {
                        devList.Add(curOrder);
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
                buyers = list.Where(x => !string.IsNullOrEmpty(x._采购员)).Select(x => x._采购员).Distinct().ToList();
                for (int idx = 0, len = buyers.Count, rowIdx = 2; idx < len; idx++, rowIdx++)
                {
                    var curBuyerName = buyers[idx];
                    var refOrders = list.Where(m => m._采购员 == curBuyerName).ToList();
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


                //sheet1.Cells[1, 15].Value = "预计可用库存";
                //sheet1.Cells[1, 16].Value = "表格导出建议采购(原预警)";
                //sheet1.Cells[1, 17].Value = "表格导出建议采购(中位数预警)";
                //sheet1.Cells[1, 18].Value = "原先算法建议采购";
                //sheet1.Cells[1, 19].Value = "中位数算法建议采购";
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = devList.Count; idx < len; idx++, rowIdx++)
                {
                    var curOrder = devList[idx];
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
                    //sheet1.Cells[rowIdx, 15].Value = curOrder._预计可用库存;
                    //sheet1.Cells[rowIdx, 16].Value = curOrder._原表格导出的建议采购数量_原预警表;
                    //sheet1.Cells[rowIdx, 17].Value = curOrder._原表格导出的建议采购数量_中位数表;
                    //sheet1.Cells[rowIdx, 18].Value = curOrder._计算后的建议采购数量_原先算法;
                    //sheet1.Cells[rowIdx, 19].Value = curOrder._计算后的建议采购数量_中位数算法;
                }
                #endregion


                buffer3 = package.GetAsByteArray();
            }
            #endregion

            #region 计算详情
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    var workbox = package.Workbook;

                    #region 详情
                    {
                        var sheet1 = workbox.Worksheets.Add("详情");

                        #region 标题行
                        sheet1.Cells[2, 1].Value = "SKU";
                        sheet1.Cells[2, 2].Value = "采购员";
                        sheet1.Cells[2, 3].Value = "含税单价";
                        sheet1.Cells[2, 4].Value = "可用数量";
                        sheet1.Cells[2, 5].Value = "采购未入库";
                        sheet1.Cells[2, 6].Value = "缺货及未派单数量";
                        sheet1.Cells[2, 7].Value = "预计可用库存";
                        sheet1.Cells[2, 8].Value = "预警销售天数";
                        sheet1.Cells[2, 9].Value = "采购到货天数";
                        sheet1.Cells[2, 10].Value = "缺货占用数量";

                        using (var rng = sheet1.Cells[1, 1, 1, 10])
                        {
                            rng.Value = "基本信息";
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        using (var rng = sheet1.Cells[1, 1, 2, 10])
                        {
                            rng.Style.WrapText = true;//自动换行
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            var colFromHex = System.Drawing.ColorTranslator.FromHtml("#11A7F5");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        sheet1.Cells[2, 11].Value = "库存上限";
                        sheet1.Cells[2, 12].Value = "库存下限";
                        sheet1.Cells[2, 13].Value = "日平均销量";
                        sheet1.Cells[2, 14].Value = "小工具建议采购";
                        sheet1.Cells[2, 15].Value = "普元建议采购";
                        using (var rng = sheet1.Cells[1, 11, 1, 15])
                        {
                            rng.Value = "原始算法";
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        using (var rng = sheet1.Cells[1,11, 2, 15])
                        {
                            rng.Style.WrapText = true;//自动换行
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            var colFromHex = System.Drawing.ColorTranslator.FromHtml("#92D050");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        sheet1.Cells[2, 16].Value = "库存上限";
                        sheet1.Cells[2, 17].Value = "库存下限";
                        sheet1.Cells[2, 18].Value = "日平均销量";
                        sheet1.Cells[2, 19].Value = "小工具建议采购";
                        sheet1.Cells[2, 20].Value = "普元建议采购";
                        using (var rng = sheet1.Cells[1, 16, 1, 20])
                        {
                            rng.Value = "中位数算法";
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        using (var rng = sheet1.Cells[1, 16, 2, 20])
                        {
                            rng.Style.WrapText = true;//自动换行
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            var colFromHex = System.Drawing.ColorTranslator.FromHtml("#E26B0A");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        using (var rng = sheet1.Cells[1, 21, 2, 21])
                        {
                            rng.Value = "Qty";
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            var colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFF00");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }


                        sheet1.Cells[2, 22].Value = "销量总和";
                        sheet1.Cells[2, 23].Value = "中位数";
                        using (var rng = sheet1.Cells[1, 22, 1, 23])
                        {
                            rng.Value = string.Format("{0}天销量", n1);
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        using (var rng = sheet1.Cells[1, 22, 2, 23])
                        {
                            rng.Style.WrapText = true;//自动换行
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            var colFromHex = System.Drawing.ColorTranslator.FromHtml("#CCC0DA");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        sheet1.Cells[2, 24].Value = "销量总和";
                        sheet1.Cells[2, 25].Value = "中位数";
                        using (var rng = sheet1.Cells[1, 24, 1, 25])
                        {
                            rng.Value = string.Format("{0}天销量", n2);
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        using (var rng = sheet1.Cells[1, 24, 2, 25])
                        {
                            rng.Style.WrapText = true;//自动换行
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            var colFromHex = System.Drawing.ColorTranslator.FromHtml("#A6A6A6");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        sheet1.Cells[2, 26].Value = "销量总和";
                        sheet1.Cells[2, 27].Value = "中位数";
                        using (var rng = sheet1.Cells[1, 26, 1, 27])
                        {
                            rng.Value = string.Format("{0}天销量", n3);
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        using (var rng = sheet1.Cells[1, 26, 2, 27])
                        {
                            rng.Style.WrapText = true;//自动换行
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            var colFromHex = System.Drawing.ColorTranslator.FromHtml("#948A54");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        using (var rng = sheet1.Cells[1, 28, 2, 28])
                        {
                            rng.Value = string.Format("{0}天销量详情", n1);
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }

                        using (var rng = sheet1.Cells[1, 29, 2, 29])
                        {
                            rng.Value = string.Format("{0}天销量详情", n2);
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }

                        using (var rng = sheet1.Cells[1, 30, 2, 30])
                        {
                            rng.Value = string.Format("{0}天销量详情", n3);
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        #endregion

                        #region 数据行
                        for (int idx = 0, len = details.Count, rowIdx = 3; idx < len; idx++, rowIdx++)
                        {
                            var data = details[idx];
                            sheet1.Cells[rowIdx, 1].Value = data._SKU;
                            sheet1.Cells[rowIdx, 2].Value = data._采购员;
                            sheet1.Cells[rowIdx, 3].Value = data._含税单价;
                            sheet1.Cells[rowIdx, 4].Value = data._可用数量;
                            sheet1.Cells[rowIdx, 5].Value = data._采购未入库;
                            sheet1.Cells[rowIdx, 6].Value = data._缺货及未派单数量;
                            sheet1.Cells[rowIdx, 7].Value = data._预计可用库存;
                            sheet1.Cells[rowIdx, 8].Value = data._预警销售天数;
                            sheet1.Cells[rowIdx, 9].Value = data._采购到货天数;
                            sheet1.Cells[rowIdx, 10].Value = data._缺货占用数量;


                            sheet1.Cells[rowIdx, 11].Value = data._库存上限_原先算法;
                            sheet1.Cells[rowIdx, 12].Value = data._库存下限_原先算法;
                            sheet1.Cells[rowIdx, 13].Value = data._日平均销量_原先算法;
                            sheet1.Cells[rowIdx, 14].Value = data._计算后的建议采购数量_原先算法;
                            sheet1.Cells[rowIdx, 15].Value = data._原表格导出的建议采购数量_原预警表;

                            sheet1.Cells[rowIdx, 16].Value = data._库存上限_中位数;
                            sheet1.Cells[rowIdx, 17].Value = data._库存下限_中位数;
                            sheet1.Cells[rowIdx, 18].Value = data._日平均销量_中位数;
                            sheet1.Cells[rowIdx, 19].Value = data._计算后的建议采购数量_中位数;
                            sheet1.Cells[rowIdx, 20].Value = data._原表格导出的建议采购数量_中位数表;

                            sheet1.Cells[rowIdx, 21].Value = data.Qty;


                            sheet1.Cells[rowIdx, 22].Value = data._参数1_销量总和;
                            sheet1.Cells[rowIdx, 23].Value = data._参数1_中位数;
                            sheet1.Cells[rowIdx, 24].Value = data._参数2_销量总和;
                            sheet1.Cells[rowIdx, 25].Value = data._参数2_中位数;
                            sheet1.Cells[rowIdx, 26].Value = data._参数3_销量总和;
                            sheet1.Cells[rowIdx, 27].Value = data._参数3_中位数;


                            sheet1.Cells[rowIdx, 28].Value = data._参数1详情;
                            sheet1.Cells[rowIdx, 29].Value = data._参数2详情;
                            sheet1.Cells[rowIdx, 30].Value = data._参数3详情;
                        }
                        #endregion

                        #region 全部边框
                        {
                            var endRow = sheet1.Dimension.End.Row;
                            var endColumn = sheet1.Dimension.End.Column;
                            using (var rng = sheet1.Cells[1, 1, endRow, endColumn])
                            {
                                rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            }
                        }
                        #endregion
                    }
                    #endregion

                    #region 异常sku
                    {
                        var sheet1 = workbox.Worksheets.Add("30天无销量SKU");
                        sheet1.Cells[1, 1].Value = "SKU";

                        #region 数据行
                        for (int idx = 0, len = errorSKUs.Count, rowIdx = 2; idx < len; idx++, rowIdx++)
                        {
                            sheet1.Cells[rowIdx, 1].Value = errorSKUs[idx];
                        }
                        #endregion
                    }
                    #endregion
                    buffer4 = package.GetAsByteArray();
                }
            }
            #endregion

            #region 导出
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
                   }
               }, null);
            #endregion
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
                    _SKU = value != null ? value.ToString().Trim().ToUpper() : "";
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

            public void _原先计算方式(List<decimal> list1, List<decimal> list2, List<decimal> list3, ref _计算过程 detail)
            {
                int count1 = list1.Count;
                int count2 = list2.Count;
                int count3 = list3.Count;
                detail._参数1_销量总和 = list1.Sum();
                detail._参数2_销量总和 = list2.Sum();
                detail._参数3_销量总和 = list3.Sum();

                detail._日平均销量 = Math.Round((detail._参数1_销量总和 / count1 + detail._参数2_销量总和 / count2 + detail._参数3_销量总和 / count3) / 3, 1);
                detail._库存上限 = _预警销售天数 * detail._日平均销量;
                detail._库存下限 = _采购到货天数 * detail._日平均销量;

                detail._建议采购 = Math.Round(detail._库存上限 + detail._库存下限 - _可用数量 - _采购未入库 + _缺货及未派单数量, 0);
            }

            public void _中位数计算方式(List<decimal> list1, List<decimal> list2, List<decimal> list3, ref _计算过程 detail)
            {
                detail._参数1_中位数 = Calcu中位数(list1);
                detail._参数2_中位数 = Calcu中位数(list2);
                detail._参数3_中位数 = Calcu中位数(list3);

                detail._日平均销量 = Math.Round((detail._参数1_中位数 + detail._参数2_中位数 + detail._参数3_中位数) / 3, 1);
                detail._库存上限 = _预警销售天数 * detail._日平均销量;
                detail._库存下限 = _采购到货天数 * detail._日平均销量;

                detail._建议采购 = Math.Round(detail._库存上限 + detail._库存下限 - _可用数量 - _采购未入库 + _缺货及未派单数量, 0);
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
                    _SKU = value != null ? value.ToString().Trim().ToUpper() : "";
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

            public decimal _日平均销量 { get; set; }
            public decimal _库存上限 { get; set; }
            public decimal _库存下限 { get; set; }
            public decimal _可用数量 { get; set; }
            public decimal _采购未入库 { get; set; }
            public decimal _缺货及未派单数量 { get; set; }
            public decimal _日平均销量_原先算法 { get; set; }
            public decimal _日平均销量_中位数 { get; set; }

            //_日平均销量
            //_库存上限 + _库存下限 - _可用数量 - _采购未入库 + _缺货及未派单数量;

        }

        class _计算详情
        {
            public string _SKU { get; set; }
            public string _采购员 { get; set; }
            public decimal _含税单价 { get; set; }
            public decimal _可用数量 { get; set; }
            public decimal _采购未入库 { get; set; }
            public decimal _缺货及未派单数量 { get; set; }
            public decimal _预计可用库存 { get; set; }
            public decimal _预警销售天数 { get; set; }
            public decimal _采购到货天数 { get; set; }
            public decimal _缺货占用数量 { get; set; }

            public decimal _库存上限_原先算法 { get; set; }
            public decimal _库存下限_原先算法 { get; set; }
            public decimal _日平均销量_原先算法 { get; set; }

            public decimal _库存上限_中位数 { get; set; }
            public decimal _库存下限_中位数 { get; set; }
            public decimal _日平均销量_中位数 { get; set; }

            public decimal _原表格导出的建议采购数量_原预警表 { get; set; }
            public decimal _原表格导出的建议采购数量_中位数表 { get; set; }
            public decimal _计算后的建议采购数量_原先算法 { get; set; }
            public decimal _计算后的建议采购数量_中位数 { get; set; }
            public decimal Qty { get; set; }

            public decimal _参数1_销量总和 { get; set; }
            public decimal _参数1_中位数 { get; set; }
            public decimal _参数2_销量总和 { get; set; }
            public decimal _参数2_中位数 { get; set; }
            public decimal _参数3_销量总和 { get; set; }
            public decimal _参数3_中位数 { get; set; }


            public string _参数1详情 { get; set; }
            public string _参数2详情 { get; set; }
            public string _参数3详情 { get; set; }
        }

        class _计算过程
        {
            public decimal _库存上限 { get; set; }
            public decimal _库存下限 { get; set; }
            public decimal _日平均销量 { get; set; }
            public decimal _建议采购 { get; set; }

            public decimal _参数1_销量总和 { get; set; }
            public decimal _参数1_中位数 { get; set; }
            public decimal _参数2_销量总和 { get; set; }
            public decimal _参数2_中位数 { get; set; }
            public decimal _参数3_销量总和 { get; set; }
            public decimal _参数3_中位数 { get; set; }
        }

        class _异常SKU信息
        {
            public string SKU { get; set; }
            public int _来源 { get; set; }
        }
    }
}
