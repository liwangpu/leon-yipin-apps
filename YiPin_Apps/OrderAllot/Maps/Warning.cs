﻿using LinqToExcel.Attributes;
using OrderAllot.Libs;
using System;

namespace OrderAllot.Maps
{
    public class Warning
    {
        private string self供应商;
        private string selfSKU;


        ////临时处理
        //[ExcelColumn("采购未入库")]
        //private string self采购未入库 { get; set; }

        [ExcelColumn("供应商")]
        public string _供应商
        {
            get
            {
                return self供应商;
            }
            set
            {
                self供应商 = string.IsNullOrEmpty(value) == true ? "" : value.ToString().Trim();
            }
        }
        [ExcelColumn("SKU码")]
        public string _SKU
        {
            get
            {
                return selfSKU;
            }
            set
            {
                selfSKU = string.IsNullOrEmpty(value) == true ? "" : value.ToString().Trim();
            }
        }
        [ExcelColumn("仓库")]
        public string _仓库 { get; set; }
        [ExcelColumn("采购员")]
        public string _采购员 { get; set; }


        [ExcelColumn("库存上限")]
        public double _库存上限 { get; set; }
        [ExcelColumn("库存下限")]
        public double _库存下限 { get; set; }
        [ExcelColumn("可用数量")]
        public double _可用数量 { get; set; }

        public double _采购未入库 { get; set; }

        [ExcelColumn("采购未入库")]
        public double? _Good { get; set; }

        //{
        //    get
        //    {
        //        double tmp = 0;
        //        try
        //        {
        //            tmp = Convert.ToDouble(self采购未入库);
        //        }
        //        catch (Exception) { }
        //        return tmp;
        //    }
        //}
        //[ExcelColumn("缺货及未派单数量")]
        public double _缺货及未派单数量 { get; set; }

        //[ExcelColumn("商品成本单价")]
        public double _商品成本单价 { get; set; }

        public double _采购金额
        {
            get
            {
                return _商品成本单价 * _最终需要采购数量;
            }
        }

        /// <summary>
        /// 这个没有经转换,是最原始的建议采购数量,用来对比库存信息是否够卖信息
        /// </summary>
        public double _建议采购数量
        {
            get
            {
                return _库存上限 + _库存下限 - _可用数量 - _采购未入库 + _缺货及未派单数量;
            }
        }

        /// <summary>
        /// 经过转换的,是最终需要采购的数量
        /// </summary>
        public double _最终需要采购数量
        {
            get
            {
                return Helper.CalAmount(_建议采购数量);
            }
        }
    }
}
