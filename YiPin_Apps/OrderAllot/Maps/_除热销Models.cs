using CommonLibs;
using LinqToExcel.Attributes;
using OrderAllot.Libs;
using System;
namespace OrderAllot.Maps
{
    public class _除热销_Warning
    {
        private string self供应商;
        private string selfSKU;

        public double _日销量 { get; set; }



        //临时处理
        [ExcelColumn("采购未入库")]
        public string org采购未入库 { get; set; }
        [ExcelColumn("缺货及未派单数量")]
        public string org缺货及未派单数量 { get; set; }

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

        public double _tmp库存上限;
        public double _tmp库存下限;

        [ExcelColumn("库存上限")]
        public double _库存上限
        {
            get
            {
                double res = 0;
                if (IsHot)
                {
                    res = _预警销售天数 * _日销量;
                }
                else
                {
                    res = _tmp库存上限;
                }
                return res;
            }
            set
            {
                _tmp库存上限 = value;
            }
        }

        [ExcelColumn("库存下限")]
        public double _库存下限
        {
            get
            {
                double res = 0;
                if (IsHot)
                {
                    res = _采购到货天数 * _日销量;
                }
                else
                {
                    res = _tmp库存下限;
                }

                return res;
            }

            set
            {
                _tmp库存下限 = value;
            }
        }





        [ExcelColumn("可用数量")]
        public double _可用数量 { get; set; }
        //[ExcelColumn("采购未入库")]
        public double _采购未入库
        {
            get
            {
                var tmp = 0.0;
                if (!string.IsNullOrEmpty(org采购未入库))
                {
                    tmp = Convert.ToDouble(org采购未入库);
                }
                return tmp;
            }
        }
        //[ExcelColumn("缺货及未派单数量")]
        public double _缺货及未派单数量
        {
            get
            {
                var tmp = 0.0;
                if (!string.IsNullOrEmpty(org缺货及未派单数量))
                {
                    tmp = Convert.ToDouble(org缺货及未派单数量);
                }
                return tmp;
            }
        }

        [ExcelColumn("商品成本单价")]
        public double _商品成本单价 { get; set; }

        [ExcelColumn("预警销售天数")]
        public double _预警销售天数 { get; set; }
        [ExcelColumn("采购到货天数")]
        public double _采购到货天数 { get; set; }
        [ExcelColumn("30天销量")]
        public double _30天销量 { get; set; }
        [ExcelColumn("15天销量")]
        public double _15天销量 { get; set; }
        [ExcelColumn("5天销量")]
        public double _5天销量 { get; set; }



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

        public bool IsHot { get; set; }
    }

    public class _热销产品
    {
        [ExcelColumn("SKU")]
        public string _SKU { get; set; }

        [ExcelColumn("销量")]
        public double _销量 { get; set; }
    }
}
