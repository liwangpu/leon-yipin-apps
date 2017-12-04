using LinqToExcel.Attributes;
using System.Text;
using System;
using CommonLibs;
namespace OrderAllot.Maps
{
    [ExcelTable("上海/昆山仓库")]
    public class _Form5延时报表
    {
        [ExcelColumn("可用数量")]
        public string str可用数量 { get; set; }
        [ExcelColumn("缺货及未派单数量")]
        public string str缺货及未派单数量 { get; set; }

        private string __sku;
        [ExcelColumn("SKU码")]
        public string _SKU
        {
            get
            {
                return __sku;
            }
            set
            {
                __sku = !string.IsNullOrEmpty(value) ? value.Trim() : "";
            }
        }
        public double _可用数量
        {
            get
            {
                double res = 0;
                if (!string.IsNullOrEmpty(str可用数量))
                    res = Convert.ToDouble(str可用数量);
                return res;
            }
        }
        public double _缺货及未派单数量
        {
            get
            {
                double res = 0;
                if (!string.IsNullOrEmpty(str缺货及未派单数量))
                    res = Convert.ToDouble(str缺货及未派单数量);
                return res;
            }
        }

    }

    [ExcelTable("延时报表")]
    public class _Form5缺货延时报表判断
    {
        private string _sku;
        [ExcelColumn("sku")]
        public string sku
        {
            get
            {
                return _sku;
            }
            set
            {
                _sku = !string.IsNullOrEmpty(value) ? value.Trim() : "";
            }
        }
        [ExcelColumn("是否停售")]
        public string _是否停售 { get; set; }
        [ExcelColumn("停售时间")]
        public string _停售时间 { get; set; }
        [ExcelColumn("店铺SKU")]
        public string _店铺SKU { get; set; }
        [ExcelColumn("订单编号")]
        public string _订单编号 { get; set; }
        [ExcelColumn("卖家简称")]
        public string _卖家简称 { get; set; }
        [ExcelColumn("Itemid")]
        public string _Itemid { get; set; }
        [ExcelColumn("延时时间")]
        public string _延时时间 { get; set; }
        [ExcelColumn("交易时间")]
        public string _交易时间 { get; set; }
        [ExcelColumn("业绩归属1")]
        public string _业绩归属1 { get; set; }
        [ExcelColumn("业绩归属2")]
        public string _业绩归属2 { get; set; }
        [ExcelColumn("采购员")]
        public string _采购员 { get; set; }
        [ExcelColumn("平台")]
        public string _平台 { get; set; }
        [ExcelColumn("本订单销售数量")]
        public string _本订单销售数量 { get; set; }
        [ExcelColumn("缺货总数")]
        public string _缺货总数 { get; set; }
        [ExcelColumn("发货仓库")]
        public string _发货仓库 { get; set; }
        [ExcelColumn("库存数量")]
        public string _库存数量 { get; set; }
        [ExcelColumn("占用数量")]
        public string _占用数量 { get; set; }
        [ExcelColumn("可用数量")]
        public string _可用数量 { get; set; }
    }
}
