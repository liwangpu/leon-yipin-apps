using LinqToExcel.Attributes;

namespace OrderAllot.Maps
{
    public class Warning
    {
        private int _SuggestAmount;

        [ExcelColumn("供应商")]
        public string _供应商 { get; set; }
        [ExcelColumn("SKU码")]
        public string _SKU { get; set; }
        [ExcelColumn("建议采购数量")]
        public int _建议采购数量
        {
            get
            {
                return CalAmount(_SuggestAmount);
            }
            set
            {
                _SuggestAmount = value;
            }
        }
        [ExcelColumn("仓库")]
        public string _仓库 { get; set; }
        [ExcelColumn("采购员")]
        public string _采购员 { get; set; }
        [ExcelColumn("商品成本单价")]
        public double _商品成本单价 { get; set; }
        public double _采购金额
        {
            get
            {
                return _商品成本单价 * _建议采购数量;
            }
        }

        #region CalAmount 计算建议采购数量
        private int CalAmount(int orgAmount)
        {
            var calAmount = 5;
            //小于5 ==>1

            if (orgAmount > 5 && orgAmount < 10)
            {
                calAmount = 10;
            }

            if (orgAmount > 10)
            {
                var bei = 0;
                var remain = orgAmount % 10;
                if (remain >= 5)
                {
                    bei = 1;
                }
                bei += (orgAmount - remain) / 10;
                calAmount = bei * 10;

            }
            return calAmount;
        }
        #endregion
    }
}
