using System.Collections.Generic;
namespace OrderAllot.Entities
{
    public class IM订单分配 : IEntity
    {
        private double _SuggestAmount;

        public string _供应商 { get; set; }
        public string _SKU { get; set; }
        public double _建议采购数量
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
        public string _仓库 { get; set; }
        public double _库存上限 { get; set; }
        public double _库存下限 { get; set; }
        public double _可用数量 { get; set; }
        public double _采购未入库 { get; set; }
        public double _缺货及未派单数量 { get; set; }
        public string _采购员 { get; set; }
        public double _商品成本单价 { get; set; }
        public double _采购金额
        {
            get
            {
                return _商品成本单价 * _建议采购数量;
            }
        }

        public double _特殊查看是否够卖
        {
            get
            {
                return _库存上限 + _库存下限 - _可用数量 - _采购未入库 + _缺货及未派单数量;
            }
        }

        public double _特殊最终库存多余数量 { get; set; }

        #region CalAmount 计算建议采购数量
        private double CalAmount(double orgAmount)
        {
            var calAmount = 5.0;
            //小于5 ==>1

            if (orgAmount > 5 && orgAmount < 10)
            {
                calAmount = 10;
            }

            if (orgAmount > 10)
            {
                var bei = 0.0;
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

        public Dictionary<string, object> ToDictionary()
        {
            throw new System.NotImplementedException();
        }
    }
}
