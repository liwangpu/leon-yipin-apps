using System.Collections.Generic;
namespace YPApps.Entities
{
    public class IM订单分配
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

        #region Dictionary
        public static Dictionary<string, string> GetMapping()
        {
            var mappingDic = new Dictionary<string, string>();
            mappingDic["供应商"] = "_供应商";
            mappingDic["SKU码"] = "_SKU";
            mappingDic["建议采购数量"] = "_建议采购数量";
            mappingDic["仓库"] = "_仓库";
            mappingDic["库存上限"] = "_库存上限";
            mappingDic["库存下限"] = "_库存下限";
            mappingDic["可用数量"] = "_可用数量";
            mappingDic["采购未入库"] = "_采购未入库";
            mappingDic["缺货及未派单数量"] = "_缺货及未派单数量";
            mappingDic["采购员"] = "_采购员";
            mappingDic["商品成本单价"] = "_商品成本单价";
         
            return mappingDic;
        } 
        #endregion
    }
}
