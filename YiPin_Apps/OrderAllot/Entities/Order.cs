using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OrderAllot.Entities
{
    public class Order
    {
        public string _供应商 { get; set; }
        public string _SKU { get; set; }
        public int _Qty { get; set; }
        public string _仓库 { get; set; }
        public string _备注 { get; set; }
        public string _合同号 { get; set; }
        public string _采购员 { get; set; }
        public double _含税单价 { get; set; }
        public double _物流费 { get; set; }
        public string _付款方式 { get; set; }
        public string _制单人 { get; set; }
        public string _到货日期 { get; set; }
        public string _1688单号 { get; set; }
        public double _预付款 { get; set; }
        public double _对应供应商采购金额 { get; set; }

    }
}
