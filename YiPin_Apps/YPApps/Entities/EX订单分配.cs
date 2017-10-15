using System.Collections.Generic;
using System;
namespace YPApps.Entities
{
    public class EX订单分配 : IExport
    {
        public string _供应商 { get; set; }
        public string _SKU { get; set; }
        public double _Qty { get; set; }
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

        public Dictionary<string, object> ToDictionary()
        {
            var dicData = new Dictionary<string, object>();
            dicData["供应商"] = this._供应商;
            dicData["SKU"] = this._SKU;
            dicData["Qty"] = this._Qty;
            dicData["仓库"] = this._仓库;
            dicData["备注"] = this._备注;
            dicData["合同号"] = this._合同号;
            dicData["采购员"] = this._采购员;
            dicData["含税单价"] = this._含税单价;
            dicData["物流费"] = this._物流费;
            dicData["付款方式"] = this._付款方式;
            dicData["制单人"] = this._制单人;
            dicData["到货日期"] = this._到货日期;
            dicData["1688单号"] = this._1688单号;
            dicData["预付款"] = this._预付款;
            dicData["对应供应商采购金额"] = this._对应供应商采购金额;
            return dicData;
        }

        Dictionary<string, object> IExport.ToDictionary()
        {
            throw new NotImplementedException();
        }
    }
}
