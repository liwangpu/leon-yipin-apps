using System.Collections.Generic;

namespace YPApps.Entities
{
    public class EX工作量 : IExport
    {
        public string _采购员 { get; set; }
        public int _订单量 { get; set; }


        public Dictionary<string, object> ToDictionary()
        {
            var mappingDic = new Dictionary<string, object>();
            mappingDic["采购员"] = this._采购员;
            mappingDic["订单量"] = this._订单量;
            return mappingDic;
        }
    }
}
