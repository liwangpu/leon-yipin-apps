using LinqToExcel.Attributes;
using System.Text;

namespace OrderAllot.Maps
{
    public class OrderState
    {
        [ExcelColumn("采购员")]
        public string _采购员 { get; set; }
        [ExcelColumn("内部标签")]
        public string _内部标签 { get; set; }
        [ExcelColumn("总数量")]
        public double _总数量 { get; set; }
        [ExcelColumn("总金额")]
        public double _总金额 { get; set; }
        [ExcelColumn("采购sku数量")]
        public int _采购sku数量 { get; set; }

        public double _完成金额 { get; set; }
        public bool _是否完成
        {
            get
            {
                var bFlag = false;
                if (!string.IsNullOrEmpty(_内部标签))
                {
                    var is合付 = _内部标签.IndexOf("合付") != -1;
                    var is可付 = _内部标签.IndexOf("可付") != -1;
                    bFlag = is合付 || is可付;
                }
                return bFlag;
            }
        }
    }
}
