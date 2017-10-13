using LinqToExcel.Attributes;
using System.Text;

namespace OrderAllot.Maps
{
    public class _Form5延时报表
    {
        [ExcelColumn("SKU码")]
        public string _SKU { get; set; }
        [ExcelColumn("可用数量")]
        public double _可用数量 { get; set; }
        [ExcelColumn("缺货及未派单数量")]
        public double _缺货及未派单数量 { get; set; }

    }
}
