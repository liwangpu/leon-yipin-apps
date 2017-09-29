using LinqToExcel.Attributes;
using System.Text;

namespace OrderAllot.Maps
{
    public class BuyersProvider
    {
        [ExcelColumn("供应商")]
        public string _供应商 { get; set; }
        [ExcelColumn("SKU数量")]
        public int _SKU数量 { get; set; }
        [ExcelColumn("采购")]
        public string _采购 { get; set; }

        public int _有几个采购 { get; set; }
    }
}
