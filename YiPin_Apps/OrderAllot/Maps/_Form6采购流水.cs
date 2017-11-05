using LinqToExcel.Attributes;
using System.Text;
using System;
using OrderAllot.Libs;

namespace OrderAllot.Maps
{
    public class _Form6采购流水
    {
        [ExcelColumn("采购sku数量")]
        public string _OrgSKU个数 { get; set; }
        [ExcelColumn("总金额")]
        public string _Org总金额 { get; set; }

        [ExcelColumn("采购员")]
        public string _采购员 { get; set; }
        [ExcelColumn("制单人")]
        public string _制单人 { get; set; }

    }

    public class _Form6采购流水Model
    {
        public string _采购员 { get; set; }
        public string _制单人 { get; set; }
        public int _SKU个数 { get; set; }
        public decimal _总金额 { get; set; }
    }

    public class _订单奖励
    {
        public string _采购员 { get; set; }

        public decimal _D1_J1 { get; set; }
        public decimal _D1_J2 { get; set; }
        public decimal _D1_J3 { get; set; }
        public decimal _D1_J4 { get; set; }
        public decimal _D1_J5 { get; set; }
        public decimal _D1_Sum
        {
            get
            {
                return _D1_J1 + _D1_J2 + _D1_J3 + _D1_J4 + _D1_J5;
            }
        }


        public decimal _D2_J1 { get; set; }
        public decimal _D2_J2 { get; set; }
        public decimal _D2_J3 { get; set; }
        public decimal _D2_J4 { get; set; }
        public decimal _D2_J5 { get; set; }
        public decimal _D2_Sum
        {
            get
            {
                return _D2_J1 + _D2_J2 + _D2_J3 + _D2_J4 + _D2_J5;
            }
        }

        public decimal _D3_J1 { get; set; }
        public decimal _D3_J2 { get; set; }
        public decimal _D3_J3 { get; set; }
        public decimal _D3_J4 { get; set; }
        public decimal _D3_J5 { get; set; }
        public decimal _D3_Sum
        {
            get
            {
                return _D3_J1 + _D3_J2 + _D3_J3 + _D3_J4 + _D3_J5;
            }
        }

        public decimal _D4_J1 { get; set; }
        public decimal _D4_J2 { get; set; }
        public decimal _D4_J3 { get; set; }
        public decimal _D4_J4 { get; set; }
        public decimal _D4_J5 { get; set; }
        public decimal _D4_Sum
        {
            get
            {
                return _D4_J1 + _D4_J2 + _D4_J3 + _D4_J4 + _D4_J5;
            }
        }

        public decimal _D5_J1 { get; set; }
        public decimal _D5_J2 { get; set; }
        public decimal _D5_J3 { get; set; }
        public decimal _D5_J4 { get; set; }
        public decimal _D5_J5 { get; set; }
        public decimal _D5_Sum
        {
            get
            {
                return _D5_J1 + _D5_J2 + _D5_J3 + _D5_J4 + _D5_J5;
            }
        }

        public decimal _产品归属奖励 { get; set; }

        public decimal _合计
        {
            get
            {
                return _D1_Sum + _D2_Sum + _D3_Sum + _D4_Sum + _D5_Sum + _产品归属奖励;
            }
        }

        public bool IsBuyer 
        {
            get 
            {
                return Helper.IsBuyer(this._采购员);
            }
        }
    }

}
