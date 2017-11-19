using CommonLibs;
using LinqToExcel.Attributes;
using OrderAllot.Libs;
using System;

namespace OrderAllot.Maps
{
    public class _SKU供应商信息
    {

        private string _sku;
        [ExcelColumn("SKU码")]
        public string _SKU码
        {
            get
            {
                return _sku;
            }
            set
            {
                _sku = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
            }
        }

        [ExcelColumn("业绩归属2")]
        public string _开发 { get; set; }

        [ExcelColumn("商品创建时间")]
        public DateTime _商品创建时间 { get; set; }

        private string _str网址1;
        [ExcelColumn("网址")]
        public string _网址1
        {
            get
            {
                return _str网址1;
            }
            set
            {
                _str网址1 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
            }
        }

        private string _str网址2;
        [ExcelColumn("网址2")]
        public string _网址2
        {
            get
            {
                return _str网址2;
            }
            set
            {
                _str网址2 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
            }
        }

        private string _str网址3;
        [ExcelColumn("网址3")]
        public string _网址3
        {
            get
            {
                return _str网址3;
            }
            set
            {
                _str网址3 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
            }
        }

        private string _str网址4;
        [ExcelColumn("网址4")]
        public string _网址4
        {
            get
            {
                return _str网址4;
            }
            set
            {
                _str网址4 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
            }
        }

        private string _str网址5;
        [ExcelColumn("网址5")]
        public string _网址5
        {
            get
            {
                return _str网址5;
            }
            set
            {
                _str网址5 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
            }
        }

        private string _str网址6;
        [ExcelColumn("网址6")]
        public string _网址6
        {
            get
            {
                return _str网址6;
            }
            set
            {
                _str网址6 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
            }
        }
    }

    public class SKU供应商数量信息
    {
        public string _开发 { get; set; }
        public int _SKU个数 { get; set; }
        public int _零个链接 { get; set; }
        public int _一个链接 { get; set; }
        public int _两个链接 { get; set; }
        public int _重复链接 { get; set; }
    }


    public class 重复链接信息
    {
        public string _开发 { get; set; }
        public string _SKU { get; set; }
        public string _网址1 { get; set; }
        public string _网址2 { get; set; }
        public string _网址3 { get; set; }
        public string _网址4 { get; set; }
        public string _网址5 { get; set; }
        public string _网址6 { get; set; }
        public DateTime _商品创建时间 { get; set; }
    }
}
