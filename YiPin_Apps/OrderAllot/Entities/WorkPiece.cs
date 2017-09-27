using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OrderAllot.Entities
{
    public class WorkPiece
    {
        public string _采购员 { get; set; }
        public double _完成金额 { get; set; }
        public double _未完成金额 { get; set; }
        public double _总计金额
        {
            get
            {
                return _完成金额 + _未完成金额;
            }
        }
        public int _完成单量 { get; set; }
        public int _未完成单量 { get; set; }
        public int _总计
        {
            get
            {
                return _完成单量 + _未完成单量;
            }
        }
        public double _平均每单金额
        {
            get
            {
                double tmp = _总计 != 0 ? _总计金额 / _总计 : 0;
                return Math.Round(tmp, 2);
            }
        }
        public double _完成率
        {
            get
            {
                var tmp = Convert.ToDouble(_完成单量) / Convert.ToDouble(_总计);
                return Math.Round(tmp, 2);
            }
        }
    }
}
