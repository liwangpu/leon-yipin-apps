﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CommonLibs
{
    public class MathHelper
    {

        /// <summary>
        /// 拉以达法则剔除离值计算总和
        /// </summary>
        /// <param name="dataList"></param>
        /// <param name="kickList"></param>
        /// <param name="outlierRatio"></param>
        /// <param name="expectation"></param>
        /// <returns></returns>
        public static decimal SumKickOutlier(List<decimal> dataList, out List<decimal> kickList, out decimal stdev
            , out decimal lowLimit, out decimal upLimit, OutlierRatio outlierRatio, decimal expectation = 0.01m)
        {
            //标准差
            stdev = CalculateStdDev(dataList);
            var rlowLimit = (expectation - stdev * (int)outlierRatio);
            var rupLimit = (expectation + stdev * (int)outlierRatio);
            kickList = dataList.Where(x => x < rlowLimit || x > rupLimit).Select(x => x).ToList();
            lowLimit = rlowLimit;
            upLimit = rupLimit;
            return dataList.Where(x => x >= rlowLimit && x <= rupLimit).Select(x => x).Sum();
        }

        /// <summary>
        /// 计算标准值
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        private static decimal CalculateStdDev(List<decimal> values)
        {
            decimal ret = 0;
            if (values.Count() > 0)
            {
                //  计算平均数   
                var avg = values.Average();
                //  计算各数值与平均数的差值的平方，然后求和 
                var sum = values.Sum(d => Math.Pow(Convert.ToDouble(d - avg), 2));
                //  除以数量，然后开方
                ret = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(sum) / values.Count()));
            }
            return ret;
        }


    }

    /// <summary>
    /// 拉以达法则倍率
    /// </summary>
    public enum OutlierRatio
    {
        Single = 1,
        Twice = 2,
        Triple = 3
    }
}
