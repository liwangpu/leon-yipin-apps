using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CommonLibs
{
    public class Helper
    {
        /**************** static method ****************/

        #region GetBuyers 获取采购员
        /// <summary>
        /// 获取采购员
        /// </summary>
        /// <returns></returns>
        public static List<string> GetBuyers()
        {
            var buyers = new List<string>();
            buyers.AddRange(GetHeFeiBuyer());
            buyers.AddRange(GetShangHaiBuyer());
            return buyers;
        }
        #endregion

        #region GetHeFeiBuyer 获取所有合肥采购
        /// <summary>
        /// 获取所有合肥采购
        /// </summary>
        /// <returns></returns>
        private static List<string> GetHeFeiBuyer()
        {
            var buyers = new List<string>();
            buyers.Add("崔侠梅");
            buyers.Add("董文丽");
            buyers.Add("李震");
            buyers.Add("侯春喜");
            buyers.Add("苏苗雨");
            buyers.Add("唐汉成");
            buyers.Add("江峰");
            buyers.Add("袁国梁");
            buyers.Add("李倩倩");
            buyers.Add("韩丽敏");
            buyers.Add("王庆飞");
            buyers.Add("王微");
            buyers.Add("刘龙敏");
            buyers.Add("蔡东玲");
            buyers.Add("张秀秀");
            buyers.Add("王腾飞");
            buyers.Add("查文亮");
            buyers.Add("张露玲");
            buyers.Add("廖延鑫");

            return buyers;
        }
        #endregion

        #region GetShangHaiBuyer 获取所有上海采购
        /// <summary>
        /// 获取所有上海采购
        /// </summary>
        /// <returns></returns>
        private static List<string> GetShangHaiBuyer()
        {
            var buyers = new List<string>();
            buyers.Add("鲍祝平");
            buyers.Add("毕玉");
            buyers.Add("王思雅");
            buyers.Add("桂娅利");
            buyers.Add("黄妍妍");
            buyers.Add("章玲玲");
            buyers.Add("邵俊丽");
            buyers.Add("潘明媛");
            buyers.Add("蔡怡雯");
            buyers.Add("秦荧");
            buyers.Add("邹晓玲");
            buyers.Add("邵羽");
            buyers.Add("何萧雪");
            buyers.Add("徐周红");
            buyers.Add("陈春梦");
            buyers.Add("张莉萍");

            return buyers;
        }
        #endregion

        #region IsBuyer 判断是否采购
        /// <summary>
        /// 判断是否采购
        /// </summary>
        /// <param name="strName"></param>
        /// <returns></returns>
        public static bool IsBuyer(string strName)
        {
            var bFlag = false;
            var buyerList = GetBuyers();
            bFlag = buyerList.Where(x => x == strName).Count() > 0;
            return bFlag;
        }
        #endregion

        #region RemoveUnBuyers 去除非采购人员
        /// <summary>
        /// 去除非采购人员
        /// </summary>
        /// <param name="buyers"></param>
        /// <returns></returns>
        public static List<string> RemoveUnBuyers(List<string> buyers)
        {
            if (buyers != null)
            {
                for (int idx = buyers.Count - 1; idx >= 0; idx--)
                {
                    if (!IsBuyer(buyers[idx]))
                        buyers.RemoveAt(idx);
                }
                return buyers;
            }
            return new List<string>();
        }
        #endregion

        #region RemoveUnBuyersByList 根据输入的采购排除人员信息
        /// <summary>
        /// 根据输入的采购排除人员信息
        /// </summary>
        /// <param name="list"></param>
        /// <param name="buyers"></param>
        /// <returns></returns>
        public static List<string> RemoveUnBuyersByList(List<string> list, List<string> buyers)
        {
            if (buyers != null && buyers.Count > 0)
            {
                for (int idx = list.Count - 1; idx >= 0; idx--)
                {
                    var curName = list[idx];
                    var bFlag = buyers.Where(x => x == curName).Count() > 0;
                    if (!bFlag)
                        list.RemoveAt(idx);
                }
                return list;
            }
            return new List<string>();
        }
        #endregion

        #region IsSpecBuyerType 判断是否为对应采购类型
        /// <summary>
        /// 判断是否为对应采购类型
        /// </summary>
        /// <param name="strBuyerName"></param>
        /// <param name="spec"></param>
        /// <returns></returns>
        public static bool IsSpecBuyerType(string strBuyerName, BuyerTypeEnum spec)
        {
            if (spec == BuyerTypeEnum.ShangHai)
            {
                return GetShangHaiBuyer().Count(x => x == strBuyerName) > 0;
            }

            if (spec == BuyerTypeEnum.HeFei)
            {
                return GetHeFeiBuyer().Count(x => x == strBuyerName) > 0;
            }

            return false;
        }
        #endregion

        #region ChangeLowerBuyer 采购员转换
        /// <summary>
        /// 采购员转换
        /// </summary>
        /// <param name="orgBuyerName"></param>
        /// <returns></returns>
        public static string ChangeLowerBuyer(string orgBuyerName)
        {
            var newBuyerName = orgBuyerName;
            switch (orgBuyerName)
            {
                case "毕玉":
                    newBuyerName = "韩丽敏";
                    break;
                case "鲍祝平":
                    newBuyerName = "王素素";
                    break;
                case "黄妍妍":
                    newBuyerName = "吴倩";
                    break;
                case "潘明媛":
                    newBuyerName = "侯春喜";
                    break;
                case "章玲玲":
                    newBuyerName = "董文丽";
                    break;
                case "蔡怡雯":
                    newBuyerName = "崔侠梅";
                    break;
                case "邹晓玲":
                    newBuyerName = "苏苗雨";
                    break;
                case "王思雅":
                    newBuyerName = "袁国梁";
                    break;
                case "王梦梦":
                    newBuyerName = "吴海燕";
                    break;
                default:
                    break;
            }
            return newBuyerName;
        }
        #endregion

        #region CalAmount 计算建议采购数量
        public static double CalAmount(double orgAmount)
        {
            if (orgAmount <= 0)
            {
                return orgAmount;
            }


            var calAmount = 5.0;
            //小于5 ==>1
            if (orgAmount > 5 && orgAmount < 10)
            {
                calAmount = 10;
            }

            if (orgAmount > 10)
            {
                var bei = 0.0;
                var remain = orgAmount % 10;
                if (remain >= 5)
                {
                    bei = 1;
                }
                bei += (orgAmount - remain) / 10;
                calAmount = bei * 10;

            }
            return calAmount;
        }

        public static decimal CalAmount(decimal orgAmount)
        {
            double b = Convert.ToDouble(orgAmount);
            return Convert.ToDecimal(CalAmount(b));
        }
        #endregion

        #region CheckCSVFileName 校验csv文件名是否符合规范
        /// <summary>
        /// 校验csv文件名是否符合规范
        /// </summary>
        /// <param name="strFileName"></param>
        /// <returns></returns>
        public static bool CheckCSVFileName(string strFileName)
        {
            var strPureFileName = Path.GetFileNameWithoutExtension(strFileName);
            if (strPureFileName.Contains('.'))
            {
                return false;
            }
            return true;
        }
        #endregion

    }

    /// <summary>
    /// 采购类型
    /// </summary>
    public enum BuyerTypeEnum
    {
        ShangHai = 0,
        HeFei = 1
    }
}
