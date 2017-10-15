using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YPApps.Libs
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
            buyers.Add("鲍祝平");
            buyers.Add("毕玉");
            buyers.Add("侯春喜");
            buyers.Add("王思雅");
            buyers.Add("曹晨晨");
            buyers.Add("黄妍妍");
            buyers.Add("章玲玲");
            buyers.Add("邵俊丽");
            buyers.Add("崔侠梅");
            buyers.Add("蔡怡雯");
            buyers.Add("桂娅利");
            buyers.Add("潘明媛");
            buyers.Add("秦荧");
            buyers.Add("邹晓玲");
            buyers.Add("董文丽");
            buyers.Add("王梦梦");
            buyers.Add("何萧雪");
            buyers.Add("苏苗雨");
            buyers.Add("王素素");
            buyers.Add("李曼曼");
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
                    newBuyerName = "李曼曼";
                    break;
                case "鲍祝平":
                    newBuyerName = "王素素";
                    break;
                case "黄妍妍":
                    newBuyerName = "曹晨晨";
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
                    newBuyerName = "韦秋菊";
                    break;
                default:
                    break;
            }
            return newBuyerName;
        }
        #endregion

    }
}
