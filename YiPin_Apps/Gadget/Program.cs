using System;
using System.Windows.Forms;

namespace Gadget
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new _库存盘点());//库存盘点
            //Application.Run(new _工资计算());//工资计算
            //Application.Run(new _商品信息统计());//商品信息统计
            //Application.Run(new _分库盘点());//分库盘点
            //Application.Run(new _移库());//移库
            //Application.Run(new _产品销量统计());//产品销量统计
            //Application.Run(new _排除侵权());//排除侵权
            //Application.Run(new _排除侵权_订单分配());//排除侵权_订单分配
            //Application.Run(new _采购订单配货());//采购订单配货
            //Application.Run(new _整合供应商人员工资统计());//整合供应商人员工资统计
            //Application.Run(new _点货绩效());//点货绩效
            //Application.Run(new _库存积压详情());//库存积压详情统计
            //Application.Run(new _采购订单配货新());//采购订单配货新
            //Application.Run(new _配货绩效());//配货绩效
            //Application.Run(new _乱单绩效());//乱单绩效
            //Application.Run(new _仓库加班考勤());//仓库加班考勤
            //Application.Run(new Form1());
            //Application.Run(new _测试());//
        }
    }
}
