using CommonLibs;
using LinqToExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Gadget.Libs
{
    public class FormHelper
    {
        /**************** static method ****************/

        #region static GetCSVPath 弹出对话框获取CSV文件路径到TextBox
        /// <summary>
        /// 弹出对话框获取CSV文件路径到TextBox
        /// </summary>
        /// <param name="txtbox"></param>
        public static void GetCSVPath(TextBox txtbox)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "CSV 文件|*.csv";//设置文件类型
            OpenFileDialog1.Title = "CSV 文件";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                if (Helper.CheckCSVFileName(OpenFileDialog1.FileName))
                {
                    txtbox.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region static GenerateTableDes 导出txt表格说明
        /// <summary>
        /// 导出txt表格说明
        /// </summary>
        /// <param name="types"></param>
        public static void GenerateTableDes(params Type[] types)
        {
            var strDesc = XlsxHelper.GetDecsipt(types);
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "记事本|*.txt";//设置文件类型
            saveFile.Title = "导出说明文件";//设置标题
            saveFile.AddExtension = true;//是否自动增加所辍名
            saveFile.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (saveFile.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                File.WriteAllText(saveFile.FileName, strDesc);
            }
        }
        #endregion

        #region static ReadCSVFile 读取CSV文件
        /// <summary>
        /// 读取CSV文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="strCSVPath"></param>
        /// <param name="list"></param>
        /// <param name="strError"></param>
        public static void ReadCSVFile<T>(string strCSVPath, ref List<T> list, ref string strError)
            where T : class,new()
        {
            strError = string.Empty;
            if (!string.IsNullOrEmpty(strCSVPath))
            {
                using (var csv = new ExcelQueryFactory(strCSVPath))
                {
                    try
                    {
                        var tmp = from c in csv.Worksheet<T>()
                                  select c;
                        list.AddRange(tmp);
                    }
                    catch (Exception ex)
                    {
                        strError = ex.Message;
                    }
                }
            }
        }
        #endregion
    }
}
