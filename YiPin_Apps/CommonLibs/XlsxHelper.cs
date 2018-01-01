using LinqToExcel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CommonLibs
{
    public class XlsxHelper
    {
        /**************** static method ****************/

        #region ExportDecsipt 获取表格列说明信息
        /// <summary>
        /// 获取表格列说明信息
        /// </summary>
        /// <param name="types"></param>
        /// <returns></returns>
        public static string GetDecsipt(params Type[] types)
        {
            var builder = new StringBuilder();
            foreach (Type typeItem in types)
            {
                var tableAttrs = typeItem.GetCustomAttributes(typeof(ExcelTableAttribute), false);
                if (tableAttrs.Length > 0)
                {
                    var defaultAttr = tableAttrs[0] as ExcelTableAttribute;
                    builder.AppendLine("【" + defaultAttr.TableName + "】");
                }


                var properties = typeItem.GetProperties();
                for (int i = 0, len = properties.Length; i < len; i++)
                {
                    var propertyItem = properties[i];
                    var attrs = propertyItem.GetCustomAttributes(typeof(ExcelColumnAttribute), false);
                    if (attrs.Length > 0)
                    {
                        var defaultAttr = attrs[0] as ExcelColumnAttribute;
                        builder.Append(defaultAttr.ColumnName + "  ");
                    }

                    if (i == len - 1)
                    {
                        builder.AppendLine("");
                        builder.AppendLine("");
                    }
                }
            }
            return builder.ToString();
        }
        #endregion

        //public static byte[] RtpExcel(params List<RtpBase>[] tables)
        //{
        //    var buffer = new byte[0];
        //    foreach (var table in tables)
        //    {
        //        var typ = table.GetType();
        //        //var list = table as List<typ>;


        //    }
        //    return buffer;
        //}
    }

    /// <summary>
    /// 扩展Excel表格名字
    /// </summary>
    public class ExcelTableAttribute : Attribute
    {
        private string _strTableName;

        public ExcelTableAttribute(string strTableName)
        {
            _strTableName = strTableName;
        }
        public string TableName { get { return _strTableName; } }
    }

    /// <summary>
    /// 导出表格特性
    /// </summary>
    public class RptTableAttribute : Attribute
    {
        private string _strTableName;

        public RptTableAttribute(string strTableName)
        {
            _strTableName = strTableName;
        }
        public string TableName { get { return _strTableName; } }
    }

    /// <summary>
    /// 导出表格列特性
    /// </summary>
    public class RtpColumnAttribute : Attribute
    {
        private string _strColumnName;

        public RtpColumnAttribute(string strColumnName)
        {
            _strColumnName = strColumnName;
        }
        public string ColumnName { get { return _strColumnName; } }
    }

}
