using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace EpplusHelper
{
    public class SheetReader<T>
        where T : class, new()
    {
        public static List<T> From(ExcelWorksheet sheet, int headerRow = 1, int dataRow = 2)
        {
            var mappingType = typeof(T);

            var endColumn = sheet.Dimension.End.Column;
            var endRow = sheet.Dimension.End.Row;

            var mapping = new Dictionary<string, Tuple<string, int>>();
            var list = new List<T>();

            #region 根据标注,获取表格匹配信息
            {
                var exAttrType = typeof(ExcelColumnAttribute);
                var mappingTypeProps = mappingType.GetProperties();

                foreach (var prop in mappingTypeProps)
                {



                    var attrs = prop.GetCustomAttributes(exAttrType, true);
                    if (attrs.Count() > 0)
                    {
                        var attr = attrs[0] as ExcelColumnAttribute;

                        var distType = "string";
                        var ptName = prop.PropertyType.Name.ToLower();
                        if (ptName.Contains("int"))
                            distType = "int";
                        else if (ptName.Contains("decimal"))
                            distType = "decimal";
                        else if (ptName.Contains("double"))
                            distType = "double";
                        else if (ptName.Contains("datetime"))
                            distType = "datetime";
                        else { }


                        var distColumn = attr.Column;
                        //验证一下列数对不对,不对需要遍历纠正
                        if (sheet.Cells[headerRow, distColumn].Value == null || sheet.Cells[headerRow, distColumn].Value.ToString().Trim() != attr.Tile)
                        {
                            for (int i = 1; i <= endColumn; i++)
                            {
                                if (sheet.Cells[1, i].Value.ToString().Trim() == attr.Tile)
                                {
                                    distColumn = i;
                                    break;
                                }
                            }
                        }



                        if (!mapping.ContainsKey(prop.Name))
                            mapping[prop.Name] = new Tuple<string, int>(distType, distColumn);

                    }
                }
                #endregion

                for (int idx = endRow; idx >= dataRow; idx--)
                {
                    if (idx == 45)
                    {

                    }
                    var instance = new T();
                    foreach (var item in mapping)
                    {
                        var cell = sheet.Cells[idx, item.Value.Item2];
                        if (cell.Value == null) continue;

                        //GetValue说失败是忽略的,但是测试下来发现并不是
                        if (item.Value.Item1 == "int")
                        {
                            try
                            {
                                mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { cell.GetValue<int>() });
                            }
                            catch
                            {
                                var str = cell.Value.ToString().Trim();
                                if (!string.IsNullOrEmpty(str))
                                {
                                    int val = 0;
                                    var b = int.TryParse(str, out val);
                                    if (b)
                                    {
                                        mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { val });
                                    }
                                }
                            }
                        }
                        else if (item.Value.Item1 == "decimal")
                        {
                            try
                            {
                                mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { cell.GetValue<decimal>() });
                            }
                            catch
                            {
                                var str = cell.Value.ToString().Trim();
                                if (!string.IsNullOrEmpty(str))
                                {
                                    decimal val = 0;
                                    var b = decimal.TryParse(str, out val);
                                    if (b)
                                    {
                                        mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { val });
                                    }
                                }
                            }
                        }
                        else if (item.Value.Item1 == "double")
                        {
                            try
                            {
                                mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { cell.GetValue<double>() });
                            }
                            catch
                            {
                                var str = cell.Value.ToString().Trim();
                                if (!string.IsNullOrEmpty(str))
                                {
                                    double val = 0;
                                    var b = double.TryParse(str, out val);
                                    if (b)
                                    {
                                        mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { val });
                                    }
                                }
                            }
                        }
                        else if (item.Value.Item1 == "datetime")
                        {
                            try
                            {
                                mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { cell.GetValue<DateTime>() });
                            }
                            catch
                            {
                                var str = cell.Value.ToString().Trim();
                                if (!string.IsNullOrEmpty(str))
                                {
                                    DateTime val;
                                    var b = DateTime.TryParse(str, out val);
                                    if (b)
                                    {
                                        mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { val });
                                    }
                                }
                            }
                        }
                        else
                        {
                            mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { cell.GetValue<string>() });
                        }

                    }

                    list.Add(instance);
                }
            }

            return list;
        }
    }
}
