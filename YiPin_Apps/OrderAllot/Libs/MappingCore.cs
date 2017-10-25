using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using LinqToExcel;

namespace OrderAllot.Libs
{
    public class MappingCore<T>
        where T : class,new()
    {
        private string _ExcelPath;
        public List<T> Datas;

        #region 构造函数
        public MappingCore(string strExcelPath)
        {
            _ExcelPath = strExcelPath;
            Datas = new List<T>();
        }
        #endregion

        #region Parse 解析数据
        public List<T> Parse()
        {
            if (!string.IsNullOrEmpty(_ExcelPath))
            {
                using (var excel = new ExcelQueryFactory(_ExcelPath))
                {
                    var sheetNames = excel.GetWorksheetNames().ToList();
                    sheetNames.ForEach(s =>
                    {
                        var tmp = from c in excel.Worksheet<T>(s)
                                  select c;
                        Datas.AddRange(tmp);
                    });
                }
            }
            return Datas;
        }
        #endregion
    }
}
