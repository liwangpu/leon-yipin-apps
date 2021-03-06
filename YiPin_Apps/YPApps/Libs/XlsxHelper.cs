﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace YPApps.Libs
{
    public class XlsxHelper
    {
        /**************** public static method ****************/

        #region Read 根据表格名称解析表格数据
        /// <summary>
        /// 根据表格名称解析表格数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="strMappings"></param>
        /// <param name="strSheetName"></param>
        /// <param name="strOpMessage"></param>
        /// <param name="iHeaderRowIdx"></param>
        /// <returns></returns>
        public static List<T> Read<T>(Stream stream, Dictionary<string, string> strMappings, string strSheetName, out string strOpMessage, int iHeaderRowIdx = 1)
            where T : class,new()
        {
            strOpMessage = string.Empty;
            var entities = new List<T>();
            var strHeaders = strMappings.ToList().Select(x => x.Key).ToList();
            var strProperties = strMappings.ToList().Select(x => x.Value).ToList();
            using (var pck = new ExcelPackage(stream))
            {
                var mappingDic = new Dictionary<string, int>();//表格标题对应表格列映射
                //指定解析具体表格
                if (!string.IsNullOrEmpty(strSheetName))
                {
                    var workBook = pck.Workbook;
                    var opSheet = workBook.Worksheets[strSheetName];
                    if (opSheet != null)
                    {
                        mappingDic = ParseMapping(opSheet, strHeaders, strProperties, iHeaderRowIdx);
                        AssignEntityValue(opSheet, entities, mappingDic, iHeaderRowIdx + 1);
                    }
                    else
                    {
                        strOpMessage = "没有找到对应表格名称的表格";
                    }
                }
            }

            return entities;
        }
        #endregion

        #region Read 解析表格所有数据
        /// <summary>
        /// 解析表格所有数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="strMappings"></param>
        /// <param name="strSheetName"></param>
        /// <param name="strOpMessage"></param>
        /// <param name="iHeaderRowIdx"></param>
        /// <returns></returns>
        public static List<T> Read<T>(Stream stream, Dictionary<string, string> strMappings, out string strOpMessage, int iHeaderRowIdx = 1)
            where T : class,new()
        {
            strOpMessage = string.Empty;
            var entities = new List<T>();
            var strHeaders = strMappings.ToList().Select(x => x.Key).ToList();
            var strProperties = strMappings.ToList().Select(x => x.Value).ToList();
            using (var pck = new ExcelPackage(stream))
            {
                var mappingDic = new Dictionary<string, int>();//表格标题对应表格列映射
                var workBook = pck.Workbook;
                var opSheets = workBook.Worksheets.ToList();
                opSheets.ForEach(curSheet =>
                {
                    mappingDic = ParseMapping(curSheet, strHeaders, strProperties, iHeaderRowIdx);
                    AssignEntityValue(curSheet, entities, mappingDic, iHeaderRowIdx + 1);
                });
            }
            return entities;
        }
        #endregion

        #region SimpleWrite 将数据简单写入表格
        /// <summary>
        /// 将数据简单写入表格
        /// </summary>
        /// <param name="dicDatas"></param>
        /// <param name="strTableName"></param>
        /// <param name="strOpMsg"></param>
        /// <param name="iSheetRowLimit"></param>
        /// <param name="iHeaderRowIdx"></param>
        /// <returns></returns>
        public static byte[] SimpleWrite(List<Dictionary<string, object>> dicDatas, out string strOpMsg, string strTableName = "", int iSheetRowLimit = 65535, int iHeaderRowIdx = 1)
        {
            strOpMsg = string.Empty;
            var buffer = new byte[0];
            using (var pck = new ExcelPackage())
            {
                var strSheetNameBase = strTableName != "" ? strTableName : "Sheet";
                var entityCount = dicDatas.Count;
                //不用分表
                if (entityCount <= iSheetRowLimit)
                {
                    var opSheet = pck.Workbook.Worksheets.Add(strSheetNameBase + "1");
                    SimpleWriteData(opSheet, dicDatas, iHeaderRowIdx);
                }
                //分表
                else
                {
                    var dif = entityCount % iSheetRowLimit;
                    var sheetSum = (entityCount - dif) / iSheetRowLimit;
                    if (dif > 0)
                        sheetSum++;
                    for (int idx = 0; idx < sheetSum; idx++)
                    {
                        var opSheet = pck.Workbook.Worksheets.Add(strSheetNameBase + (idx + 1).ToString());
                        var refEntities = dicDatas.Skip(idx * iSheetRowLimit).Take(iSheetRowLimit).ToList();
                        SimpleWriteData(opSheet, refEntities, iHeaderRowIdx);
                    }
                }
                buffer = pck.GetAsByteArray();
            }
            return buffer;
        }
        #endregion

        #region SaveWorkBook 根据路径文件信息保存工作簿
        /// <summary>
        /// 根据路径文件信息保存工作簿
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="strFileName"></param>
        /// <param name="strOpMsg"></param>
        public static void SaveWorkBook(byte[] buffer, string strFileName, out string strOpMsg)
        {
            strOpMsg = string.Empty;
            var len = buffer.Length;
            using (var fs = File.Create(strFileName, len))
            {
                fs.Write(buffer, 0, len);
            }
        }
        #endregion

        /**************** protected static method ****************/

        #region ParseMapping 解析实体属性对应表格标题的列标映射
        /// <summary>
        /// 解析实体属性对应表格标题的列标映射
        /// </summary>
        /// <param name="oSheet"></param>
        /// <param name="strHeaders"></param>
        /// <param name="strPropertis"></param>
        /// <param name="iTitleRowPosition"></param>
        /// <returns></returns>
        protected static Dictionary<string, int> ParseMapping(ExcelWorksheet oSheet, List<string> strHeaders, List<string> strPropertis, int iTitleRowPosition)
        {
            var mappingDic = new Dictionary<string, int>();
            var endColumnIdx = oSheet.Dimension.End.Column;
            for (int colIdx = 1/*EPPlus列从1开始*/; colIdx <= endColumnIdx; colIdx++)
            {
                var value = oSheet.Cells[iTitleRowPosition, colIdx].Value;
                if (value != null && !string.IsNullOrEmpty(value.ToString()))
                {
                    for (int idx = 0, len = strPropertis.Count; idx < len; idx++)
                    {
                        var curProperty = strPropertis[idx];
                        //为了防止传入的headers和propertis对不上
                        if (strHeaders.Count >= idx + 1)
                        {
                            var curHeader = strHeaders[idx];
                            if (curHeader == value.ToString().Trim())
                                mappingDic.Add(curProperty, colIdx);
                        }
                    }
                }
            }
            return mappingDic;
        }
        #endregion

        #region AssignEntityValue 根据映射获取表格数据
        /// <summary>
        /// 根据映射获取表格数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="oSheet"></param>
        /// <param name="entities"></param>
        /// <param name="mappingDic"></param>
        /// <param name="iDataRowIdx"></param>
        protected static void AssignEntityValue<T>(ExcelWorksheet oSheet, List<T> entities, Dictionary<string, int> mappingDic, int iDataRowIdx)
            where T : class,new()
        {
            var entityType = typeof(T);
            var entityProperties = entityType.GetProperties();
            var endRowIdx = oSheet.Dimension.End.Row;
            for (int rowIdx = iDataRowIdx; rowIdx <= endRowIdx; rowIdx++)
            {
                var entity = new T();
                foreach (var item in entityProperties)
                {
                    var curPropertyName = item.Name;
                    var refMapping = mappingDic.ToList().Where(x => x.Key == curPropertyName).FirstOrDefault();
                    if (!string.IsNullOrEmpty(refMapping.Key))
                    {
                        try
                        {
                            var value = Convert.ChangeType(oSheet.Cells[rowIdx, refMapping.Value].Value, item.PropertyType);
                            item.SetValue(entity, value, null);
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
                entities.Add(entity);
            }
        }
        #endregion

        #region SimpleWriteData 将数据简单写入表
        /// <summary>
        /// 将数据简单写入表
        /// </summary>
        /// <param name="oSheet"></param>
        /// <param name="dicDatas"></param>
        /// <param name="iHeaderRowIdx"></param>
        protected static void SimpleWriteData(ExcelWorksheet oSheet, IList<Dictionary<string, object>> dicDatas, int iHeaderRowIdx)
        {
            if (dicDatas != null && dicDatas.Count > 0)
            {
                #region 标题行
                {
                    var defaultDic = dicDatas[0].ToArray();
                    for (int idx = 0, len = defaultDic.Length; idx < len; idx++)
                    {
                        var curKeyValue = defaultDic[idx];
                        oSheet.Cells[iHeaderRowIdx, idx + 1].Value = curKeyValue.Key;
                    }
                }
                #endregion

                #region 数据行
                for (int idx = 0, len = dicDatas.Count; idx < len; idx++)
                {
                    var curDic = dicDatas[idx].ToArray();
                    for (int kvIdx = 0, dicLen = curDic.Length; kvIdx < dicLen; kvIdx++)
                    {
                        var curKeyValue = curDic[kvIdx];
                        oSheet.Cells[iHeaderRowIdx + 1 + idx, kvIdx + 1].Value = curKeyValue.Value;
                    }
                }
                #endregion
            }
        }
        #endregion
    }
}
