﻿using LinqToExcel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CommonLibs;
using LinqToExcel.Attributes;
using OfficeOpenXml.Style;

namespace Gadget
{
    public partial class _库存盘点 : Form
    {
        public _库存盘点()
        {
            InitializeComponent();
        }

        private void _库存盘点_Load(object sender, EventArgs e)
        {
            //txtUpJiaoHuo.Text = @"C:\Users\Leon\Desktop\demo\拣货单.csv";
            //txtUpKucun.Text = @"C:\Users\Leon\Desktop\demo\6-15昆山仓全部库存.csv";
        }

        /**************** button event ****************/

        #region 上传拣货表
        private void btnUpJiaoHuo_Click(object sender, EventArgs e)
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
                    txtUpJiaoHuo.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传库存表
        private void btnUpKucun_Click(object sender, EventArgs e)
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
                    txtUpKucun.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传出入库差
        private void btnChuRuKu_Click(object sender, EventArgs e)
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
                    txtChuRuKu.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 数据处理
        private void btnAnalyze_Click(object sender, EventArgs e)
        {

            var list拣货信息 = new List<_拣货表>();
            var list库存信息 = new List<_库存表>();
            var list出入库差 = new List<_出入库差>();
            var list结果信息 = new List<_导出表>();
            btnAnalyze.Enabled = false;
            #region 读取数据
            var actReadData = new Action(() =>
            {
                ShowMsg("开始读取表格数据");

                var str拣货表Path = txtUpJiaoHuo.Text;
                var str库存表Path = txtUpKucun.Text;
                var str出入库差Path = txtChuRuKu.Text;

                #region 读取拣货表
                if (!string.IsNullOrEmpty(str拣货表Path))
                {
                    using (var csv = new ExcelQueryFactory(str拣货表Path))
                    {
                        try
                        {
                            var tmp = from c in csv.Worksheet<_拣货表>()
                                      select c;
                            list拣货信息.AddRange(tmp);
                        }
                        catch (Exception ex)
                        {
                            ShowMsg(ex.Message);
                        }
                    }
                }
                #endregion

                #region 读取库存表
                if (!string.IsNullOrEmpty(str库存表Path))
                {
                    using (var csv = new ExcelQueryFactory(str库存表Path))
                    {
                        try
                        {
                            var tmp = from c in csv.Worksheet<_库存表>()
                                      select c;
                            list库存信息.AddRange(tmp);
                        }
                        catch (Exception ex)
                        {
                            ShowMsg(ex.Message);
                        }
                    }
                }
                #endregion

                #region 出入库差
                if (!string.IsNullOrEmpty(str出入库差Path))
                {
                    using (var csv = new ExcelQueryFactory(str出入库差Path))
                    {
                        try
                        {
                            var tmp = from c in csv.Worksheet<_出入库差>()
                                      select c;
                            list出入库差.AddRange(tmp);
                        }
                        catch (Exception ex)
                        {
                            ShowMsg(ex.Message);
                        }
                    }
                }
                #endregion

            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((ob) =>
            {
                ShowMsg("开始计算数据");
                var list拣货SKU = list拣货信息.Select(x => x._SKU).Distinct().ToList();

                #region 拣货单匹配子SKU
                list拣货SKU.ForEach(curSKU =>
                {
                    //if (curSKU == "FEDA14A20HP")
                    //{

                    //}

                    /*
                     * 两类sku,一类为CLBA10A94-B,CLBA10A94,其中第一个是第二个的子sku,我们称第一个是父sku
                     * 父sku一般不含中间横线,且一般以数值结尾
                     * 另一类为CLBA12A2F,CLBA12A2A,CLBA12A2B,不含中间横线,一般这个最后的字母是代表子类
                     * 所以截取这个字母前的字符作为父sku
                     */
                    var stParentPart = string.Empty;

                    #region 匹配第一类sku
                    if (curSKU.Contains('-'))
                    {
                        var indx = curSKU.IndexOf('-');
                        if (indx != -1)
                            stParentPart = curSKU.Substring(0, indx);
                    }
                    #endregion

                    #region 匹配第二类sku
                    if (string.IsNullOrEmpty(stParentPart))
                    {
                        stParentPart = CutParent(curSKU);
                    }
                    #endregion

                    stParentPart = !string.IsNullOrEmpty(stParentPart) ? stParentPart : curSKU;

                    var refStoreItems = list库存信息.Where(sto => sto._SKU.Contains(stParentPart)).ToList();
                    if (refStoreItems != null && refStoreItems.Count > 0)
                    {
                        #region 第一个默认是拣货单sku
                        {
                            var defaultItem = refStoreItems.Where(x => x._SKU == curSKU).FirstOrDefault();
                            if (defaultItem != null)
                            {
                                var bExist = list结果信息.Count(x => x._SKU == defaultItem._SKU) > 0;
                                if (!bExist)
                                {
                                    var data = new _导出表();
                                    data._SKU = defaultItem._SKU;
                                    data._商品名称 = defaultItem._商品名称;
                                    data._单位 = defaultItem._单位;
                                    data._可用数量 = defaultItem._可用数量;
                                    data._库存数量 = defaultItem._库存数量;
                                    data._占用数量 = defaultItem._占用数量;
                                    data._库位 = defaultItem._库位;
                                    var ref拣货Item = list拣货信息.Where(x => x._SKU == defaultItem._SKU).FirstOrDefault();
                                    if (ref拣货Item != null)
                                    {
                                        data._拣货单数量 = ref拣货Item._拣货单数量;
                                        data._缺货数量 = ref拣货Item._缺货数量;
                                    }
                                    list结果信息.Add(data);
                                }
                            }
                        }
                        #endregion

                        #region 余下的子sku
                        {
                            var remainItems = refStoreItems.Where(x => x._SKU != curSKU).ToList();
                            if (remainItems != null && remainItems.Count > 0)
                            {
                                foreach (var item in remainItems)
                                {
                                    var bExist = list结果信息.Count(x => x._SKU == item._SKU) > 0;
                                    if (!bExist)
                                    {
                                        var data = new _导出表();
                                        data._SKU = item._SKU;
                                        data._商品名称 = item._商品名称;
                                        data._单位 = item._单位;
                                        data._可用数量 = item._可用数量;
                                        data._库存数量 = item._库存数量;
                                        data._占用数量 = item._占用数量;
                                        data._库位 = item._库位;
                                        var ref拣货Item = list拣货信息.Where(x => x._SKU == item._SKU).FirstOrDefault();
                                        if (ref拣货Item != null)
                                        {
                                            data._拣货单数量 = ref拣货Item._拣货单数量;
                                            data._缺货数量 = ref拣货Item._缺货数量;
                                        }
                                        list结果信息.Add(data);
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                });
                #endregion

                #region 匹配出入库差
                {
                    if (list出入库差.Count > 0 && list结果信息.Count > 0)
                    {
                        for (int i = 0, len = list结果信息.Count; i < len; i++)
                        {
                            var curItem = list结果信息[i];
                            var refDiff = list出入库差.Where(x => x._SKU == curItem._SKU).FirstOrDefault();
                            if (refDiff != null)
                            {
                                curItem._出库数量 = refDiff._出库数量;
                                curItem._入库数量 = refDiff._入库数量;
                                curItem._出入库差值 = refDiff._出入库差值;
                            }
                        }
                    }
                }
                #endregion

                Export(list结果信息);
            }, null);
            #endregion
        }
        #endregion

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_拣货表), typeof(_库存表), typeof(_出入库差));

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

        /**************** common method ****************/

        #region 匹配第二 类sku
        private string CutParent(string strSKU)
        {
            var msg = 0;
            var iCutLength = 0;
            if (!string.IsNullOrEmpty(strSKU))
            {
                for (int idx = strSKU.Length - 1; idx >= 0; idx--)
                {
                    var cschar = strSKU.Substring(idx, 1);
                    var bIsNumber = int.TryParse(cschar, out msg);
                    if (bIsNumber)
                    {
                        break;
                    }
                    iCutLength++;
                }
            }
            if (iCutLength > 0)
            {
                return strSKU.Substring(0, strSKU.Length - iCutLength);
            }
            return strSKU;
        }
        #endregion

        #region Export 导出结果表格
        private void Export(List<_导出表> list)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];

            #region 数据表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;


                #region 总表
                {
                    var sheet1 = workbox.Worksheets.Add("总表");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "库位";
                    sheet1.Cells[1, 2].Value = "SKU";
                    sheet1.Cells[1, 3].Value = "商品名称";
                    sheet1.Cells[1, 4].Value = "单位";
                    sheet1.Cells[1, 5].Value = "库存数量";
                    sheet1.Cells[1, 6].Value = "占用数量";
                    sheet1.Cells[1, 7].Value = "可用数量";
                    sheet1.Cells[1, 8].Value = "盘点数量";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._库位;
                        sheet1.Cells[rowIdx, 2].Value = info._SKU;
                        sheet1.Cells[rowIdx, 3].Value = info._商品名称;
                        sheet1.Cells[rowIdx, 4].Value = info._单位;
                        sheet1.Cells[rowIdx, 5].Value = info._库存数量;
                        sheet1.Cells[rowIdx, 6].Value = info._占用数量;
                        sheet1.Cells[rowIdx, 7].Value = info._可用数量;
                    }
                    #endregion

                    #region 样式设置
                    {
                        sheet1.Cells[1, 1, 1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中

                        sheet1.Column(1).Width = 13;
                        sheet1.Column(2).Width = 18.43;
                        sheet1.Column(3).Width = 40;
                        sheet1.Column(4).Width = 4.86;

                    }
                    #endregion
                }
                #endregion


                var areas = list.Select(x => !string.IsNullOrWhiteSpace(x._库位) ? x._库位.Substring(0, 1).ToUpper() : "").Distinct().OrderBy(x => x).ToList();

                foreach (var areaName in areas)
                {
                    if (string.IsNullOrWhiteSpace(areaName))
                        continue;
                    var sheet1 = workbox.Worksheets.Add(string.Format("{0}区", areaName));
                    var referDatas = list.Where(x => x._库位.Substring(0, 1).ToUpper() == areaName).ToList();

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "库位";
                    sheet1.Cells[1, 2].Value = "SKU";
                    sheet1.Cells[1, 3].Value = "商品名称";
                    sheet1.Cells[1, 4].Value = "单位";
                    sheet1.Cells[1, 5].Value = "盘点数量";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = referDatas.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = referDatas[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._库位;
                        sheet1.Cells[rowIdx, 2].Value = info._SKU;
                        sheet1.Cells[rowIdx, 3].Value = info._商品名称;
                        sheet1.Cells[rowIdx, 4].Value = info._单位;

                    }
                    #endregion

                    #region 样式设置
                    {
                        sheet1.Cells[1,1,1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中

                        sheet1.Column(1).Width = 13;
                        sheet1.Column(2).Width = 18.43;
                        sheet1.Column(3).Width = 40;
                        sheet1.Column(4).Width = 4.86;
                        sheet1.Column(5).Width = 9.43;

                        var endRow = sheet1.Dimension.End.Row;
                        var endColumn = sheet1.Dimension.End.Column;
                        using (var rng = sheet1.Cells[1, 1, endRow, endColumn])
                        {
                            rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        }
                    }
                    #endregion

                }




                buffer = package.GetAsByteArray();
            }
            #endregion



            InvokeMainForm((obj) =>
            {
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
                saveFile.Title = "导出数据";//设置标题
                saveFile.AddExtension = true;//是否自动增加所辍名
                saveFile.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
                if (saveFile.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
                {
                    var FileName = saveFile.FileName;//得到文件路径   
                    var saveFilName = Path.GetFileNameWithoutExtension(FileName);

                    txtExport.Text = FileName;
                    var len = buffer.Length;
                    using (var fs = File.Create(FileName, len))
                    {
                        fs.Write(buffer, 0, len);
                    }
                    ShowMsg("表格生成完毕");
                    btnAnalyze.Enabled = true;
                }
            }, null);
        }
        #endregion

        #region ShowMsg 消息提示
        /// <summary>
        /// 消息提示
        /// </summary>
        /// <param name="strMsg"></param>
        private void ShowMsg(string strMsg)
        {
            if (this.InvokeRequired)
            {
                var act = new Action<string>(ShowMsg);
                this.Invoke(act, strMsg);
            }
            else
            {
                this.lbMsg.Text = strMsg;
            }
        }
        #endregion

        #region InvokeMainForm 调用主线程
        protected void InvokeMainForm(Action act)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(act);
            }
            else
            {
                act.Invoke();
            }
        }
        protected void InvokeMainForm(Action<object> act, object obj)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(act, obj);
            }
            else
            {
                act.Invoke(obj);
            }
        }
        #endregion

        /**************** common class ****************/

        [ExcelTable("拣货表")]
        class _拣货表
        {
            private string strSKU;
            [ExcelColumn("SKU")]
            public string _SKU
            {
                get
                {
                    return strSKU;
                }
                set
                {
                    strSKU = !string.IsNullOrEmpty(value) ? value.Trim() : "";
                }
            }
            [ExcelColumn("拣货单数量")]
            public decimal _拣货单数量 { get; set; }
            [ExcelColumn("缺货数量")]
            public decimal _缺货数量 { get; set; }
        }

        [ExcelTable("库存表")]
        class _库存表
        {
            private string strSKU;
            [ExcelColumn("SKU码")]
            public string _SKU
            {
                get
                {
                    return strSKU;
                }
                set
                {
                    strSKU = !string.IsNullOrEmpty(value) ? value.Trim() : "";
                }
            }
            [ExcelColumn("库存数量")]
            public decimal _库存数量 { get; set; }
            [ExcelColumn("占用数量")]
            public decimal _占用数量 { get; set; }
            [ExcelColumn("可用数量")]
            public decimal _可用数量 { get; set; }
            [ExcelColumn("库位")]
            public string _库位 { get; set; }
            [ExcelColumn("商品名称")]
            public string _商品名称 { get; set; }
            [ExcelColumn("单位")]
            public string _单位 { get; set; }
        }

        [ExcelTable("出入库差")]
        class _出入库差
        {
            private string strSKU;
            [ExcelColumn("SKU码")]
            public string _SKU
            {
                get
                {
                    return strSKU;
                }
                set
                {
                    strSKU = !string.IsNullOrEmpty(value) ? value.Trim() : "";
                }
            }
            [ExcelColumn("入库数量")]
            public decimal _入库数量 { get; set; }
            [ExcelColumn("出库数量")]
            public decimal _出库数量 { get; set; }
            [ExcelColumn("入库数量-出库数量")]
            public decimal _出入库差值 { get; set; }
        }

        class _导出表
        {
            public string _SKU { get; set; }
            public decimal _库存数量 { get; set; }
            public decimal _占用数量 { get; set; }
            public decimal _可用数量 { get; set; }
            public decimal _拣货单数量 { get; set; }
            public decimal _缺货数量 { get; set; }
            public string _库位 { get; set; }
            public decimal _入库数量 { get; set; }
            public decimal _出库数量 { get; set; }
            public decimal _出入库差值 { get; set; }
            public string _商品名称 { get; set; }
            public string _单位 { get; set; }
        }
    }
}
