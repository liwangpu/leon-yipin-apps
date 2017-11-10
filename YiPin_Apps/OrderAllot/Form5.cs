using LinqToExcel;
using OfficeOpenXml;
using OrderAllot.Maps;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OrderAllot
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();


            txtUpShRp.Text = @"C:\Users\pulw\Desktop\延时报表\上海所有库存1.xlsx";
            txtUpKsRp.Text = @"C:\Users\pulw\Desktop\延时报表\昆山所有库存1.xlsx";
            txtQueh.Text = @"C:\Users\pulw\Desktop\延时报表\缺货延时报表1.xlsx";
            btnAnalyze.Enabled = true;

        }

        #region 上传上海仓库延时报表
        private void btnUpShRp_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpShRp.Text = OpenFileDialog1.FileName;
                btnAnalyze.Enabled = true;
            }
        }
        #endregion

        #region 上传上昆山仓库延时报表
        private void btnUpKsRp_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpKsRp.Text = OpenFileDialog1.FileName;
                btnAnalyze.Enabled = true;
            }
        }
        #endregion

        #region 上传缺货延时报表
        private void btnQueh_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtQueh.Text = OpenFileDialog1.FileName;
            }
        }
        #endregion

        #region 处理数据
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            try
            {
                #region 解析并计算
                var list上海库存 = new List<_Form5延时报表>();
                var list昆山库存 = new List<_Form5延时报表>();
                var quehList = new List<_Form5缺货延时报表判断>();
                var str上海库存Path = txtUpShRp.Text;
                var str昆山库存Path = txtUpKsRp.Text;
                var str缺货延时报表Path = txtQueh.Text;

                var actRead = new Action(() =>
                {
                    ShowMsg("开始读取表格数据");
                    if (!string.IsNullOrEmpty(str上海库存Path))
                    {
                        using (var excel = new ExcelQueryFactory(str上海库存Path))
                        {
                         

                            try
                            {
                                var sheetNames = excel.GetWorksheetNames().ToList();
                                sheetNames.ForEach(s =>
                                {
                                    try
                                    {
                                        var tmp = from c in excel.Worksheet<_Form5延时报表>(s)
                                                  select c;
                                        list上海库存.AddRange(tmp);
                                    }
                                    catch (Exception ex)
                                    {
                                        ShowMsg(ex.Message);
                                    }
                                });
                            }
                            catch (Exception ex)
                            {

                                var aaaa = 1;
                            }

                
                        }
                    }
                    if (!string.IsNullOrEmpty(str昆山库存Path))
                    {
                        using (var excel = new ExcelQueryFactory(str昆山库存Path))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<_Form5延时报表>(s)
                                              select c;
                                    list昆山库存.AddRange(tmp);
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }
                    if (!string.IsNullOrEmpty(str缺货延时报表Path))
                    {
                        using (var excel = new ExcelQueryFactory(str缺货延时报表Path))
                        {
                            var sheetNames = excel.GetWorksheetNames().ToList();
                            sheetNames.ForEach(s =>
                            {
                                try
                                {
                                    var tmp = from c in excel.Worksheet<_Form5缺货延时报表判断>(s)
                                              select c;
                                    quehList.AddRange(tmp);
                                }
                                catch (Exception ex)
                                {
                                    ShowMsg(ex.Message);
                                }
                            });
                        }
                    }

                });


                actRead.BeginInvoke((obj) =>
                {
                    ShowMsg("开始计算表格数据");
                    var cmSkus = new List<string>();
                    for (int idx = list上海库存.Count - 1; idx >= 0; idx--)
                    {
                        var shItem = list上海库存[idx];
                        var refksItem = list昆山库存.Where(k => k._SKU == shItem._SKU).FirstOrDefault();
                        if (refksItem != null)
                        {
                            cmSkus.Add(shItem._SKU);
                            //两个仓库都有,将两个仓库的 可用库存+可用库存-缺货及未派单-缺货及未派单>0即为有库存,需要导出表格
                            var amount = shItem._可用数量 + refksItem._可用数量 - shItem._缺货及未派单数量 - refksItem._缺货及未派单数量;
                            if (amount < 0)
                            {
                                list上海库存.RemoveAt(idx);
                            }
                            else
                            {
                                var _可用数量 = (shItem._可用数量 + refksItem._可用数量).ToString();
                                var _缺货以及未派单 = (shItem._缺货及未派单数量 + refksItem._缺货及未派单数量).ToString();
                                shItem.str可用数量 = _可用数量;
                                shItem.str缺货及未派单数量 = _缺货以及未派单;
                            }
                        }
                        else
                        {
                            //只有上海仓库有,也判断 可用库存-缺货及未派单>0 即导出
                            var amount = shItem._可用数量 - shItem._缺货及未派单数量;
                            if (amount < 0)
                            {
                                list上海库存.RemoveAt(idx);
                            }
                        }
                    }
                    //上海仓库计算完了,然后把昆山仓库加上去
                    list昆山库存.ForEach(ks =>
                    {
                        var isDuplicate = cmSkus.Count(c => c == ks._SKU);
                        if (isDuplicate == 0)
                        {
                            var amount = ks._可用数量 - ks._缺货及未派单数量;
                            if (amount >= 0)
                                list上海库存.Add(ks);
                        }
                    });


                    if (quehList != null && quehList.Count > 0)
                    {
                        for (int idx = quehList.Count - 1; idx >= 0; idx--)
                        {
                            var curQueh = quehList[idx];
                            if (curQueh._是否停售 != "停售")
                            {
                                var isExist = list上海库存.Where(x => x._SKU == curQueh.sku).Count() > 0;
                                if (isExist)
                                {
                                    quehList.RemoveAt(idx);
                                }
                            }
                            else
                            {
                                quehList.RemoveAt(idx);
                            }
                        }
                    }


                    //计算完毕,开始导出数据
                    ExportExcel(list上海库存, quehList);
                }, null);
                #endregion
            }
            catch (Exception ex)
            {
                ShowMsg(ex.Message);
            }
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

        #region ExportExcel 导出Excel表格
        /// <summary>
        /// 导出Excel表格
        /// </summary>
        /// <param name="workPiece"></param>
        private void ExportExcel(List<_Form5延时报表> rpList, List<_Form5缺货延时报表判断> queList)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer1 = new byte[0];
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[1, 1].Value = "SKU码";
                sheet1.Cells[1, 2].Value = "库存数量";
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = rpList.Count; idx < len; idx++, rowIdx++)
                {
                    var curItem = rpList[idx];
                    sheet1.Cells[rowIdx, 1].Value = curItem._SKU;
                    sheet1.Cells[rowIdx, 2].Value = curItem._可用数量 - curItem._缺货及未派单数量;
                }
                #endregion


                buffer = package.GetAsByteArray();
            }

            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[1, 1].Value = "sku";
                sheet1.Cells[1, 2].Value = "是否停售";
                sheet1.Cells[1, 3].Value = "停售时间";
                sheet1.Cells[1, 4].Value = "店铺SKU";
                sheet1.Cells[1, 5].Value = "订单编号";
                sheet1.Cells[1, 6].Value = "卖家简称";
                sheet1.Cells[1, 7].Value = "Itemid";
                sheet1.Cells[1, 8].Value = "延时时间";
                sheet1.Cells[1, 9].Value = "交易时间";
                sheet1.Cells[1, 10].Value = "业绩归属1";
                sheet1.Cells[1, 11].Value = "业绩归属2";
                sheet1.Cells[1, 12].Value = "采购员";
                sheet1.Cells[1, 13].Value = "平台";
                sheet1.Cells[1, 14].Value = "本订单销售数量";
                sheet1.Cells[1, 15].Value = "库存数量";
                sheet1.Cells[1, 15].Value = "占用数量";
                sheet1.Cells[1, 17].Value = "缺货总数";
                sheet1.Cells[1, 18].Value = "可用数量";
                sheet1.Cells[1, 19].Value = "发货仓库";
                #endregion


                #region 数据行
                for (int idx = 0, rowIdx = 2, len = queList.Count; idx < len; idx++, rowIdx++)
                {
                    var curItem = queList[idx];
                    sheet1.Cells[rowIdx, 1].Value = curItem.sku;
                    sheet1.Cells[rowIdx, 2].Value = curItem._是否停售;
                    sheet1.Cells[rowIdx, 3].Value = curItem._停售时间;
                    sheet1.Cells[rowIdx, 4].Value = curItem._店铺SKU;
                    sheet1.Cells[rowIdx, 5].Value = curItem._订单编号;
                    sheet1.Cells[rowIdx, 6].Value = curItem._卖家简称;
                    sheet1.Cells[rowIdx, 7].Value = curItem._Itemid;
                    sheet1.Cells[rowIdx, 8].Value = curItem._延时时间;
                    sheet1.Cells[rowIdx, 9].Value = curItem._交易时间;
                    sheet1.Cells[rowIdx, 10].Value = curItem._业绩归属1;
                    sheet1.Cells[rowIdx, 11].Value = curItem._业绩归属2;
                    sheet1.Cells[rowIdx, 12].Value = curItem._采购员;
                    sheet1.Cells[rowIdx, 13].Value = curItem._平台;
                    sheet1.Cells[rowIdx, 14].Value = curItem._本订单销售数量;
                    sheet1.Cells[rowIdx, 15].Value = curItem._库存数量;
                    sheet1.Cells[rowIdx, 16].Value = curItem._占用数量;
                    sheet1.Cells[rowIdx, 17].Value = curItem._缺货总数;
                    sheet1.Cells[rowIdx, 18].Value = curItem._可用数量;
                    sheet1.Cells[rowIdx, 19].Value = curItem._发货仓库;

                }
                #endregion

                buffer1 = package.GetAsByteArray();
            }



            InvokeMainForm((obj) =>
            {

                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
                saveFile.Title = "导出数据";//设置标题
                saveFile.AddExtension = true;//是否自动增加所辍名
                saveFile.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
                if (saveFile.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
                {
                    string FileName = saveFile.FileName;//得到文件路径   
                    txtExport.Text = FileName;

                    var saveFilName = Path.GetFileNameWithoutExtension(FileName);
                    var savePath = Path.GetDirectoryName(FileName);
                    var FileName1 = Path.Combine(savePath, saveFilName + "延时报表.xlsx");


                    try
                    {
                        var len = buffer.Length;
                        using (var fs = File.Create(FileName, len))
                        {
                            fs.Write(buffer, 0, len);
                        }

                        var len1 = buffer1.Length;
                        using (var fs = File.Create(FileName1, len1))
                        {
                            fs.Write(buffer1, 0, len1);
                        }
                    }
                    catch (Exception ex)
                    {
                        ShowMsg(ex.Message);
                    }

                    ShowMsg("表格生成完毕");
                    btnAnalyze.Enabled = false;
                }
            }, null);
        }
        #endregion


    }
}
