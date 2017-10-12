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
        }

        #region 上传上海仓库延时报表
        private void btnUpShRp_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
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
            OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
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

        #region 处理数据
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            try
            {
                #region 解析并计算
                var shRpList = new List<_Form5延时报表>();
                var ksRpList = new List<_Form5延时报表>();
                var shExcelPath = txtUpShRp.Text;
                var ksExcelPath = txtUpShRp.Text;

                var actRead = new Action(() =>
                {
                    ShowMsg("开始读取表格数据");
                    using (var excel = new ExcelQueryFactory(shExcelPath))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<_Form5延时报表>(s)
                                          select c;
                                shRpList.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                    }
                    using (var excel = new ExcelQueryFactory(ksExcelPath))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<_Form5延时报表>(s)
                                          select c;
                                ksRpList.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                    }
                });


                actRead.BeginInvoke((obj) =>
                {
                    ShowMsg("开始计算表格数据");
                    var cmSkus = new List<string>();
                    for (int idx = shRpList.Count - 1; idx >= 0; idx--)
                    {
                        var shItem = shRpList[idx];
                        var refksItem = ksRpList.Where(k => k._SKU == shItem._SKU).FirstOrDefault();
                        if (refksItem != null)
                        {
                            cmSkus.Add(shItem._SKU);
                            //两个仓库都有,将两个仓库的 可用库存+可用库存-缺货及未派单-缺货及未派单>0即为有库存,需要导出表格
                            var amount = shItem._可用数量 + refksItem._可用数量 - shItem._缺货及未派单数量 - refksItem._缺货及未派单数量;
                            if (amount < 0)
                            {
                                shRpList.RemoveAt(idx);
                            }
                        }
                        else
                        {
                            //只有上海仓库有,也判断 可用库存-缺货及未派单>0 即导出
                            var amount = shItem._可用数量 - shItem._缺货及未派单数量;
                            if (amount < 0)
                            {
                                shRpList.RemoveAt(idx);
                            }
                        }
                    }
                    //上海仓库计算完了,然后把昆山仓库加上去
                    ksRpList.ForEach(ks =>
                    {
                        var isDuplicate = cmSkus.Count(c => c == ks._SKU);
                        if (isDuplicate == 0)
                        {
                            var amount = ks._可用数量 - ks._缺货及未派单数量;
                            if (amount > 0)
                                shRpList.Add(ks);
                        }
                    });
                    //计算完毕,开始导出数据
                    ExportExcel(shRpList);
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
        private void ExportExcel(List<_Form5延时报表> rpList)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
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
                    try
                    {
                        var len = buffer.Length;
                        using (var fs = File.Create(FileName, len))
                        {
                            fs.Write(buffer, 0, len);
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
