using LinqToExcel;
using OfficeOpenXml;
using OrderAllot.Entities;
using OrderAllot.Maps;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OrderAllot
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }


        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpload.Text = OpenFileDialog1.FileName;
                btnAnalyze.Enabled = true;
            }
        }

        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            try
            {
                #region 解析并计算
                var workPieceList = new List<WorkPiece>();
                var orderStateList = new List<OrderState>();
                var buyers = new List<string>();//采购员唯一队列
                var excelPath = txtUpload.Text;
                var actRead = new Action(() =>
                {
                    ShowMsg("开始读取表格数据");
                    using (var excel = new ExcelQueryFactory(excelPath))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<OrderState>(s)
                                          select c;
                                orderStateList.AddRange(tmp);
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
                    buyers = orderStateList.Where(s => !string.IsNullOrEmpty(s._采购员)).Select(s => s._采购员).Distinct().ToList();
                    buyers.ForEach(bu =>
                    {
                        var curStateInfos = orderStateList.Where(x => x._采购员 == bu);
                        var curFinishStateInfo = curStateInfos.Where(x => x._是否完成 == true);
                        var curUnFinishStateInfo = curStateInfos.Where(x => x._是否完成 == false);

                        var workItem = new WorkPiece();
                        workItem._采购员 = bu;
                        workItem._完成金额 = curFinishStateInfo.Select(x => x._总金额).Sum();
                        workItem._未完成金额 = curUnFinishStateInfo.Select(x => x._总金额).Sum();
                        workItem._完成单量 = curFinishStateInfo.Count();
                        workItem._未完成单量 = curUnFinishStateInfo.Count();
                        workPieceList.Add(workItem);
                    });

                    //计算完毕,开始导出数据
                    ExportExcel(workPieceList.OrderByDescending(u=>u._完成单量).ToList());

                }, null);
                #endregion
            }
            catch (Exception ex)
            {
                ShowMsg(ex.Message);
            }
        }



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
        private void ExportExcel(List<WorkPiece> workPiece)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[1, 1].Value = "采购员";
                sheet1.Cells[1, 2].Value = "完成金额";
                sheet1.Cells[1, 3].Value = "未完成金额";
                sheet1.Cells[1, 4].Value = "总计金额";
                sheet1.Cells[1, 5].Value = "完成单量";
                sheet1.Cells[1, 6].Value = "未完成单量";
                sheet1.Cells[1, 7].Value = "订单总计";
                sheet1.Cells[1, 8].Value = "平均每单金额";
                sheet1.Cells[1, 9].Value = "完成率";
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = workPiece.Count; idx < len; idx++, rowIdx++)
                {
                    var curWorkItem = workPiece[idx];
                    sheet1.Cells[rowIdx, 1].Value = curWorkItem._采购员;
                    sheet1.Cells[rowIdx, 2].Value = curWorkItem._完成金额;
                    sheet1.Cells[rowIdx, 3].Value = curWorkItem._未完成金额;
                    sheet1.Cells[rowIdx, 4].Value = curWorkItem._总计金额;
                    sheet1.Cells[rowIdx, 5].Value = curWorkItem._完成单量;
                    sheet1.Cells[rowIdx, 6].Value = curWorkItem._未完成单量;
                    sheet1.Cells[rowIdx, 7].Value = curWorkItem._总计;
                    sheet1.Cells[rowIdx, 8].Value = curWorkItem._平均每单金额;
                    sheet1.Cells[rowIdx, 9].Value = curWorkItem._完成率;
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
