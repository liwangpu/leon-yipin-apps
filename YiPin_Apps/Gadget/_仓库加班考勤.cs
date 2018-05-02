﻿using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Gadget
{
    public partial class _仓库加班考勤 : Form
    {
        public _仓库加班考勤()
        {
            InitializeComponent();
        }

        private void _仓库加班考勤_Load(object sender, EventArgs e)
        {

        }

        /**************** button event ****************/

        #region 上传考勤信息
        private void btn上传考勤表_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txt考勤.Text = OpenFileDialog1.FileName;
                btn计算考勤.Enabled = true;
            }
        }
        #endregion

        #region 计算考勤
        private void btn当天考勤_Click(object sender, EventArgs e)
        {
            btn计算考勤.Enabled = false;
            var strError = string.Empty;
            var list考勤数据 = new List<_考勤数据>();
            var list加班绩效 = new List<_加班绩效>();

            #region 读取数据
            var actReadData = new Action(() =>
            {
                ShowMsg("开始读取当天考勤信息");
                using (var package = new ExcelPackage(new FileInfo(txt考勤.Text)))
                {
                    var worksheet = package.Workbook.Worksheets["考勤记录"];//创建worksheet
                    var endRow = worksheet.Dimension.End.Row;
                    var endColumn = worksheet.Dimension.End.Column;
                    for (int idx = 5; idx <= endRow; idx = idx + 2)
                    {
                        var md = new _考勤数据();
                        md._姓名 = worksheet.Cells[idx, 11].Value.ToString();
                        md._员工序号 = Convert.ToInt32(worksheet.Cells[idx, 3].Value);
                        var list = new List<string>();
                        for (int cll = 1; cll <= endColumn; cll++)
                        {
                            var vl = worksheet.Cells[idx + 1, cll].Value != null ? worksheet.Cells[idx + 1, cll].Value.ToString() : "";
                            list.Add(vl);
                        }
                        md._加班信息 = list;
                        list考勤数据.Add(md);
                    }
                }
            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                ShowMsg("考勤数据读取完毕,即将开始计算");
                for (int idx = list考勤数据.Count - 1; idx >= 0; idx--)
                {
                    var cur = list考勤数据[idx];
                    var md = new _加班绩效();
                    md._员工序号 = cur._员工序号;
                    md._姓名 = cur._姓名;

                    //if (md._姓名 == "吴赞")
                    //{

                    //}
                    var list = new List<double>();
                    var _list打卡情况 = new List<string>();
                    for (int nnn = 0, count = cur._加班信息.Count; nnn < count; nnn++)
                    {
                        var timeStr = !string.IsNullOrWhiteSpace(cur._加班信息[nnn]) ? cur._加班信息[nnn].Trim() : "";
                        if (!string.IsNullOrWhiteSpace(timeStr))
                        {

                            //if (nnn == 8)
                            //{

                            //}
                            double total = 0;
                            var d上班时间 = Convert.ToDateTime("2018-08-08 09:00:00");
                            var d下班时间 = Convert.ToDateTime("2018-08-08 17:30:00");
                            //上班时间加班
                            {
                                var startStr = string.Format("2018-08-08 {0}:00", timeStr.Substring(0, 5));
                                var startTime = Convert.ToDateTime(startStr);
                                var timespan = (d上班时间 - startTime).TotalMinutes;
                                if (timespan >= 55)
                                {
                                    var mm = timespan % 60;
                                    var hh = (timespan - mm) / 60;
                                    total += hh;
                                    if (mm >= 55)
                                        total += 1;
                                }
                            }
                            //下班时间加班
                            {
                                if (timeStr.Length > 5)
                                {
                                    var _d包饭时间 = Convert.ToDateTime("2018-08-08 21:00:00");
                                    var endStr = string.Format("2018-08-08 {0}:00", timeStr.Substring(5, 5));
                                    var endTime = Convert.ToDateTime(endStr);
                                    var timespan = (endTime - d下班时间).TotalMinutes;
                                    if (timespan >= 55)
                                    {
                                        var mm = timespan % 60;
                                        var hh = (timespan - mm) / 60;
                                        total += hh;
                                        if (mm >= 55)
                                            total += 1;
                                        if (mm >= 30 && mm < 55)
                                            total += 0.5;

                                        if (endTime >= _d包饭时间)
                                            total -= 0.5;
                                    }
                                }
                            }
                            list.Add(total);
                            _list打卡情况.Add("已打卡");
                        }
                        else
                        {
                            _list打卡情况.Add("未打卡");
                            list.Add(0);
                        }
                    }
                    md._加班时长 = list;
                    md._打卡情况 = _list打卡情况;
                    list加班绩效.Add(md);
                }
                ExportExcel(list加班绩效);
            }, null);
            #endregion

        }
        #endregion

        #region 导出表格说明事件
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //FormHelper.GenerateTableDes(typeof(_拣货单), typeof(_拣货时间), typeof(_拣货人员配置));
        }
        #endregion

        /**************** common method ****************/
        #region 导出表格
        private void ExportExcel(List<_加班绩效> list加班绩效)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];

            #region 订单分配
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                using (var rng = sheet1.Cells[1, 1, 3, 1])
                {
                    rng.Value = "姓名";
                    rng.Merge = true;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                }
                sheet1.Cells[1, 2].Value = "星期";
                using (var rng = sheet1.Cells[2, 2, 3, 2])
                {
                    rng.Value = "项目";
                    rng.Merge = true;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                }

                if (list加班绩效[0] != null)
                {
                    string[] Day = new string[] { "周日", "周一", "周二", "周三", "周四", "周五", "周六" };
                    var days = list加班绩效[0]._加班时长.Count + 3;
                    for (int column = 3, idx = 1; column < days; column++, idx++)
                    {
                        var ct = dtp考勤时间.Value;
                        var dateStr = string.Format("{0}-{1}-{2}", ct.Year, ct.Month > 9 ? "" + ct.Month : "0" + ct.Month, idx > 9 ? "" + idx : "0" + idx);
                        var date = DateTime.MinValue;
                        var isValid = DateTime.TryParse(dateStr, out date);
                        if (isValid)
                        {
                            sheet1.Column(column).Width = 5;//设置列宽
                            using (var rng = sheet1.Cells[2, column, 3, column])
                            {
                                rng.Value = idx;
                                rng.Merge = true;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                                rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            }

                            sheet1.Cells[1, column].Value = Day[Convert.ToInt32(Convert.ToDateTime(dateStr).DayOfWeek.ToString("d"))].ToString();
                        }
                    }

                    using (var rng = sheet1.Cells[1, days, 2, days + 2])
                    {
                        rng.Value = "合计加班";
                        rng.Merge = true;
                        sheet1.Column(days).Width = 5;//设置列宽
                        sheet1.Column(days + 1).Width = 5;//设置列宽
                        sheet1.Column(days + 2).Width = 5;//设置列宽
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                    }

                    using (var rng = sheet1.Cells[3, days])
                    {
                        rng.Value = "平时（H）";
                        rng.Style.WrapText = true;//自动换行
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                    }

                    using (var rng = sheet1.Cells[3, days + 1])
                    {
                        rng.Value = "周末（日）";
                        rng.Style.WrapText = true;//自动换行
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                    }

                    using (var rng = sheet1.Cells[3, days + 2])
                    {
                        rng.Value = "节日（日）";
                        rng.Style.WrapText = true;//自动换行
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                    }
                }
                sheet1.Row(3).Height = 79;//设置行高
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 4, len = list加班绩效.Count; idx < len; idx++)
                {
                    var curOrder = list加班绩效[idx];
                    using (var rng = sheet1.Cells[rowIdx, 1, rowIdx + 2, 1])
                    {
                        rng.Value = curOrder._姓名;
                        rng.Merge = true;
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                    }
                    sheet1.Cells[rowIdx, 2].Value = "出勤";
                    sheet1.Cells[rowIdx + 1, 2].Value = "请假";
                    sheet1.Cells[rowIdx + 2, 2].Value = "加班";

                    var _i请假总计 = 0;
                    for (int nnn = 0, nlen = curOrder._加班时长.Count; nnn < nlen; nnn++)
                    {
                        sheet1.Cells[rowIdx + 2, 3 + nnn].Value = curOrder._加班时长[nnn];

                        if (sheet1.Cells[1, 3 + nnn].Value.ToString() != "周日" && curOrder._打卡情况[nnn] == "未打卡")
                        {
                            _i请假总计++;
                            sheet1.Cells[rowIdx + 1, 3 + nnn].Value = 1;
                        }
                    }
                    sheet1.Cells[rowIdx + 2, 3 + curOrder._加班时长.Count].Value = curOrder._加班时长.Sum();
                    sheet1.Cells[rowIdx + 1, 3 + curOrder._加班时长.Count].Value = _i请假总计;
                    rowIdx += 3;
                }
                #endregion

                #region 全部边框
                {
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
                    var savePath = Path.GetDirectoryName(FileName);

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
                }
                btn计算考勤.Enabled = true;
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

        class _考勤数据
        {
            public int _员工序号 { get; set; }
            public string _姓名 { get; set; }
            public List<string> _加班信息 { get; set; }
        }

        class _加班绩效
        {
            public int _员工序号 { get; set; }
            public string _姓名 { get; set; }
            public List<double> _加班时长 { get; set; }
            public List<string> _打卡情况 { get; set; }
        }

    }
}
