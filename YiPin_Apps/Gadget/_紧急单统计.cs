using CommonLibs;
using Gadget.Libs;
using LinqToExcel.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Gadget
{
    public partial class _紧急单统计 : Form
    {
        public _紧急单统计()
        {
            InitializeComponent();
        }

        private void _紧急单统计_Load(object sender, EventArgs e)
        {
            //txt紧急单.Text = @"C:\Users\Leon\Desktop\666.csv";
            //btn计算工作情况.Enabled = true;
        }

        /**************** button event ****************/

        #region 上传紧急单
        private void btn上传紧急单_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt紧急单, () =>
            {
                btn计算工作情况.Enabled = true;
            });
        }
        #endregion

        #region 计算
        private void btn计算工作情况_Click(object sender, EventArgs e)
        {
            var _List原始数据 = new List<_库位明细信息>();
            var settime = dtp截至时间.Value;
            #region 读取数据
            var actReadData = new Action(() =>
            {
                var strError = string.Empty;
                ShowMsg("开始读取紧急单数据");
                FormHelper.ReadCSVFile(txt紧急单.Text, ref _List原始数据, ref strError);

                #region 过滤截至时间之后的数据
                if (_List原始数据.Count > 0)
                {
                    for (int idx = _List原始数据.Count - 1; idx >= 0; idx--)
                    {
                        var curData = _List原始数据[idx];
                        if (!string.IsNullOrWhiteSpace(curData._内部标签))
                        {
                            var strArr = curData._内部标签.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                            if (strArr.Length >= 3)
                            {
                                var strTime = strArr[2].Substring(0, 8);
                                var cTime = Convert.ToDateTime(string.Format("{0} {1}", DateTime.Now.ToString("yyyy-MM-dd"), strTime));
                                if (cTime > settime)
                                    _List原始数据.RemoveAt(idx);
                            }
                        }
                        //else
                        //{
                        //    _List原始数据.RemoveAt(idx);
                        //}
                    }
                }
                #endregion

            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                ShowMsg("紧急单数据读取完毕,即将开始计算");
                var _采购员姓名List = _List原始数据.Where(x => !string.IsNullOrWhiteSpace(x._采购员)).Select(x => x._采购员).Distinct().ToList();
                var _list统计结果 = new List<_统计结果>();
                if (_采购员姓名List.Count > 0)
                    _采购员姓名List.ForEach(name =>
                    {
                        //if (name == "唐汉成")
                        //{

                        //}
                        var referDatas = _List原始数据.Where(x => x._采购员 == name).ToList();
                        var list紧急订单 = referDatas.Where(x => x._备注.Contains("紧急")).ToList();
                        var list完成订单 = referDatas.Where(x => x._内部标签.Contains("付")).ToList();

                        var mode = new _统计结果();
                        mode._采购员 = name;
                        mode._紧急订单数 = list紧急订单.Select(x => x._采购订单号).Distinct().Count();
                        mode._完成订单数 = list完成订单.Select(x => x._采购订单号).Distinct().Count();
                        mode._异常订单 = referDatas.Where(x => x._备注.Contains("标") || x._备注.Contains("缺")).Select(x => x._采购订单号).Distinct().Count();
                        mode._紧急单SKU个数 = list紧急订单.Select(x => x.SKU).Distinct().Count();
                        mode._完成SKU个数 = list完成订单.Select(x => x.SKU).Distinct().Count();

                        var d订单完成占比 = mode._紧急订单数 - mode._异常订单 > 0 ? Math.Round(mode._完成订单数 * 1.0m / (mode._紧急订单数 - mode._异常订单), 4) : 0;
                        mode._订单完成占比 = d订单完成占比;
                        mode._SKU完成占比 = mode._紧急单SKU个数 > 0 ? Math.Round(mode._完成SKU个数 * 1.0m / mode._紧急单SKU个数, 4) : 0;


                        _list统计结果.Add(mode);
                    });

                ExportExcel(_list统计结果);
            }, null);
            #endregion
        } 
        #endregion

        #region 导出说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormHelper.GenerateTableDes(typeof(_库位明细信息));
        } 
        #endregion

        /**************** common method ****************/

        #region 导出表格
        private void ExportExcel(List<_统计结果> list)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];

            #region 订单分配
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                using (var rng = sheet1.Cells[1, 1, 1, 9])
                {
                    rng.Value = DateTime.Now.ToString("yyyy-MM-dd");
                    rng.Merge = true;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                }


                sheet1.Cells[2, 1].Value = "采购员";
                sheet1.Cells[2, 2].Value = "紧急订单数";
                sheet1.Cells[2, 3].Value = "完成订单数";
                sheet1.Cells[2, 4].Value = "异常订单";
                sheet1.Cells[2, 5].Value = "订单完成占比";
                sheet1.Cells[2, 6].Value = "紧急单SKU个数";
                sheet1.Cells[2, 7].Value = "完成SKU个数";
                sheet1.Cells[2, 8].Value = "SKU完成占比";
                sheet1.Cells[2, 9].Value = "缺货订单";

                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 3, len = list.Count; idx < len; idx++)
                {
                    var curData = list[idx];
                    sheet1.Cells[rowIdx, 1].Value = curData._采购员;
                    sheet1.Cells[rowIdx, 2].Value = curData._紧急订单数;
                    sheet1.Cells[rowIdx, 3].Value = curData._完成订单数;
                    sheet1.Cells[rowIdx, 4].Value = curData._异常订单;
                    sheet1.Cells[rowIdx, 5].Value = curData._订单完成占比;
                    sheet1.Cells[rowIdx, 6].Value = curData._紧急单SKU个数;
                    sheet1.Cells[rowIdx, 7].Value = curData._完成SKU个数;
                    sheet1.Cells[rowIdx, 8].Value = curData._SKU完成占比;
                    rowIdx++;
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

                sheet1.Cells[sheet1.Dimension.Address].AutoFitColumns();

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
                btn计算工作情况.Enabled = true;
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

        [ExcelTable("紧急单")]
        class _库位明细信息
        {
            private string _Org备注;
            private string _OrgSku;
            private string _Org采购员;
            private string _Org内部标签;
            private string _Org采购订单号;

            [ExcelColumn("备注")]
            public string _备注
            {
                get
                {
                    return _Org备注;
                }
                set
                {
                    _Org备注 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("商品SKU码_细")]
            public string SKU
            {
                get
                {
                    return _OrgSku;
                }
                set
                {
                    _OrgSku = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("采购员")]
            public string _采购员
            {
                get
                {
                    return _Org采购员;
                }
                set
                {
                    _Org采购员 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("内部标签")]
            public string _内部标签
            {
                get
                {
                    return _Org内部标签;
                }
                set
                {
                    _Org内部标签 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("采购订单号")]
            public string _采购订单号
            {
                get
                {
                    return _Org采购订单号;
                }
                set
                {
                    _Org采购订单号 = value != null ? value.ToString().Trim() : "";
                }
            }
        }

        class _统计结果
        {
            public string _采购员 { get; set; }
            public int _紧急订单数 { get; set; }
            public int _完成订单数 { get; set; }
            public int _异常订单 { get; set; }
            public int _紧急单SKU个数 { get; set; }
            public int _完成SKU个数 { get; set; }
            public decimal _订单完成占比 { get; set; }
            public decimal _SKU完成占比 { get; set; }
        }
    }
}
