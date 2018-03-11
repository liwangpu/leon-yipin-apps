using CommonLibs;
using LinqToExcel;
using LinqToExcel.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Gadget
{
    public partial class _点货绩效 : Form
    {
        const string _未指定人员 = "未指定人员";
        public _点货绩效()
        {
            InitializeComponent();
        }

        private void _点货绩效_Load(object sender, EventArgs e)
        {
            //txt入库明细.Text = @"C:\Users\Leon\Desktop\点货绩效 - 副本\入库明细.csv";
            //txt人员代号.Text = @"C:\Users\Leon\Desktop\点货绩效 - 副本\人员代号.csv";
            //txt积分参数.Text = @"C:\Users\Leon\Desktop\点货绩效 - 副本\积分参数.csv";
            //txt热销订单.Text = @"C:\Users\Leon\Desktop\点货绩效 - 副本\热销订单.csv";
        }

        /**************** button event ****************/

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_入库明细Mapping), typeof(_人员代号Mapping), typeof(_积分参数Mapping), typeof(_热销订单Mapping));

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

        #region 上传入库明细
        private void btn上传入库明细_Click(object sender, EventArgs e)
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
                    txt入库明细.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传人员代号
        private void btn上传人员代号_Click(object sender, EventArgs e)
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
                    txt人员代号.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传积分参数
        private void btn上传积分参数_Click(object sender, EventArgs e)
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
                    txt积分参数.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传热销订单
        private void btn上传热销订单_Click(object sender, EventArgs e)
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
                    txt热销订单.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 处理数据
        private void btn处理_Click(object sender, EventArgs e)
        {
            var list入库明细 = new List<_入库明细Mapping>();
            var list人员代号 = new List<_人员代号Mapping>();
            var list积分参数 = new List<_积分参数Mapping>();
            var list热销订单 = new List<_热销订单Mapping>();
            var list点货绩效 = new List<_点货绩效Model>();


            #region 读取数据
            var actReadData = new Action(() =>
            {
                ShowMsg("开始读取表格信息");

                #region 读取入库明细
                {
                    var strCSVPath = txt入库明细.Text;
                    if (!string.IsNullOrEmpty(strCSVPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCSVPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_入库明细Mapping>()
                                          select c;
                                list入库明细.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 读取人员代号
                {
                    var strCSVPath = txt人员代号.Text;
                    if (!string.IsNullOrEmpty(strCSVPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCSVPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_人员代号Mapping>()
                                          select c;
                                list人员代号.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 读取库存明细
                {
                    var strCSVPath = txt积分参数.Text;
                    if (!string.IsNullOrEmpty(strCSVPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCSVPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_积分参数Mapping>()
                                          select c;
                                list积分参数.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 读取热销订单
                {
                    var strCSVPath = txt热销订单.Text;
                    if (!string.IsNullOrEmpty(strCSVPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCSVPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_热销订单Mapping>()
                                          select c;
                                list热销订单.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion
            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                ShowMsg("开始处理数据");

                #region 匹配人员信息
                {
                    for (int idx = 0, len = list入库明细.Count; idx < len; idx++)
                    {
                        var item = list入库明细[idx];

                        if (string.IsNullOrEmpty(item._人员代码))
                        {

                        }

                        if (!string.IsNullOrEmpty(item._人员代码))
                        {
                            var ref人员 = list人员代号.Where(x => x._代号 == item._人员代码).FirstOrDefault();
                            if (ref人员 != null)
                                item._人员姓名 = ref人员._姓名;
                            else
                                item._人员姓名 = _未指定人员;
                        }
                    }
                }
                #endregion

                #region 匹配积分
                {
                    for (int idx = 0, len = list入库明细.Count; idx < len; idx++)
                    {
                        var item = list入库明细[idx];
                        var ref积分 = list积分参数.Where(x => x._左区间 < item._总数量 && x._右区间 >= item._总数量).FirstOrDefault();
                        if (ref积分 != null)
                        {
                            item._盘点积分 = ref积分._积分;
                        }

                        item._是否热销订单 = list热销订单.Where(x => x._热销单号 == item._入库单退回单号).Count() > 0;
                    }
                }
                #endregion

                #region 盘点绩效
                {
                    var inventors = list入库明细.Select(x => x._人员姓名).Distinct().ToList();
                    inventors.ForEach(name =>
                    {
                        var refList订单 = list入库明细.Where(x => x._人员姓名 == name).ToList();
                        var model = new _点货绩效Model();
                        model._点货人 = name;
                        model._入库单数 = refList订单.Select(x => x._入库单退回单号).Distinct().Count();
                        model._总积分 = refList订单.Select(x => x._盘点积分).Sum();
                        list点货绩效.Add(model);
                    });
                }
                #endregion

                Export(list点货绩效.OrderByDescending(x => x._总积分).ToList(), list入库明细);
            }, null);
            #endregion
        }
        #endregion

        /**************** common method ****************/

        #region 导出表格
        private void Export(List<_点货绩效Model> resultList, List<_入库明细Mapping> detailList)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer1 = new byte[0];

            #region 绩效结果
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                #region 汇总表
                {
                    var sheet1 = workbox.Worksheets.Add(string.Format("昆山仓{0}月绩效", DateTime.Now.Month));

                    #region 标题行

                    //                    昆山仓2月绩效							绩效工资		
                    //排序	点货人	总积分	工作天数	工作时间	入库单数	每小时平均积分	平均值倍数	主管评分	绩效工资

                    sheet1.Cells[2, 1].Value = "排序";
                    sheet1.Cells[2, 2].Value = "点货人";
                    sheet1.Cells[2, 3].Value = "总积分";
                    sheet1.Cells[2, 4].Value = "工作天数";
                    sheet1.Cells[2, 5].Value = "工作时间";
                    sheet1.Cells[2, 6].Value = "入库单数";
                    sheet1.Cells[2, 7].Value = "每小时平均积分";
                    sheet1.Cells[2, 8].Value = "平均值倍数";
                    sheet1.Cells[2, 9].Value = "主管评分";
                    sheet1.Cells[2, 10].Value = "绩效工资";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 3, len = resultList.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = resultList[idx];
                        sheet1.Cells[rowIdx, 1].Value = idx + 1;
                        sheet1.Cells[rowIdx, 2].Value = info._点货人;
                        sheet1.Cells[rowIdx, 3].Value = info._总积分;
                        sheet1.Cells[rowIdx, 6].Value = info._入库单数;
                    }
                    #endregion

                    #region 表格样式
                    {
                        using (var rng = sheet1.Cells[1, 1, 1, 7])
                        {
                            rng.Value = string.Format("昆山仓{0}月绩效", DateTime.Now.Month);
                            rng.Merge = true;
                        }

                        using (var rng = sheet1.Cells[1, 8, 1, 10])
                        {
                            rng.Value = "绩效工资";
                            rng.Merge = true;
                        }

                        using (var rng = sheet1.Cells[1, 1, 2, 10])
                        {
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#ACB9CA");
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        }

                        using (var rng = sheet1.Cells[1, 1, resultList.Count + 2, 10])
                        {
                            rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        }
                    }
                    #endregion

                }
                #endregion

                buffer = package.GetAsByteArray();
            }
            #endregion

            #region 绩效详情
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 汇总表
                {
                    var sheet1 = workbox.Worksheets.Add("所有");

                    #region 标题行

                    sheet1.Cells[1, 1].Value = "入库单号";
                    sheet1.Cells[1, 2].Value = "点货人";
                    sheet1.Cells[1, 3].Value = "数量";
                    sheet1.Cells[1, 4].Value = "是否热销单";
                    sheet1.Cells[1, 5].Value = "积分";
                    sheet1.Cells[1, 6].Value = "积分减半";
                    sheet1.Cells[1, 7].Value = "修改后积分";

                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = detailList.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = detailList[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._入库单退回单号;
                        sheet1.Cells[rowIdx, 2].Value = info._人员姓名;
                        sheet1.Cells[rowIdx, 3].Value = info._总数量;

                        if (info._是否热销订单)
                        {
                            sheet1.Cells[rowIdx, 4].Value = "是";
                            sheet1.Cells[rowIdx, 6].Value = info._最后积分;
                        }
                        sheet1.Cells[rowIdx, 5].Value = info._盘点积分;
                        sheet1.Cells[rowIdx, 7].Value = info._最后积分;

                    }
                    #endregion

                }
                #endregion

                buffer1 = package.GetAsByteArray();
            }
            #endregion

            #region 导出结果
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
                    var pureFilName = Path.GetFileNameWithoutExtension(FileName);
                    var tmp = FileName.Split(new string[] { pureFilName }, StringSplitOptions.RemoveEmptyEntries);
                    var notCulcPath = Path.Combine(tmp[0], pureFilName + "(结果).xlsx");
                    var culcPath = Path.Combine(tmp[0], pureFilName + "(详情).xlsx");

                    var len = buffer.Length;
                    using (var fs = File.Create(notCulcPath, len))
                    {
                        fs.Write(buffer, 0, len);
                    }

                    var len1 = buffer1.Length;
                    if (len1 > 0)
                    {
                        using (var fs = File.Create(culcPath, len1))
                        {
                            fs.Write(buffer1, 0, len1);
                        }
                    }

                    ShowMsg("表格生成完毕");
                }
            }, null);
            #endregion
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

        [ExcelTable("入库明细")]
        class _入库明细Mapping
        {
            private string org内部便签;
            private string org入库单退回单号;

            [ExcelColumn("内部便签")]
            public string _内部便签
            {
                get
                {
                    return org内部便签;
                }
                set
                {
                    org内部便签 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("入库单/退回单号")]
            public string _入库单退回单号
            {
                get
                {
                    return org入库单退回单号;
                }
                set
                {
                    org入库单退回单号 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("总数量")]
            public decimal _总数量 { get; set; }
            public string _人员代码
            {
                get
                {
                    var tmp = _内部便签.Split(new string[] { ":", "：" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    if (tmp.Count >= 1)
                        return tmp[tmp.Count - 1].Trim();
                    return "";
                }
            }
            public string _人员姓名 { get; set; }
            public decimal _盘点积分 { get; set; }

            public decimal _最后积分
            {
                get
                {
                    if (_是否热销订单)
                        return Math.Round(_盘点积分 / 2, 4);
                    else
                        return _盘点积分;
                }
            }
            public bool _是否热销订单 { get; set; }
        }

        [ExcelTable("人员代号")]
        class _人员代号Mapping
        {
            private string org姓名;
            private string org代号;

            [ExcelColumn("姓名")]
            public string _姓名
            {
                get
                {
                    return org姓名;
                }
                set
                {
                    org姓名 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("代号")]
            public string _代号
            {
                get
                {
                    return org代号;
                }
                set
                {
                    org代号 = value != null ? value.ToString().Trim() : "";
                }
            }
        }

        [ExcelTable("积分参数")]
        class _积分参数Mapping
        {
            [ExcelColumn("左区间")]
            public decimal _左区间 { get; set; }

            [ExcelColumn("右区间")]
            public decimal _右区间 { get; set; }

            [ExcelColumn("积分")]
            public decimal _积分 { get; set; }
        }

        [ExcelTable("热销订单")]
        class _热销订单Mapping
        {
            private string org热销单号;

            [ExcelColumn("热销单号")]
            public string _热销单号
            {
                get
                {
                    return org热销单号;
                }
                set
                {
                    org热销单号 = value != null ? value.ToString().Trim() : "";
                }
            }
        }

        class _点货绩效Model
        {
            public string _点货人 { get; set; }
            public decimal _总积分 { get; set; }
            public int _入库单数 { get; set; }
        }


    }
}
