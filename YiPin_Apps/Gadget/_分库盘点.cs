using LinqToExcel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CommonLibs;
using LinqToExcel.Attributes;

namespace Gadget
{
    public partial class _分库盘点 : Form
    {
        public _分库盘点()
        {
            InitializeComponent();
        }

        private void _分库盘点_Load(object sender, EventArgs e)
        {
            //txt库存明细.Text = @"C:\Users\Leon\Desktop\所有商品信息.csv";
            //txt盘点结果.Text = @"C:\Users\Leon\Desktop\盘点.csv";
        }

        /**************** button event ****************/

        #region 上传库存明细
        private void btn库存明细_Click(object sender, EventArgs e)
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
                    txt库存明细.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传盘点结果
        private void btn盘点结果_Click(object sender, EventArgs e)
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
                    txt盘点结果.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 处理数据
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            var list库存明细 = new List<_库存明细Mapping>();
            var list盘点结果 = new List<_盘点表Mapping>();
            var list导出结果 = new List<_分析结果Model>();
            var i遗留上海天数 = nd遗留上海天数.Value;
            var b只取盘点对应 = cb只要盘点结果.Checked;

            ShowMsg("开始读取表格信息");

            #region 读取数据
            var actReadData = new Action(() =>
            {
                #region 读取库存明细
                {
                    var strCSVPath = txt库存明细.Text;
                    if (!string.IsNullOrEmpty(strCSVPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCSVPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_库存明细Mapping>()
                                          select c;
                                list库存明细.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 读取盘点表
                {
                    var strCSVPath = txt盘点结果.Text;
                    if (!string.IsNullOrEmpty(strCSVPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCSVPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_盘点表Mapping>()
                                          select c;
                                list盘点结果.AddRange(tmp);
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

            ShowMsg("开始处理数据");

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                list库存明细.ForEach(dtItem =>
                {
                    //if (dtItem.SKU=="DNFD12D46-S2")
                    //{
                        
                    //}

                    var model = new _分析结果Model();
                    model._SKU = dtItem.SKU;
                    model._商品名称 = dtItem._商品名称;
                    model._库位 = dtItem._库位;
                    model._可用数量 = dtItem._可用数量;
                    model._日平均销量 = dtItem._平均日销量;
                    model._遗留上海天数 = i遗留上海天数;

                    var ref盘点Item = list盘点结果.Where(x => x.SKU == dtItem.SKU).FirstOrDefault();
                    if (ref盘点Item != null)
                    {
                        model._盘点数量 = ref盘点Item._数量;
                    }

                    if (b只取盘点对应)
                    {
                        if (ref盘点Item != null)
                        {
                            list导出结果.Add(model);
                        }
                    }
                    else
                    {
                        list导出结果.Add(model);
                    }
                });

                Export(list导出结果.OrderBy(x => x._库位).ToList());
            }, null);
            #endregion
        }
        #endregion

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_库存明细Mapping), typeof(_盘点表Mapping));

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

        #region Export 导出结果表格
        private void Export(List<_分析结果Model> list)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];

            #region 数据表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 汇总表
                {
                    var sheet1 = workbox.Worksheets.Add("所有");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "商品名称";
                    sheet1.Cells[1, 3].Value = "库位";
                    sheet1.Cells[1, 4].Value = "仓库";
                    sheet1.Cells[1, 5].Value = "货架";
                    sheet1.Cells[1, 6].Value = "可用数量";
                    sheet1.Cells[1, 7].Value = "盘点数量";
                    sheet1.Cells[1, 8].Value = "留在上海数量";
                    sheet1.Cells[1, 9].Value = "移往昆山数量";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._SKU;
                        sheet1.Cells[rowIdx, 2].Value = info._商品名称;
                        sheet1.Cells[rowIdx, 3].Value = info._库位;
                        sheet1.Cells[rowIdx, 4].Value = info._仓库;
                        sheet1.Cells[rowIdx, 5].Value = info._货架;
                        sheet1.Cells[rowIdx, 6].Value = info._可用数量;
                        sheet1.Cells[rowIdx, 7].Value = info._盘点数量;
                        sheet1.Cells[rowIdx, 8].Value = info._留在上海数量;
                        sheet1.Cells[rowIdx, 9].Value = info._移往昆山数量;
                    }
                    #endregion

                }
                #endregion

                #region 分表
                {
                    var storeNames = list.Select(x => x._仓库).Distinct().ToList();
                    storeNames.ForEach(storeName =>
                    {
                        var sheet1 = workbox.Worksheets.Add(!string.IsNullOrEmpty(storeName) ? storeName : "空仓库");

                        #region 标题行
                        sheet1.Cells[1, 1].Value = "SKU";
                        sheet1.Cells[1, 2].Value = "商品名称";
                        sheet1.Cells[1, 3].Value = "库位";
                        sheet1.Cells[1, 4].Value = "仓库";
                        sheet1.Cells[1, 5].Value = "货架";
                        sheet1.Cells[1, 6].Value = "可用数量";
                        sheet1.Cells[1, 7].Value = "盘点数量";
                        sheet1.Cells[1, 8].Value = "留在上海数量";
                        sheet1.Cells[1, 9].Value = "移往昆山数量";
                        #endregion

                        #region 数据行
                        var refList = list.Where(x => x._仓库 == storeName).ToList();
                        for (int idx = 0, rowIdx = 2, len = refList.Count; idx < len; idx++, rowIdx++)
                        {
                            var info = refList[idx];
                            sheet1.Cells[rowIdx, 1].Value = info._SKU;
                            sheet1.Cells[rowIdx, 2].Value = info._商品名称;
                            sheet1.Cells[rowIdx, 3].Value = info._库位;
                            sheet1.Cells[rowIdx, 4].Value = info._仓库;
                            sheet1.Cells[rowIdx, 5].Value = info._货架;
                            sheet1.Cells[rowIdx, 6].Value = info._可用数量;
                            sheet1.Cells[rowIdx, 7].Value = info._盘点数量;
                            sheet1.Cells[rowIdx, 8].Value = info._留在上海数量;
                            sheet1.Cells[rowIdx, 9].Value = info._移往昆山数量;
                        }
                        #endregion
                    });
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

                    txtExport.Text = FileName;
                    var len = buffer.Length;
                    using (var fs = File.Create(FileName, len))
                    {
                        fs.Write(buffer, 0, len);
                    }
                    ShowMsg("表格生成完毕");
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

        [ExcelTable("库存明细")]
        class _库存明细Mapping
        {
            private string orgSKU;
            private string org库位;
            private string org商品名称;

            [ExcelColumn("SKU码")]
            public string SKU
            {
                get
                {
                    return orgSKU;
                }
                set
                {
                    orgSKU = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("商品名称")]
            public string _商品名称
            {
                get
                {
                    return org商品名称;
                }
                set
                {
                    org商品名称 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("库位")]
            public string _库位
            {
                get
                {
                    return org库位;
                }
                set
                {
                    org库位 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("可用数量")]
            public decimal _可用数量 { get; set; }

            [ExcelColumn("30天销量")]
            public decimal _30天销量 { get; set; }

            [ExcelColumn("15天销量")]
            public decimal _15天销量 { get; set; }

            [ExcelColumn("5天销量")]
            public decimal _5天销量 { get; set; }

            public decimal _平均日销量
            {
                get
                {
                    return Math.Round((_30天销量 / 30 + _15天销量 / 15 + _5天销量 / 5) / 3, 0);
                }
            }

        }

        [ExcelTable("库存明细")]
        class _盘点表Mapping
        {
            private string orgSKU;

            [ExcelColumn("SKU码")]
            public string SKU
            {
                get
                {
                    return orgSKU;
                }
                set
                {
                    orgSKU = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("数量")]
            public decimal _数量 { get; set; }
        }

        class _分析结果Model
        {
            public string _SKU { get; set; }
            public string _商品名称 { get; set; }
            public string _库位 { get; set; }
            public decimal _可用数量 { get; set; }
            public decimal _盘点数量 { get; set; }

            public decimal _遗留上海天数 { get; set; }
            public decimal _日平均销量 { get; set; }

            public string _仓库
            {
                get
                {
                    var tmp = "";
                    if (!string.IsNullOrEmpty(_库位))
                    {
                        tmp = _库位.Substring(0, 2);
                    }
                    return tmp;
                }
            }

            public string _货架
            {
                get
                {
                    var tmp = "";
                    if (!string.IsNullOrEmpty(_库位))
                    {
                        tmp = _库位.Substring(0, 3);
                    }
                    return tmp;
                }
            }

            public decimal _移往昆山数量
            {
                get
                {
                    decimal tmp = 0;
                    if (_盘点数量 != 0)
                    {
                        tmp = _盘点数量 - _留在上海数量;
                    }
                    else
                    {
                        tmp = _可用数量 - _留在上海数量;
                    }
                    return tmp;
                }
            }

            public decimal _留在上海数量
            {
                get
                {
                    return Math.Round(_遗留上海天数 * _日平均销量, 0);
                }
            }
        }


    }
}
