using CommonLibs;
using LinqToExcel;
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
    public partial class _商品信息统计 : Form
    {
        public _商品信息统计()
        {
            InitializeComponent();
        }

        private void _商品信息统计_Load(object sender, EventArgs e)
        {
            //txt商品明细.Text = @"C:\Users\Leon\Desktop\yyy\aaa.csv";
            //txt商品明细.Text = @"C:\Users\Leon\Desktop\yyy\aaa.xlsx";
        }

        /**************** button event ****************/

        #region 上传商品明细事件
        private void btn商品明细_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "CSV 文件|*.csv";//设置文件类型
            OpenFileDialog1.Title = "CSV 文件";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                if (Helper.CheckCSVFileName(OpenFileDialog1.FileName))
                    txt商品明细.Text = OpenFileDialog1.FileName;
                else
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
            }
        }
        #endregion

        #region 分析按钮事件
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            SetButtonState(false);

            var d供应商详情金额下限 = ndLower.Value;
            var _商品明细List = new List<_商品明细Mapping>();
            var _类目统计List = new List<_类目统计Model>();
            var _供应商统计List = new List<_供应商统计Model>();
            var _供应商详情List = new List<_供应商统计Model>();

            #region 读取数据
            var analyzeAct = new Action(() =>
            {
                ShowMsg("开始读取表格数据");
                #region 读取商品明细
                {
                    var strCSVPath = txt商品明细.Text;
                    if (!string.IsNullOrEmpty(strCSVPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCSVPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_商品明细Mapping>()
                                          select c;
                                _商品明细List.AddRange(tmp);
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

            #region 分析数据
            analyzeAct.BeginInvoke((obj) =>
            {
                var list类目列表 = _商品明细List.Where(x => !string.IsNullOrEmpty(x._商品类别)).Select(x => x._商品类别).Distinct().OrderBy(x => x).ToList();
                var list供应商列表 = _商品明细List.Where(x => !string.IsNullOrEmpty(x._供应商)).Select(x => x._供应商).Distinct().OrderBy(x => x).ToList();

                ShowMsg("正在统计类目信息");
                list类目列表.ForEach(className =>
                {
                    var model = new _类目统计Model();
                    model._类目 = className;

                    var refRecords = _商品明细List.Where(x => x._商品类别 == className).ToList();
                    model._供应商个数 = refRecords.Select(x => x._供应商).Distinct().Count();
                    model._SKU个数 = refRecords.Select(x => x._SKU码).Distinct().Count();
                    model._月销量 = refRecords.Sum(x => x._30天销量);
                    model._月销售额 = refRecords.Sum(x => x._30销售金额);
                    _类目统计List.Add(model);
                });

                ShowMsg("正在统计供应商信息");
                list供应商列表.ForEach(providerName =>
                {
                    var model = new _供应商统计Model();
                    model._供应商 = providerName;

                    var refRecords = _商品明细List.Where(x => x._供应商 == providerName).ToList();
                    model._SKU个数 = refRecords.Select(x => x._SKU码).Distinct().Count();
                    model._类目个数 = refRecords.Select(x => x._商品类别).Distinct().Count();
                    model._月销量 = refRecords.Sum(x => x._30天销量);
                    model._月销售额 = refRecords.Sum(x => x._30销售金额);
                    var _采购员List= refRecords.Select(x => x._采购员).Distinct().ToList();
                    var _开发Array = refRecords.Select(x => x._业绩归属2).Distinct().ToArray();
                    var _类目Array = refRecords.Select(x => x._商品类别).Distinct().ToArray();



                    model._采购详细 = _采购员List.Count() > 0 ? string.Join(",", Helper.RemoveUnBuyers(_采购员List).ToArray()) : "";
                    model._开发详细 = _开发Array.Count() > 0 ? string.Join(",", _开发Array) : "";
                    model._类目详细 = _类目Array.Count() > 0 ? string.Join(",", _类目Array) : "";

                    #region 如果月销售额达到,加入 _供应商详情List,并且计算出详细信息
                    {
                        if (model._月销售额 >= d供应商详情金额下限)
                        {
                            var _curRefClassNameList = refRecords.Select(x => x._商品类别).Distinct().ToList();
                            _curRefClassNameList.ForEach(curClass =>
                            {
                                var detail = new _供应商类目详情();
                                detail._类目名称 = curClass;
                                var curClassRefRecord = refRecords.Where(x => x._商品类别 == curClass).ToList();
                                detail._SKU个数 = curClassRefRecord.Select(x => x._SKU码).Distinct().Count();
                                detail._月销量 = curClassRefRecord.Sum(x => x._30天销量);
                                detail._月销售额 = curClassRefRecord.Sum(x => x._30销售金额);
                                model.AddDetail(detail);
                            });
                            _供应商详情List.Add(model);
                        }
                    }
                    #endregion

                    _供应商统计List.Add(model);
                });

                Export(_类目统计List.OrderByDescending(x => x._月销售额).ToList(), _供应商统计List.OrderByDescending(x => x._月销售额).ToList(), _供应商详情List.OrderByDescending(x => x._月销售额).ToList());
            }, null);
            #endregion
        }
        #endregion

        #region 导出表格说明按钮事件
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_商品明细Mapping));

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

        #region 导出表格
        private void Export(List<_类目统计Model> classList, List<_供应商统计Model> providerList, List<_供应商统计Model> detailList)
        {
            ShowMsg("开始生成表格");

            var buffer = new byte[0];

            #region 生成数据表格
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    var workbox = package.Workbook;

                    #region 类目统计表
                    {
                        var sheet1 = workbox.Worksheets.Add("按类目统计");

                        #region 标题行
                        sheet1.Cells[1, 1].Value = "类别";
                        sheet1.Cells[1, 2].Value = "供应商个数";
                        sheet1.Cells[1, 3].Value = "SKU个数";
                        sheet1.Cells[1, 4].Value = "月销量";
                        sheet1.Cells[1, 5].Value = "金额";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 2, len = classList.Count; idx < len; idx++, rowIdx++)
                        {
                            var info = classList[idx];
                            sheet1.Cells[rowIdx, 1].Value = info._类目;
                            sheet1.Cells[rowIdx, 2].Value = info._供应商个数;
                            sheet1.Cells[rowIdx, 3].Value = info._SKU个数;
                            sheet1.Cells[rowIdx, 4].Value = info._月销量;
                            sheet1.Cells[rowIdx, 5].Value = info._月销售额;
                        }
                        #endregion
                    }
                    #endregion

                    #region 供应商统计表
                    {
                        var sheet1 = workbox.Worksheets.Add("按供应商统计");

                        #region 标题行
                        sheet1.Cells[1, 1].Value = "供应商";
                        sheet1.Cells[1, 2].Value = "SKU个数";
                        sheet1.Cells[1, 3].Value = "类别个数";
                        sheet1.Cells[1, 4].Value = "月销量";
                        sheet1.Cells[1, 5].Value = "金额";
                        sheet1.Cells[1, 6].Value = "采购员";
                        sheet1.Cells[1, 7].Value = "开发";
                        sheet1.Cells[1, 8].Value = "类目";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 2, len = providerList.Count; idx < len; idx++, rowIdx++)
                        {
                            var info = providerList[idx];
                            sheet1.Cells[rowIdx, 1].Value = info._供应商;
                            sheet1.Cells[rowIdx, 2].Value = info._SKU个数;
                            sheet1.Cells[rowIdx, 3].Value = info._类目个数;
                            sheet1.Cells[rowIdx, 4].Value = info._月销量;
                            sheet1.Cells[rowIdx, 5].Value = info._月销售额;
                            sheet1.Cells[rowIdx, 6].Value = info._采购详细;
                            sheet1.Cells[rowIdx, 7].Value = info._开发详细;
                            sheet1.Cells[rowIdx, 8].Value = info._类目详细;
                        }
                        #endregion
                    }
                    #endregion

                    #region 供应商详情
                    {
                        var sheet1 = workbox.Worksheets.Add("供应商详情");

                        #region 标题行
                        sheet1.Cells[1, 1].Value = "供应商";
                        sheet1.Cells[1, 2].Value = "类别";
                        sheet1.Cells[1, 3].Value = "SKU个数";
                        sheet1.Cells[1, 4].Value = "月销量";
                        sheet1.Cells[1, 5].Value = "金额";
                        sheet1.Cells[1, 6].Value = "总金额";
                        sheet1.Cells[1, 7].Value = "总SKU个数";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 2, len = detailList.Count; idx < len; idx++)
                        {
                            var info = detailList[idx];
                            var detailCount = info._详细信息.Count();
                            using (var rng = sheet1.Cells[rowIdx, 1, rowIdx + detailCount - 1, 1])
                            {
                                rng.Merge = true;
                                rng.Value = info._供应商;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                                rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            }
                            using (var rng = sheet1.Cells[rowIdx, 6, rowIdx + detailCount - 1, 6])
                            {
                                rng.Merge = true;
                                rng.Value = info._月销售额;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                                rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            }
                            using (var rng = sheet1.Cells[rowIdx, 7, rowIdx + detailCount - 1, 7])
                            {
                                rng.Merge = true;
                                rng.Value = info._SKU个数;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                                rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            }

                            for (int i = 0; i < detailCount; i++)
                            {
                                var dtItem = info._详细信息[i];
                                sheet1.Cells[rowIdx + i, 2].Value = dtItem._类目名称;
                                sheet1.Cells[rowIdx + i, 3].Value = dtItem._SKU个数;
                                sheet1.Cells[rowIdx + i, 4].Value = dtItem._月销量;
                                sheet1.Cells[rowIdx + i, 5].Value = dtItem._月销售额;
                            }
                            rowIdx += detailCount;
                        }
                        #endregion
                    }
                    #endregion

                    buffer = package.GetAsByteArray();
                }
            }
            #endregion

            #region 导出表格
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
                    SetButtonState(true);
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

        #region SetButtonState 刷新按钮状态
        private void SetButtonState(bool bEnable)
        {
            btn商品明细.Enabled = bEnable;
            btnAnalyze.Enabled = bEnable;
        }
        #endregion

        /**************** common class ****************/

        [ExcelTable("商品明细")]
        class _商品明细Mapping
        {
            private string _org商品编码;
            private string _orgSKU码;

            [ExcelColumn("商品编码")]
            public string _商品编码
            {
                get
                {
                    return _org商品编码;
                }
                set
                {
                    _org商品编码 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("SKU码")]
            public string _SKU码
            {
                get
                {
                    return _orgSKU码;
                }
                set
                {
                    _orgSKU码 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("30天销量")]
            public decimal _30天销量 { get; set; }

            [ExcelColumn("15天销量")]
            public decimal _15天销量 { get; set; }

            [ExcelColumn("5天销量")]
            public decimal _5天销量 { get; set; }

            [ExcelColumn("供应商")]
            public string _供应商 { get; set; }

            [ExcelColumn("商品类别")]
            public string _商品类别 { get; set; }

            [ExcelColumn("采购员")]
            public string _采购员 { get; set; }

            [ExcelColumn("业绩归属2")]
            public string _业绩归属2 { get; set; }

            [ExcelColumn("商品成本单价")]
            public decimal _单价 { get; set; }

            public decimal _30销售金额
            {
                get
                {
                    return _单价 * _30天销量;
                }
            }
        }

        class _类目统计Model
        {
            public string _类目 { get; set; }
            public int _供应商个数 { get; set; }
            public int _SKU个数 { get; set; }
            public decimal _月销量 { get; set; }
            public decimal _月销售额 { get; set; }
        }

        class _供应商统计Model
        {
            public _供应商统计Model()
            {
                _详细信息 = new List<_供应商类目详情>();
            }

            public string _供应商 { get; set; }
            public int _SKU个数 { get; set; }
            public int _类目个数 { get; set; }
            public string _类目详细 { get; set; }
            public decimal _月销量 { get; set; }
            public decimal _月销售额 { get; set; }
            public string _采购详细 { get; set; }
            public string _开发详细 { get; set; }

            public List<_供应商类目详情> _详细信息 { get; set; }

            public void AddDetail(_供应商类目详情 detail)
            {
                _详细信息.Add(detail);
            }
        }

        class _供应商类目详情
        {
            public string _类目名称 { get; set; }
            public int _SKU个数 { get; set; }
            public decimal _月销量 { get; set; }
            public decimal _月销售额 { get; set; }
        }
    }
}
