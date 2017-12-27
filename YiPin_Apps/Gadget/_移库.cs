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
    public partial class _移库 : Form
    {
        public _移库()
        {
            InitializeComponent();
        }

        private void _移库_Load(object sender, EventArgs e)
        {
            //txt上海在售.Text = @"C:\Users\Leon\Desktop\aa\上海在售.csv";
            //txt昆山在售.Text = @"C:\Users\Leon\Desktop\aa\昆山在售.csv";
            //txt上海停售.Text = @"C:\Users\Leon\Desktop\aa\上海停售.csv";
            //txt昆山停售.Text = @"C:\Users\Leon\Desktop\aa\昆山停售.csv";
            //txt侵权产品.Text = @"C:\Users\Leon\Desktop\aa\侵权.csv";

        }

        /**************** button event ****************/

        #region 上传上海在售
        private void btn上海在售_Click(object sender, EventArgs e)
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
                    txt上海在售.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传昆山在售
        private void btn昆山在售_Click(object sender, EventArgs e)
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
                    txt昆山在售.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传上海停售
        private void btn上海停售_Click(object sender, EventArgs e)
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
                    txt上海停售.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传昆山停售
        private void btn昆山停售_Click(object sender, EventArgs e)
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
                    txt昆山停售.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传侵权产品
        private void btn侵权产品_Click(object sender, EventArgs e)
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
                    txt侵权产品.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 处理按钮事件
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            var str指定仓库 = txt过滤.Text.Trim().ToUpper();

            var _list上海在售 = new List<_上海库位Mapping>();
            var _list上海停售 = new List<_上海库位Mapping>();
            var _list昆山在售 = new List<_昆山库位apping>();
            var _list昆山停售 = new List<_昆山库位apping>();
            var _list侵权产品 = new List<_侵权Mapping>();

            var _list在售 = new List<_在售And停售Model>();
            var _list停售 = new List<_在售And停售Model>();
            var _list侵权 = new List<_在售And停售Model>();

            ShowMsg("开始读取数据");

            #region 读取数据
            var actReadData = new Action(() =>
            {
                #region 读取上海在售
                {
                    var strCsvPath = txt上海在售.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_上海库位Mapping>()
                                          select c;
                                _list上海在售.AddRange(tmp);

                                if (!string.IsNullOrEmpty(str指定仓库))
                                {
                                    for (int idx = _list上海在售.Count - 1; idx >= 0; idx--)
                                    {
                                        if (_list上海在售[idx]._子仓库 != str指定仓库)
                                            _list上海在售.RemoveAt(idx);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 读取上海停售
                {
                    var strCsvPath = txt上海停售.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_上海库位Mapping>()
                                          select c;
                                _list上海停售.AddRange(tmp);

                                if (!string.IsNullOrEmpty(str指定仓库))
                                {
                                    for (int idx = _list上海停售.Count - 1; idx >= 0; idx--)
                                    {
                                        if (_list上海停售[idx]._子仓库 != str指定仓库)
                                            _list上海停售.RemoveAt(idx);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 读取库山在售
                {
                    var strCsvPath = txt昆山在售.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_昆山库位apping>()
                                          select c;
                                _list昆山在售.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 读取库山停售
                {
                    var strCsvPath = txt昆山停售.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_昆山库位apping>()
                                          select c;
                                _list昆山停售.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 读取侵权
                {
                    var strCsvPath = txt侵权产品.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_侵权Mapping>()
                                          select c;
                                _list侵权产品.AddRange(tmp);
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

            ShowMsg("开始计算数据");

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                #region 计算在售
                {
                    /*
                     * 遍历上海在售,把昆山在售的匹配出来
                     */
                    _list上海在售.ForEach(shItem =>
                    {
                        var ref昆山对应记录 = _list昆山在售.Where(x => x.SKU == shItem.SKU).FirstOrDefault();

                        var model = new _在售And停售Model();
                        model._SKU = shItem.SKU;
                        model._商品名称 = shItem._商品名称;
                        model._上海仓库 = shItem._子仓库;
                        model._上海库位 = shItem._原始库位;
                        model._上海货架 = shItem._货架;
                        model._可用数量 = shItem._可用数量;

                        if (ref昆山对应记录 != null)
                        {
                            model._昆山库位 = ref昆山对应记录._原始库位;
                            model._昆山区域 = ref昆山对应记录._区域;
                        }
                        _list在售.Add(model);
                    });

                }
                #endregion

                #region 计算停售
                {
                    _list上海停售.ForEach(shItem =>
                    {
                        //if (shItem.SKU=="DNFK10K38-Y")
                        //{

                        //}


                        /*
                        * 遍历上海停售,把昆山停售的匹配出来
                        */
                        var ref昆山对应记录 = _list昆山停售.Where(x => x.SKU == shItem.SKU).FirstOrDefault();
                        var model = new _在售And停售Model();
                        model._SKU = shItem.SKU;
                        model._商品名称 = shItem._商品名称;
                        model._上海仓库 = shItem._子仓库;
                        model._上海库位 = shItem._原始库位;
                        model._上海货架 = shItem._货架;
                        model._可用数量 = shItem._可用数量;

                        //匹配出来的话,并且有库位,即为库位不为空,要加到在售列表,备注写明:停售
                        if (ref昆山对应记录 != null)
                        {
                            model._昆山库位 = ref昆山对应记录._原始库位;
                            model._昆山区域 = ref昆山对应记录._区域;
                            model._备注 = "停售";
                            if (!string.IsNullOrEmpty(model._昆山库位))
                            {
                                _list在售.Add(model);
                            }
                        }
                        else
                        {
                            //没有匹配出来,看一下是否在侵权sku里面,如果在,放到侵权列表
                            var bQinQuan = _list侵权产品.Where(x => x.SKU == shItem.SKU).Count() > 0;
                            if (bQinQuan)
                            {
                                model._备注 = "停售";
                                _list侵权.Add(model);
                            }
                            else
                            {
                                model._备注 = "侵权";
                                _list停售.Add(model);
                            }
                        }

                    });

                }
                #endregion

                Export(_list在售, _list停售, _list侵权);
            }, null);
            #endregion

        }
        #endregion

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_上海库位Mapping), typeof(_昆山库位apping), typeof(_侵权Mapping));

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

        #region Export 导出表
        private void Export(List<_在售And停售Model> _在售List, List<_在售And停售Model> _停售List, List<_在售And停售Model> _侵权List)
        {
            ShowMsg("开始生成表格");
            var buffer1 = new byte[0];

            #region 生成表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 在售
                {
                    var sheet1 = workbox.Worksheets.Add("在售");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "商品名称";
                    sheet1.Cells[1, 3].Value = "上海库位";
                    sheet1.Cells[1, 4].Value = "上海仓库";
                    sheet1.Cells[1, 5].Value = "上海货架";
                    sheet1.Cells[1, 6].Value = "昆山库位";
                    sheet1.Cells[1, 7].Value = "昆山区域";
                    sheet1.Cells[1, 8].Value = "可用数量";
                    sheet1.Cells[1, 9].Value = "备注";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = _在售List.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = _在售List[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._SKU;
                        sheet1.Cells[rowIdx, 2].Value = info._商品名称;
                        sheet1.Cells[rowIdx, 3].Value = info._上海库位;
                        sheet1.Cells[rowIdx, 4].Value = info._上海仓库;
                        sheet1.Cells[rowIdx, 5].Value = info._上海货架;
                        sheet1.Cells[rowIdx, 6].Value = info._昆山库位;
                        sheet1.Cells[rowIdx, 7].Value = info._昆山区域;
                        sheet1.Cells[rowIdx, 8].Value = info._可用数量;
                        sheet1.Cells[rowIdx, 9].Value = info._备注;
                    }
                    #endregion

                }
                #endregion

                #region 停售
                {
                    var sheet1 = workbox.Worksheets.Add("停售");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "商品名称";
                    sheet1.Cells[1, 3].Value = "上海库位";
                    sheet1.Cells[1, 4].Value = "上海仓库";
                    sheet1.Cells[1, 5].Value = "上海货架";
                    sheet1.Cells[1, 6].Value = "昆山库位";
                    sheet1.Cells[1, 7].Value = "昆山区域";
                    sheet1.Cells[1, 8].Value = "可用数量";
                    sheet1.Cells[1, 9].Value = "备注";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = _停售List.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = _停售List[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._SKU;
                        sheet1.Cells[rowIdx, 2].Value = info._商品名称;
                        sheet1.Cells[rowIdx, 3].Value = info._上海库位;
                        sheet1.Cells[rowIdx, 4].Value = info._上海仓库;
                        sheet1.Cells[rowIdx, 5].Value = info._上海货架;
                        sheet1.Cells[rowIdx, 6].Value = info._昆山库位;
                        sheet1.Cells[rowIdx, 7].Value = info._昆山区域;
                        sheet1.Cells[rowIdx, 8].Value = info._可用数量;
                        sheet1.Cells[rowIdx, 9].Value = info._备注;
                    }
                    #endregion

                }
                #endregion

                #region 侵权
                {
                    var sheet1 = workbox.Worksheets.Add("侵权");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "商品名称";
                    sheet1.Cells[1, 3].Value = "上海库位";
                    sheet1.Cells[1, 4].Value = "上海仓库";
                    sheet1.Cells[1, 5].Value = "上海货架";
                    sheet1.Cells[1, 6].Value = "昆山库位";
                    sheet1.Cells[1, 7].Value = "昆山区域";
                    sheet1.Cells[1, 8].Value = "可用数量";
                    sheet1.Cells[1, 9].Value = "备注";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = _侵权List.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = _侵权List[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._SKU;
                        sheet1.Cells[rowIdx, 2].Value = info._商品名称;
                        sheet1.Cells[rowIdx, 3].Value = info._上海库位;
                        sheet1.Cells[rowIdx, 4].Value = info._上海仓库;
                        sheet1.Cells[rowIdx, 5].Value = info._上海货架;
                        sheet1.Cells[rowIdx, 6].Value = info._昆山库位;
                        sheet1.Cells[rowIdx, 7].Value = info._昆山区域;
                        sheet1.Cells[rowIdx, 8].Value = info._可用数量;
                        sheet1.Cells[rowIdx, 9].Value = info._备注;
                    }
                    #endregion

                }
                #endregion

                buffer1 = package.GetAsByteArray();
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
                    var len = buffer1.Length;
                    using (var fs = File.Create(FileName, len))
                    {
                        fs.Write(buffer1, 0, len);
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

        [ExcelTable("侵权产品表")]
        class _侵权Mapping
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
        }

        [ExcelTable("上海在售/停售表")]
        class _上海库位Mapping
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

            [ExcelColumn("仓库")]
            public string _仓库 { get; set; }

            [ExcelColumn("库位")]
            public string _原始库位
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

            /// <summary>
            /// 详细的仓库,比如上海是一个大仓库,里面包含很多子仓库,分这个概念出来,怕混淆,上海子仓库取库位前2位字符
            /// </summary>
            public string _子仓库
            {
                get
                {
                    var tmp = "";
                    if (!string.IsNullOrEmpty(_原始库位))
                    {
                        tmp = _原始库位.Substring(0, 2);
                    }
                    return tmp;
                }
            }

            /// <summary>
            /// 货架 上海货架取库位信息前3位字符
            /// </summary>
            public string _货架
            {
                get
                {
                    var tmp = "";
                    if (!string.IsNullOrEmpty(_原始库位))
                    {
                        tmp = _原始库位.Substring(0, 3);
                    }
                    return tmp;
                }
            }

        }

        [ExcelTable("昆山在售/停售表")]
        class _昆山库位apping
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

            [ExcelColumn("仓库")]
            public string _仓库 { get; set; }

            [ExcelColumn("库位")]
            public string _原始库位
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

            /// <summary>
            /// 昆山没有子仓库,区域差不多同一个层级
            /// </summary>
            public string _区域
            {
                get
                {
                    var tmp = "";
                    if (!string.IsNullOrEmpty(_原始库位))
                    {
                        tmp = _原始库位.Substring(0, 2);
                    }
                    return tmp;
                }
            }

        }

        class _在售And停售Model
        {
            public string _SKU { get; set; }
            public string _商品名称 { get; set; }
            public string _上海库位 { get; set; }
            public string _上海仓库 { get; set; }
            public string _上海货架 { get; set; }
            public string _昆山库位 { get; set; }
            public string _昆山区域 { get; set; }
            public decimal _可用数量 { get; set; }
            public string _备注 { get; set; }
        }

    }
}
