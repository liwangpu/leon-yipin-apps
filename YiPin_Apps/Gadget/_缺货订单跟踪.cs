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
using Newtonsoft.Json;
namespace Gadget
{
    public partial class _缺货订单跟踪 : Form
    {
        private string _CacheFolder { get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "缓存信息"); } }
        private string _缺货记录信息Path { get { return Path.Combine(_CacheFolder, "缺货记录信息.json"); } }
        private List<_缺货记录信息> _List缺货记录 = new List<_缺货记录信息>();

        public _缺货订单跟踪()
        {
            InitializeComponent();
        }

        private void _缺货订单跟踪_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists(_CacheFolder))
                Directory.CreateDirectory(_CacheFolder);
            if (File.Exists(_缺货记录信息Path))
                using (var fs = new FileStream(_缺货记录信息Path, FileMode.Open))
                using (var reader = new StreamReader(fs))
                {
                    var str = reader.ReadToEnd();
                    _List缺货记录 = JsonConvert.DeserializeObject<List<_缺货记录信息>>(str);
                }


            txt缺货订单.Text = @"C:\Users\Leon\Desktop\缺货订单7月17号\缺货订单7月17号.csv";
            txt在售产品.Text = @"C:\Users\Leon\Desktop\缺货订单7月17号\所有在售产品信息7月17号.csv";
            btn计算.Enabled = true;
        }

        /**************** button event ****************/

        #region 上传缺货订单
        private void btn上传缺货订单_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt缺货订单, () =>
            {
                btn计算.Enabled = CanCalcu();
            });
        }
        #endregion

        #region 上传在售产品
        private void btn上传在售产品_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt在售产品, () =>
            {
                btn计算.Enabled = CanCalcu();
            });
        }
        #endregion

        #region 处理数据
        private void btn计算_Click(object sender, EventArgs e)
        {
            var list缺货订单 = new List<_缺货订单>();
            var list缺货详细信息 = new List<_缺货信息>();
            var list产品信息 = new List<_产品信息>();
            var list当前缺货产品信息 = new List<_产品信息>();
            var list当天已经解决的缺货信息 = new List<_缺货记录信息>();
            var list报表1 = new List<_报表1>();
            var list报表2 = new List<_报表2>();

            #region 读取数据
            var actReadData = new Action(() =>
            {
                var strError = string.Empty;
                ShowMsg("开始读取缺货信息数据");
                FormHelper.ReadCSVFile(txt缺货订单.Text, ref list缺货订单, ref strError);
                ShowMsg("开始读取在售产品数据");
                FormHelper.ReadCSVFile(txt在售产品.Text, ref list产品信息, ref strError);

                list当前缺货产品信息 = list产品信息.Where(x => x._是否缺货 == true).ToList();

                //提取缺货订单数据详细信息(已经过滤不缺货的产品)
                if (list缺货订单.Count > 0)
                {
                    foreach (var item in list缺货订单)
                    {
                        foreach (var dt in item._缺货详情)
                        {
                            //if (dt.SKU== "MRZC2C66-FC")
                            //{

                            //}
                            var refer产品信息 = list产品信息.FirstOrDefault(x => x.SKU == dt.SKU);
                            if (refer产品信息 == null)
                            {
                                dt._已停售 = true;
                                list缺货详细信息.Add(dt);
                            }
                            else
                            {
                                var refer缺货信息 = list当前缺货产品信息.FirstOrDefault(x => x.SKU == dt.SKU);
                                if (refer缺货信息 != null)
                                {
                                    dt._采购员 = refer缺货信息._采购员;
                                    list缺货详细信息.Add(dt);
                                }
                            }
                        }
                    }

                    ////遍历List缺货记录,把今天不缺货的统计出来,同时刷新缺货记录上的最早缺货时间
                    //for (int idx = _List缺货记录.Count - 1; idx >= 0; idx--)
                    //{
                    //    var curData = _List缺货记录[idx];
                    //    var b缺货 = list当前缺货产品信息.Any(x => x.SKU == curData.SKU);
                    //    if (!b缺货)
                    //    {
                    //        //不缺货了,加入当天已经解决的缺货信息
                    //        list当天已经解决的缺货信息.Add(curData);
                    //        //同时清除List缺货记录
                    //        _List缺货记录.RemoveAt(idx);
                    //    }
                    //}

                }

            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                ShowMsg("开始计算数据");

                #region 统计报表1
                {
                    var skus = list缺货详细信息.Select(x => x.SKU).Distinct().ToList();
                    foreach (var sku in skus)
                    {
                        var refers = list缺货详细信息.Where(x => x.SKU == sku).ToList();
                        var defaultData = refers[0];//取一个值出来获取基本信息
                        var model = new _报表1();
                        model.SKU = sku;
                        model._缺货数量 = refers.Sum(x => x._缺货数量);
                        model._缺货订单数量 = refers.Select(x => x._订单编号).Distinct().Count();
                        model._已停售 = defaultData._已停售;
                        model._采购员 = defaultData._采购员;
                        list报表1.Add(model);
                    }
                }
                #endregion

                #region 统计报表2
                {
                    var buyerNames = list报表1.Select(x => x._采购员).Distinct().ToList();
                    var m开发汇总 = new _报表2();
                    m开发汇总._采购员 = "开发";
                    foreach (var name in buyerNames)
                    {
                        var refers = list报表1.Where(x => x._采购员 == name).ToList();
                        if (Helper.IsBuyer(name) || string.IsNullOrWhiteSpace(name))
                        {
                            var model = new _报表2();
                            model._采购员 = !string.IsNullOrWhiteSpace(name) ? name : "停售";
                            model._缺货SKU个数 = refers.Count;
                            model._缺货数量 = refers.Select(x => x._缺货数量).Sum();
                            model._缺货订单数量 = refers.Select(x => x._缺货订单数量).Sum();
                            list报表2.Add(model);
                        }
                        else
                        {
                            m开发汇总._缺货SKU个数 += refers.Count;
                            m开发汇总._缺货数量 += refers.Select(x => x._缺货数量).Sum();
                            m开发汇总._缺货订单数量 += refers.Select(x => x._缺货订单数量).Sum();
                        }
                    }
                    list报表2.Add(m开发汇总);

                }
                #endregion

                ExportExcel(list报表1.OrderByDescending(x => x._采购员).ToList(), list报表2.OrderByDescending(x => x._缺货订单数量).ToList());
            }, null);
            #endregion

        }
        #endregion


        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormHelper.GenerateTableDes(typeof(_缺货订单), typeof(_产品信息));
        }
        #endregion

        /**************** common method ****************/

        #region ExportExcel 导出Excel表格
        /// <summary>
        /// 导出Excel表格
        /// </summary>
        /// <param name="orders"></param>
        private void ExportExcel(List<_报表1> list1, List<_报表2> list2)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];

            #region 报表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 报表1
                {
                    var sheet1 = workbox.Worksheets.Add("缺货详情");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "缺货数量";
                    sheet1.Cells[1, 3].Value = "订单数量";
                    sheet1.Cells[1, 4].Value = "采购员";
                    sheet1.Cells[1, 5].Value = "停售与否";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list1.Count; idx < len; idx++)
                    {
                        var curData = list1[idx];
                        sheet1.Cells[rowIdx, 1].Value = curData.SKU;
                        sheet1.Cells[rowIdx, 2].Value = curData._缺货数量;
                        sheet1.Cells[rowIdx, 3].Value = curData._缺货订单数量;
                        sheet1.Cells[rowIdx, 4].Value = curData._采购员;
                        sheet1.Cells[rowIdx, 5].Value = curData._已停售 ? "停售" : "";
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
                }
                #endregion

                #region 报表2
                {
                    var sheet1 = workbox.Worksheets.Add("缺货统计");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "采购员";
                    sheet1.Cells[1, 2].Value = "缺货SKU数量";
                    sheet1.Cells[1, 3].Value = "缺货数量";
                    sheet1.Cells[1, 4].Value = "缺货订单";
                    sheet1.Cells[1, 5].Value = "整合人员";
                    #endregion



                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list2.Count; idx < len; idx++)
                    {
                        var curData = list2[idx];
                        sheet1.Cells[rowIdx, 1].Value = curData._采购员;
                        sheet1.Cells[rowIdx, 2].Value = curData._缺货SKU个数;
                        sheet1.Cells[rowIdx, 3].Value = curData._缺货数量;
                        sheet1.Cells[rowIdx, 4].Value = curData._缺货订单数量;

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
                    var FileName2 = Path.Combine(savePath, saveFilName + "工作量.xlsx");
                    var FileName3 = Path.Combine(savePath, saveFilName + "(开发订单).xlsx");
                    var FileName4 = Path.Combine(savePath, saveFilName + "(详情表).xlsx");
                    //var FileName5 = Path.Combine(savePath, saveFilName + "(判断两个库存).xlsx");

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
            }, null);
        }
        #endregion

        #region CanCalcu 判断是否可以开始计算数据
        private bool CanCalcu()
        {
            var b上传缺货订单 = !string.IsNullOrWhiteSpace(txt缺货订单.Text);
            var b上传在售产品 = !string.IsNullOrWhiteSpace(txt在售产品.Text);
            return b上传缺货订单 && b上传在售产品;
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

        class _缺货信息
        {
            public string _订单编号 { get; set; }
            public string SKU { get; set; }
            public decimal _缺货数量 { get; set; }
            public DateTime _缺货时间 { get; set; }
            public DateTime _最早缺货时间 { get; set; }
            public bool _已停售 { get; set; }
            public string _采购员 { get; set; }
            public double _延时_天
            {
                get
                {
                    var currentTime = DateTime.Now;
                    return (currentTime - _最早缺货时间).TotalDays;
                }
            }
        }

        [ExcelTable("缺货订单")]
        class _缺货订单
        {
            [ExcelColumn("商品明细")]
            public string _商品明细 { get; set; }
            [ExcelColumn("交易时间(中国)")]
            public string _交易时间 { get; set; }
            [ExcelColumn("订单编号")]
            public string _订单编号 { get; set; }
            public List<_缺货信息> _缺货详情
            {
                get
                {
                    var list = new List<_缺货信息>();
                    if (!string.IsNullOrWhiteSpace(_商品明细))
                    {
                        var item = _商品明细.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                        if (item.Length > 0)
                        {
                            foreach (var str in item)
                            {
                                var arr = str.Split(new string[] { "*" }, StringSplitOptions.RemoveEmptyEntries);
                                if (arr.Length > 1)
                                {
                                    var model = new _缺货信息();
                                    model._订单编号 = _订单编号;
                                    model.SKU = arr[0];
                                    model._缺货时间 = Convert.ToDateTime(_交易时间);
                                    model._缺货数量 = Convert.ToDecimal(arr[1]);
                                    list.Add(model);
                                }
                            }
                        }
                    }
                    return list;
                }
            }

        }

        [ExcelTable("在售产品信息")]
        class _产品信息
        {
            private string _OrgSku;

            [ExcelColumn("SKU码")]
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
            public string _采购员 { get; set; }

            [ExcelColumn("可用数量")]
            public decimal _可用数量 { get; set; }

            [ExcelColumn("缺货及未派单数量")]
            public decimal _缺货及未派单数量 { get; set; }

            [ExcelColumn("缺货占用数量")]
            public decimal _缺货占用数量 { get; set; }

            public bool _是否缺货
            {
                get
                {
                    return _可用数量 - _缺货及未派单数量 + _缺货占用数量 < 0;
                }
            }
        }

        class _报表1
        {
            public string SKU { get; set; }
            public decimal _缺货数量 { get; set; }
            public int _缺货订单数量 { get; set; }
            public string _采购员 { get; set; }
            public bool _已停售 { get; set; }
        }

        class _报表2
        {
            public string _采购员 { get; set; }
            public decimal _缺货SKU个数 { get; set; }
            public decimal _缺货数量 { get; set; }
            public int _缺货订单数量 { get; set; }
        }

        class _缺货记录信息
        {
            public string SKU { get; set; }
            public DateTime _最早缺货时间 { get; set; }
            public double _延时_天
            {
                get
                {
                    var currentTime = DateTime.Now;
                    return (currentTime - _最早缺货时间).TotalDays;
                }
            }
        }

    }
}
