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
    public partial class _库存积压详情 : Form
    {
        public _库存积压详情()
        {
            InitializeComponent();
        }

        private void _库存积压详情_Load(object sender, EventArgs e)
        {
            txt库存周转率.Text = @"C:\Users\Leon\Desktop\source\库存周转率(周转天数大于等于100).csv";
            txt入库明细表.Text = @"C:\Users\Leon\Desktop\source\采购入库明细表.csv|";
        }

        /**************** button event ****************/

        #region 上传库存周转率
        private void btn上传库存周转率_Click(object sender, EventArgs e)
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
                    txt库存周转率.Text = OpenFileDialog1.FileName;
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传入库明细
        private void btn上传入库明细_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Filter = "CSV 文件|*.csv";//设置文件类型
            OpenFileDialog1.Title = "CSV 文件";//设置标题
            OpenFileDialog1.Multiselect = true;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                if (Helper.CheckCSVFileName(OpenFileDialog1.FileName))
                {
                    txt入库明细表.Text = string.Join("|", OpenFileDialog1.FileNames);
                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 处理
        private void btn处理_Click(object sender, EventArgs e)
        {
            var list库存周转率 = new List<_库存周转率Mapping>();
            var list入库明细 = new List<_入库明细Mapping>();
            var list滞销情况汇总 = new List<_滞销情况汇总Model>();
            var list库存积压详情 = new List<_库存积压详情Model>();

            #region 读取数据
            var actReadData = new Action(() =>
            {

                ShowMsg("开始读取库存周转率");
                #region 读取库存周转率
                {
                    var strCSVPath = txt库存周转率.Text;
                    if (!string.IsNullOrEmpty(strCSVPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCSVPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_库存周转率Mapping>()
                                          select c;
                                list库存周转率.AddRange(tmp);

                                for (int idx = list库存周转率.Count - 1; idx >= 0; idx--)
                                {
                                    var curItem = list库存周转率[idx];
                                    if (curItem._可用数量 <= 0 || string.IsNullOrEmpty(curItem._SKU))
                                    {
                                        list库存周转率.RemoveAt(idx);
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

                ShowMsg("开始读取入库明细");
                #region 读取入库明细
                {
                    var strCSVPathArr = !string.IsNullOrEmpty(txt入库明细表.Text) ? txt入库明细表.Text.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries).ToList() : new List<string>();
                    if (strCSVPathArr.Count > 0)
                    {
                        foreach (var strCSVPath in strCSVPathArr)
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
                }
                #endregion
            });
            #endregion


            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                ShowMsg("开始统计滞销汇总");

                #region 统计滞销情况汇总
                {
                    #region _2016年前开发的产品
                    {
                        var refList = list库存周转率.Where(x => x._开发时间 < Convert.ToDateTime("2016-01-01")).ToList();
                        var ref在售 = refList.Where(x => !x._是否停售).ToList();
                        var ref停售 = refList.Where(x => x._是否停售).ToList();
                        var model = new _滞销情况汇总Model();
                        model._类型 = _Enum统计类型._2016年前开发的产品;

                        model._在售SKU个数 = ref在售.Count();
                        model._在售库存金额 = ref在售.Sum(x => x._积压金额);

                        model._停售SKU个数 = ref停售.Count();
                        model._停售库存金额 = ref停售.Sum(x => x._积压金额);
                        list滞销情况汇总.Add(model);
                    }
                    #endregion

                    #region _2016年开发的产品
                    {
                        var refList = list库存周转率.Where(x => x._开发时间 >= Convert.ToDateTime("2016-01-01") && x._开发时间 < Convert.ToDateTime("2017-01-01")).ToList();
                        var ref在售 = refList.Where(x => !x._是否停售).ToList();
                        var ref停售 = refList.Where(x => x._是否停售).ToList();
                        var model = new _滞销情况汇总Model();
                        model._类型 = _Enum统计类型._2016年开发的产品;

                        model._在售SKU个数 = ref在售.Count();
                        model._在售库存金额 = ref在售.Sum(x => x._积压金额);

                        model._停售SKU个数 = ref停售.Count();
                        model._停售库存金额 = ref停售.Sum(x => x._积压金额);
                        list滞销情况汇总.Add(model);
                    }
                    #endregion

                    #region _2017年1_6月份开发的产品
                    {
                        var refList = list库存周转率.Where(x => x._开发时间 >= Convert.ToDateTime("2017-01-01") && x._开发时间 < Convert.ToDateTime("2017-07-01")).ToList();
                        var ref在售 = refList.Where(x => !x._是否停售).ToList();
                        var ref停售 = refList.Where(x => x._是否停售).ToList();
                        var model = new _滞销情况汇总Model();
                        model._类型 = _Enum统计类型._2017年1_6月份开发的产品;

                        model._在售SKU个数 = ref在售.Count();
                        model._在售库存金额 = ref在售.Sum(x => x._积压金额);

                        model._停售SKU个数 = ref停售.Count();
                        model._停售库存金额 = ref停售.Sum(x => x._积压金额);
                        list滞销情况汇总.Add(model);
                    }
                    #endregion

                    #region _2017年7_12月份开发的产品
                    {
                        var refList = list库存周转率.Where(x => x._开发时间 >= Convert.ToDateTime("2017-07-01") && x._开发时间 < Convert.ToDateTime("2018-01-01")).ToList();
                        var ref在售 = refList.Where(x => !x._是否停售).ToList();
                        var ref停售 = refList.Where(x => x._是否停售).ToList();
                        var model = new _滞销情况汇总Model();
                        model._类型 = _Enum统计类型._2017年7_12月份开发的产品;

                        model._在售SKU个数 = ref在售.Count();
                        model._在售库存金额 = ref在售.Sum(x => x._积压金额);

                        model._停售SKU个数 = ref停售.Count();
                        model._停售库存金额 = ref停售.Sum(x => x._积压金额);
                        list滞销情况汇总.Add(model);
                    }
                    #endregion

                    #region _2018年开发的产品
                    {
                        var refList = list库存周转率.Where(x => x._开发时间 >= Convert.ToDateTime("2018-01-01")).ToList();
                        var ref在售 = refList.Where(x => !x._是否停售).ToList();
                        var ref停售 = refList.Where(x => x._是否停售).ToList();
                        var model = new _滞销情况汇总Model();
                        model._类型 = _Enum统计类型._2018年开发的产品;

                        model._在售SKU个数 = ref在售.Count();
                        model._在售库存金额 = ref在售.Sum(x => x._积压金额);

                        model._停售SKU个数 = ref停售.Count();
                        model._停售库存金额 = ref停售.Sum(x => x._积压金额);
                        list滞销情况汇总.Add(model);
                    }
                    #endregion
                }
                #endregion

                #region MyRegion
                { 
                
                }
                #endregion
            }, null);
            #endregion

        }
        #endregion

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_入库明细Mapping), typeof(_库存周转率Mapping));

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

        [ExcelTable("库存周转率")]
        class _库存周转率Mapping
        {
            private string _orgSKU;
            private bool _org是否停售;
            private DateTime? _org开发时间;

            [ExcelColumn("是否停售")]
            public string M是否停售
            {
                set
                {
                    _org是否停售 = value.ToString().IndexOf("是") != -1;
                }
            }

            [ExcelColumn("开发时间")]
            public string M开发时间
            {
                set
                {
                    if (!string.IsNullOrEmpty(value))
                        _org开发时间 = Convert.ToDateTime(value);
                }
            }

            [ExcelColumn("SKU")]
            public string _SKU
            {
                get
                {
                    return _orgSKU;
                }
                set
                {
                    _orgSKU = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("可用数量")]
            public decimal _可用数量 { get; set; }

            [ExcelColumn("成本价")]
            public decimal _成本价 { get; set; }

            public DateTime? _开发时间
            {
                get
                {
                    return _org开发时间;
                }
            }
            public bool _是否停售
            {
                get
                {
                    return _org是否停售;
                }
            }
            public decimal _积压金额
            {
                get
                {
                    return Math.Round(_可用数量 * _成本价, 2);
                }
            }
        }

        [ExcelTable("入库明细")]
        class _入库明细Mapping
        {
            private string _orgSKU;
            private DateTime? _org入库时间;

            [ExcelColumn("商品SKU")]
            public string _SKU
            {
                get
                {
                    return _orgSKU;
                }
                set
                {
                    _orgSKU = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("入库审核时间")]
            public string M入库审核时间
            {
                set
                {
                    if (!string.IsNullOrEmpty(value))
                        _org入库时间 = Convert.ToDateTime(value);
                }
            }
            public DateTime? _入库时间
            {
                get
                {
                    return _org入库时间;
                }
            }
        }

        class _滞销情况汇总Model
        {
            public _Enum统计类型 _类型 { get; set; }
            public int _在售SKU个数 { get; set; }
            public decimal _在售SKU占比 { get; set; }
            public int _停售SKU个数 { get; set; }
            public decimal _停售SKU占比 { get; set; }
            public decimal _在售库存金额 { get; set; }
            public decimal _在售库存金额占比 { get; set; }
            public decimal _停售库存金额 { get; set; }
            public decimal _停售库存金额占比 { get; set; }

        }

        class _库存积压详情Model
        {
            public string _SKU { get; set; }

            public int _积压数量 { get; set; }
            public decimal _单价 { get; set; }
            public decimal _积压总金额
            {
                get
                {
                    return Math.Round((_积压数量 * _单价) * 1.0m, 2);
                }
            }
            public int _积压天数 { get; set; }
        }
        enum _Enum统计类型
        {
            _2016年前开发的产品 = 1,
            _2016年开发的产品 = 2,
            _2017年1_6月份开发的产品 = 3,
            _2017年7_12月份开发的产品 = 4,
            _2018年开发的产品 = 5
        }
    }
}
