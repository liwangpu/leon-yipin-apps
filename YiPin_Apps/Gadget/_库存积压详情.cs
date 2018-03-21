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
            //txt库存周转率.Text = @"C:\Users\Leon\Desktop\source\库存周转率(周转天数大于等于100).csv";
            //txt入库明细表.Text = @"C:\Users\Leon\Desktop\source\采购入库明细表.csv|";
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
            var list库存积压统计_按类别 = new List<_库存积压统计Model>();
            var list库存积压统计_按采购员 = new List<_库存积压统计Model>();
            var list库存积压统计_按开发 = new List<_库存积压统计Model>();

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
                                    if (curItem._可用数量 > 0 && !string.IsNullOrEmpty(curItem._SKU))
                                    {
                                        var isRemove = false;
                                        if (ndp周转天数.Value > 0)
                                        {
                                            if (curItem._库存周转天数 < ndp周转天数.Value)
                                                isRemove = true;
                                        }

                                        if (ndp库存金额.Value > 0)
                                        {
                                            if (curItem._积压金额 < ndp库存金额.Value)
                                                isRemove = true;
                                        }

                                        if (ndp可用数量.Value > 0)
                                        {
                                            if (curItem._可用数量 < ndp可用数量.Value)
                                                isRemove = true;
                                        }

                                        if (isRemove)
                                            list库存周转率.RemoveAt(idx);
                                    }
                                    else
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
                    //#region _2016年前开发的产品
                    //{
                    //    var refList = list库存周转率.Where(x => x._开发时间 < Convert.ToDateTime("2016-01-01")).ToList();
                    //    var ref在售 = refList.Where(x => !x._是否停售).ToList();
                    //    var ref停售 = refList.Where(x => x._是否停售).ToList();
                    //    var model = new _滞销情况汇总Model();
                    //    model._类型 = _Enum统计类型._2016年前开发的产品;

                    //    model._在售SKU个数 = ref在售.Count();
                    //    model._在售库存金额 = ref在售.Sum(x => x._积压金额);

                    //    model._停售SKU个数 = ref停售.Count();
                    //    model._停售库存金额 = ref停售.Sum(x => x._积压金额);
                    //    list滞销情况汇总.Add(model);
                    //}
                    //#endregion

                    //#region _2016年开发的产品
                    //{
                    //    var refList = list库存周转率.Where(x => x._开发时间 >= Convert.ToDateTime("2016-01-01") && x._开发时间 < Convert.ToDateTime("2017-01-01")).ToList();
                    //    var ref在售 = refList.Where(x => !x._是否停售).ToList();
                    //    var ref停售 = refList.Where(x => x._是否停售).ToList();
                    //    var model = new _滞销情况汇总Model();
                    //    model._类型 = _Enum统计类型._2016年开发的产品;

                    //    model._在售SKU个数 = ref在售.Count();
                    //    model._在售库存金额 = ref在售.Sum(x => x._积压金额);

                    //    model._停售SKU个数 = ref停售.Count();
                    //    model._停售库存金额 = ref停售.Sum(x => x._积压金额);
                    //    list滞销情况汇总.Add(model);
                    //}
                    //#endregion

                    //#region _2017年1_6月份开发的产品
                    //{
                    //    var refList = list库存周转率.Where(x => x._开发时间 >= Convert.ToDateTime("2017-01-01") && x._开发时间 < Convert.ToDateTime("2017-07-01")).ToList();
                    //    var ref在售 = refList.Where(x => !x._是否停售).ToList();
                    //    var ref停售 = refList.Where(x => x._是否停售).ToList();
                    //    var model = new _滞销情况汇总Model();
                    //    model._类型 = _Enum统计类型._2017年1_6月份开发的产品;

                    //    model._在售SKU个数 = ref在售.Count();
                    //    model._在售库存金额 = ref在售.Sum(x => x._积压金额);

                    //    model._停售SKU个数 = ref停售.Count();
                    //    model._停售库存金额 = ref停售.Sum(x => x._积压金额);
                    //    list滞销情况汇总.Add(model);
                    //}
                    //#endregion

                    //#region _2017年7_12月份开发的产品
                    //{
                    //    var refList = list库存周转率.Where(x => x._开发时间 >= Convert.ToDateTime("2017-07-01") && x._开发时间 < Convert.ToDateTime("2018-01-01")).ToList();
                    //    var ref在售 = refList.Where(x => !x._是否停售).ToList();
                    //    var ref停售 = refList.Where(x => x._是否停售).ToList();
                    //    var model = new _滞销情况汇总Model();
                    //    model._类型 = _Enum统计类型._2017年7_12月份开发的产品;

                    //    model._在售SKU个数 = ref在售.Count();
                    //    model._在售库存金额 = ref在售.Sum(x => x._积压金额);

                    //    model._停售SKU个数 = ref停售.Count();
                    //    model._停售库存金额 = ref停售.Sum(x => x._积压金额);
                    //    list滞销情况汇总.Add(model);
                    //}
                    //#endregion

                    //#region _2018年开发的产品
                    //{
                    //    var refList = list库存周转率.Where(x => x._开发时间 >= Convert.ToDateTime("2018-01-01")).ToList();
                    //    var ref在售 = refList.Where(x => !x._是否停售).ToList();
                    //    var ref停售 = refList.Where(x => x._是否停售).ToList();
                    //    var model = new _滞销情况汇总Model();
                    //    model._类型 = _Enum统计类型._2018年开发的产品;

                    //    model._在售SKU个数 = ref在售.Count();
                    //    model._在售库存金额 = ref在售.Sum(x => x._积压金额);

                    //    model._停售SKU个数 = ref停售.Count();
                    //    model._停售库存金额 = ref停售.Sum(x => x._积压金额);
                    //    list滞销情况汇总.Add(model);
                    //}
                    //#endregion
                }
                #endregion

                #region 库存积压统计
                {
                    var tmp = (from it in list库存周转率
                               join cv in list入库明细 on it._SKU equals cv._SKU
                               group cv by it._SKU into r
                               select new
                               {
                                   SKU = r.Key,
                                   LastTime = r.Max(d => d._入库时间)
                               }).Where(x => x.LastTime != null).ToList();

                    var query = (from it in tmp
                                 join cv in list库存周转率 on it.SKU equals cv._SKU
                                 select new _库存积压详情Model()
                                 {
                                     _SKU = cv._SKU,
                                     _积压数量 = cv._可用数量,
                                     _积压总金额 = cv._积压金额,
                                     _积压天数 = (DateTime.Now - (DateTime)it.LastTime).Days,
                                     _采购员 = cv._采购员,
                                     _业绩归属2 = cv._业绩归属2,
                                     _是否停售 = cv._是否停售,
                                     _开发时间 = cv._开发时间
                                 }).Where(x => x._积压天数 >= ndp积压天数.Value && x._采购员 != "李玲玲" && !string.IsNullOrEmpty(x._采购员)).ToList();
                    list库存积压详情.AddRange(query);
                }
                #endregion

                #region 积压总结
                {
                    #region 以类别分
                    {
                        #region 统计在售
                        {
                            var _停售 = false;
                            var model = new _库存积压统计Model();
                            model._统计类别 = _Enum积压类别._在售;
                            ///////
                            model._可用数量_30以下 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压数量 < 30);
                            model._可用数量_30_50 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压数量 >= 30 && x._积压数量 < 50);
                            model._可用数量_50_100 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压数量 >= 50 && x._积压数量 < 100);
                            model._可用数量_100_200 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压数量 >= 100 && x._积压数量 < 200);
                            model._可用数量_200以上 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压数量 >= 200);
                            ///////
                            model._库存金额_30以下 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压总金额 < 30);
                            model._库存金额_30_50 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压总金额 >= 30 && x._积压总金额 < 50);
                            model._库存金额_50_100 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压总金额 >= 50 && x._积压总金额 < 100);
                            model._库存金额_100_200 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压总金额 >= 100 && x._积压总金额 < 200);
                            model._库存金额_200以上 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压总金额 >= 200);
                            ///////
                            model._积压天数_30以下 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压天数 < 30);
                            model._积压天数_30_50 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压天数 >= 30 && x._积压天数 < 50);
                            model._积压天数_50_100 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压天数 >= 50 && x._积压天数 < 100);
                            model._积压天数_100_200 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压天数 >= 100 && x._积压天数 < 200);
                            model._积压天数_200以上 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压天数 >= 200);
                            ///////
                            list库存积压统计_按类别.Add(model);
                        }
                        #endregion

                        #region 统计停售
                        {
                            var _停售 = true;
                            var model = new _库存积压统计Model();
                            model._统计类别 = _Enum积压类别._停售;
                            ///////
                            model._可用数量_30以下 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压数量 < 30);
                            model._可用数量_30_50 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压数量 >= 30 && x._积压数量 < 50);
                            model._可用数量_50_100 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压数量 >= 50 && x._积压数量 < 100);
                            model._可用数量_100_200 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压数量 >= 100 && x._积压数量 < 200);
                            model._可用数量_200以上 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压数量 >= 200);
                            ///////
                            model._库存金额_30以下 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压总金额 < 30);
                            model._库存金额_30_50 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压总金额 >= 30 && x._积压总金额 < 50);
                            model._库存金额_50_100 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压总金额 >= 50 && x._积压总金额 < 100);
                            model._库存金额_100_200 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压总金额 >= 100 && x._积压总金额 < 200);
                            model._库存金额_200以上 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压总金额 >= 200);
                            ///////
                            model._积压天数_30以下 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压天数 < 30);
                            model._积压天数_30_50 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压天数 >= 30 && x._积压天数 < 50);
                            model._积压天数_50_100 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压天数 >= 50 && x._积压天数 < 100);
                            model._积压天数_100_200 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压天数 >= 100 && x._积压天数 < 200);
                            model._积压天数_200以上 = list库存积压详情.Where(x => x._是否停售 == _停售).Count(x => x._积压天数 >= 200);
                            ///////
                            list库存积压统计_按类别.Add(model);
                        }
                        #endregion
                    }
                    #endregion

                    #region 以采购员分
                    {
                        var buyers = list库存积压详情.Select(x => x._采购员).Distinct().ToList();
                        foreach (var byer in buyers)
                        {
                            var model = new _库存积压统计Model();
                            model._采购员 = byer;
                            ///////
                            model._可用数量_30以下 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压数量 < 30);
                            model._可用数量_30_50 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压数量 >= 30 && x._积压数量 < 50);
                            model._可用数量_50_100 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压数量 >= 50 && x._积压数量 < 100);
                            model._可用数量_100_200 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压数量 >= 100 && x._积压数量 < 200);
                            model._可用数量_200以上 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压数量 >= 200);
                            ///////
                            model._库存金额_30以下 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压总金额 < 30);
                            model._库存金额_30_50 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压总金额 >= 30 && x._积压总金额 < 50);
                            model._库存金额_50_100 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压总金额 >= 50 && x._积压总金额 < 100);
                            model._库存金额_100_200 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压总金额 >= 100 && x._积压总金额 < 200);
                            model._库存金额_200以上 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压总金额 >= 200);
                            ///////
                            model._积压天数_30以下 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压天数 < 30);
                            model._积压天数_30_50 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压天数 >= 30 && x._积压天数 < 50);
                            model._积压天数_50_100 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压天数 >= 50 && x._积压天数 < 100);
                            model._积压天数_100_200 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压天数 >= 100 && x._积压天数 < 200);
                            model._积压天数_200以上 = list库存积压详情.Where(x => x._采购员 == byer).Count(x => x._积压天数 >= 200);
                            ///////
                            list库存积压统计_按采购员.Add(model);
                        }
                    }
                    #endregion

                    #region 以开发分
                    {
                        var developer = list库存积压详情.Select(x => x._业绩归属2).Distinct().ToList();
                        foreach (var dvl in developer)
                        {
                            var model = new _库存积压统计Model();
                            model._开发 = dvl;
                            ///////
                            model._可用数量_30以下 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压数量 < 30);
                            model._可用数量_30_50 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压数量 >= 30 && x._积压数量 < 50);
                            model._可用数量_50_100 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压数量 >= 50 && x._积压数量 < 100);
                            model._可用数量_100_200 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压数量 >= 100 && x._积压数量 < 200);
                            model._可用数量_200以上 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压数量 >= 200);
                            ///////
                            model._库存金额_30以下 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压总金额 < 30);
                            model._库存金额_30_50 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压总金额 >= 30 && x._积压总金额 < 50);
                            model._库存金额_50_100 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压总金额 >= 50 && x._积压总金额 < 100);
                            model._库存金额_100_200 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压总金额 >= 100 && x._积压总金额 < 200);
                            model._库存金额_200以上 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压总金额 >= 200);
                            ///////
                            model._积压天数_30以下 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压天数 < 30);
                            model._积压天数_30_50 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压天数 >= 30 && x._积压天数 < 50);
                            model._积压天数_50_100 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压天数 >= 50 && x._积压天数 < 100);
                            model._积压天数_100_200 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压天数 >= 100 && x._积压天数 < 200);
                            model._积压天数_200以上 = list库存积压详情.Where(x => x._业绩归属2 == dvl).Count(x => x._积压天数 >= 200);
                            ///////
                            list库存积压统计_按开发.Add(model);
                        }
                    }
                    #endregion
                }
                #endregion

                Export(list库存积压详情.OrderByDescending(x => x._积压总金额).ToList(), list库存积压统计_按类别, list库存积压统计_按采购员, list库存积压统计_按开发);
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
        private void Export(List<_库存积压详情Model> list积压详情, List<_库存积压统计Model> list积压统计_类别, List<_库存积压统计Model> list积压统计_采购, List<_库存积压统计Model> list积压统计_开发)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];

            #region 库存积压详情
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    var workbox = package.Workbook;

                    #region 详情表
                    {
                        var sheet1 = workbox.Worksheets.Add("库存积压详情");

                        #region 标题行
                        sheet1.Cells[1, 1].Value = "SKU";
                        sheet1.Cells[1, 2].Value = "可用数量";
                        sheet1.Cells[1, 3].Value = "库存金额";
                        sheet1.Cells[1, 4].Value = "积压天数";
                        sheet1.Cells[1, 5].Value = "采购员";
                        sheet1.Cells[1, 6].Value = "业绩归属2";
                        sheet1.Cells[1, 7].Value = "是否停售";
                        sheet1.Cells[1, 8].Value = "开发时间";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 2, len = list积压详情.Count; idx < len; idx++, rowIdx++)
                        {
                            var info = list积压详情[idx];
                            sheet1.Cells[rowIdx, 1].Value = info._SKU;
                            sheet1.Cells[rowIdx, 2].Value = info._积压数量;
                            sheet1.Cells[rowIdx, 3].Value = info._积压总金额;
                            sheet1.Cells[rowIdx, 4].Value = info._积压天数;
                            sheet1.Cells[rowIdx, 5].Value = info._采购员;
                            sheet1.Cells[rowIdx, 6].Value = info._业绩归属2;
                            sheet1.Cells[rowIdx, 7].Value = info._是否停售 ? "是" : "";
                            sheet1.Cells[rowIdx, 8].Value = info._开发时间 != null ? ((DateTime)info._开发时间).ToString("yyyy/MM/dd") : "";
                        }
                        #endregion

                    }
                    #endregion

                    #region 统计表_类别
                    {
                        var sheet1 = workbox.Worksheets.Add("类别划分");

                        #region 标题行
                        sheet1.Cells[2, 1].Value = "类别";
                        sheet1.Cells[2, 2].Value = "30以下";
                        sheet1.Cells[2, 3].Value = "30-50";
                        sheet1.Cells[2, 4].Value = "50-100";
                        sheet1.Cells[2, 5].Value = "100-200";
                        sheet1.Cells[2, 6].Value = "200以上";

                        sheet1.Cells[2, 7].Value = "30以下";
                        sheet1.Cells[2, 8].Value = "30-50";
                        sheet1.Cells[2, 9].Value = "50-100";
                        sheet1.Cells[2, 10].Value = "100-200";
                        sheet1.Cells[2, 11].Value = "200以上";

                        sheet1.Cells[2, 12].Value = "30以下";
                        sheet1.Cells[2, 13].Value = "30-50";
                        sheet1.Cells[2, 14].Value = "50-100";
                        sheet1.Cells[2, 15].Value = "100-200";
                        sheet1.Cells[2, 16].Value = "200以上";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 3, len = list积压统计_类别.Count; idx < len; idx++, rowIdx++)
                        {
                            var info = list积压统计_类别[idx];
                            sheet1.Cells[rowIdx, 1].Value = info._统计类别 == _Enum积压类别._在售 ? "在售" : "停售";
                            sheet1.Cells[rowIdx, 2].Value = info._可用数量_30以下;
                            sheet1.Cells[rowIdx, 3].Value = info._可用数量_30_50;
                            sheet1.Cells[rowIdx, 4].Value = info._可用数量_50_100;
                            sheet1.Cells[rowIdx, 5].Value = info._可用数量_100_200;
                            sheet1.Cells[rowIdx, 6].Value = info._可用数量_200以上;

                            sheet1.Cells[rowIdx, 7].Value = info._库存金额_30以下;
                            sheet1.Cells[rowIdx, 8].Value = info._库存金额_30_50;
                            sheet1.Cells[rowIdx, 9].Value = info._库存金额_50_100;
                            sheet1.Cells[rowIdx, 10].Value = info._库存金额_100_200;
                            sheet1.Cells[rowIdx, 11].Value = info._库存金额_200以上;

                            sheet1.Cells[rowIdx, 12].Value = info._积压天数_30以下;
                            sheet1.Cells[rowIdx, 13].Value = info._积压天数_30_50;
                            sheet1.Cells[rowIdx, 14].Value = info._积压天数_50_100;
                            sheet1.Cells[rowIdx, 15].Value = info._积压天数_100_200;
                            sheet1.Cells[rowIdx, 16].Value = info._积压天数_200以上;
                        }
                        #endregion

                        #region 样式
                        {
                            using (var rng = sheet1.Cells[1, 2, 1, 6])
                            {
                                rng.Value = "可用数量";
                                rng.Merge = true;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            using (var rng = sheet1.Cells[1, 7, 1, 11])
                            {
                                rng.Value = "积压金额";
                                rng.Merge = true;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            using (var rng = sheet1.Cells[1, 12, 1, 16])
                            {
                                rng.Value = "积压天数";
                                rng.Merge = true;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            using (var rng = sheet1.Cells[1, 1, list积压统计_类别.Count + 2, 16])
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

                    #region 统计表_采购员
                    {
                        var sheet1 = workbox.Worksheets.Add("采购员划分");

                        #region 标题行
                        sheet1.Cells[2, 1].Value = "采购员";
                        sheet1.Cells[2, 2].Value = "30以下";
                        sheet1.Cells[2, 3].Value = "30-50";
                        sheet1.Cells[2, 4].Value = "50-100";
                        sheet1.Cells[2, 5].Value = "100-200";
                        sheet1.Cells[2, 6].Value = "200以上";

                        sheet1.Cells[2, 7].Value = "30以下";
                        sheet1.Cells[2, 8].Value = "30-50";
                        sheet1.Cells[2, 9].Value = "50-100";
                        sheet1.Cells[2, 10].Value = "100-200";
                        sheet1.Cells[2, 11].Value = "200以上";

                        sheet1.Cells[2, 12].Value = "30以下";
                        sheet1.Cells[2, 13].Value = "30-50";
                        sheet1.Cells[2, 14].Value = "50-100";
                        sheet1.Cells[2, 15].Value = "100-200";
                        sheet1.Cells[2, 16].Value = "200以上";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 3, len = list积压统计_采购.Count; idx < len; idx++, rowIdx++)
                        {
                            var info = list积压统计_采购[idx];
                            sheet1.Cells[rowIdx, 1].Value = info._采购员;
                            sheet1.Cells[rowIdx, 2].Value = info._可用数量_30以下;
                            sheet1.Cells[rowIdx, 3].Value = info._可用数量_30_50;
                            sheet1.Cells[rowIdx, 4].Value = info._可用数量_50_100;
                            sheet1.Cells[rowIdx, 5].Value = info._可用数量_100_200;
                            sheet1.Cells[rowIdx, 6].Value = info._可用数量_200以上;

                            sheet1.Cells[rowIdx, 7].Value = info._库存金额_30以下;
                            sheet1.Cells[rowIdx, 8].Value = info._库存金额_30_50;
                            sheet1.Cells[rowIdx, 9].Value = info._库存金额_50_100;
                            sheet1.Cells[rowIdx, 10].Value = info._库存金额_100_200;
                            sheet1.Cells[rowIdx, 11].Value = info._库存金额_200以上;

                            sheet1.Cells[rowIdx, 12].Value = info._积压天数_30以下;
                            sheet1.Cells[rowIdx, 13].Value = info._积压天数_30_50;
                            sheet1.Cells[rowIdx, 14].Value = info._积压天数_50_100;
                            sheet1.Cells[rowIdx, 15].Value = info._积压天数_100_200;
                            sheet1.Cells[rowIdx, 16].Value = info._积压天数_200以上;
                        }
                        #endregion

                        #region 样式
                        {
                            using (var rng = sheet1.Cells[1, 2, 1, 6])
                            {
                                rng.Value = "可用数量";
                                rng.Merge = true;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            using (var rng = sheet1.Cells[1, 7, 1, 11])
                            {
                                rng.Value = "积压金额";
                                rng.Merge = true;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            using (var rng = sheet1.Cells[1, 12, 1, 16])
                            {
                                rng.Value = "积压天数";
                                rng.Merge = true;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            using (var rng = sheet1.Cells[1, 1, list积压统计_采购.Count + 2, 16])
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

                    #region 统计表_开发
                    {
                        var sheet1 = workbox.Worksheets.Add("开发划分");

                        #region 标题行
                        sheet1.Cells[2, 1].Value = "开发";
                        sheet1.Cells[2, 2].Value = "30以下";
                        sheet1.Cells[2, 3].Value = "30-50";
                        sheet1.Cells[2, 4].Value = "50-100";
                        sheet1.Cells[2, 5].Value = "100-200";
                        sheet1.Cells[2, 6].Value = "200以上";

                        sheet1.Cells[2, 7].Value = "30以下";
                        sheet1.Cells[2, 8].Value = "30-50";
                        sheet1.Cells[2, 9].Value = "50-100";
                        sheet1.Cells[2, 10].Value = "100-200";
                        sheet1.Cells[2, 11].Value = "200以上";

                        sheet1.Cells[2, 12].Value = "30以下";
                        sheet1.Cells[2, 13].Value = "30-50";
                        sheet1.Cells[2, 14].Value = "50-100";
                        sheet1.Cells[2, 15].Value = "100-200";
                        sheet1.Cells[2, 16].Value = "200以上";
                        #endregion

                        #region 数据行
                        for (int idx = 0, rowIdx = 3, len = list积压统计_开发.Count; idx < len; idx++, rowIdx++)
                        {
                            var info = list积压统计_开发[idx];
                            sheet1.Cells[rowIdx, 1].Value = info._开发;
                            sheet1.Cells[rowIdx, 2].Value = info._可用数量_30以下;
                            sheet1.Cells[rowIdx, 3].Value = info._可用数量_30_50;
                            sheet1.Cells[rowIdx, 4].Value = info._可用数量_50_100;
                            sheet1.Cells[rowIdx, 5].Value = info._可用数量_100_200;
                            sheet1.Cells[rowIdx, 6].Value = info._可用数量_200以上;

                            sheet1.Cells[rowIdx, 7].Value = info._库存金额_30以下;
                            sheet1.Cells[rowIdx, 8].Value = info._库存金额_30_50;
                            sheet1.Cells[rowIdx, 9].Value = info._库存金额_50_100;
                            sheet1.Cells[rowIdx, 10].Value = info._库存金额_100_200;
                            sheet1.Cells[rowIdx, 11].Value = info._库存金额_200以上;

                            sheet1.Cells[rowIdx, 12].Value = info._积压天数_30以下;
                            sheet1.Cells[rowIdx, 13].Value = info._积压天数_30_50;
                            sheet1.Cells[rowIdx, 14].Value = info._积压天数_50_100;
                            sheet1.Cells[rowIdx, 15].Value = info._积压天数_100_200;
                            sheet1.Cells[rowIdx, 16].Value = info._积压天数_200以上;
                        }
                        #endregion

                        #region 样式
                        {
                            using (var rng = sheet1.Cells[1, 2, 1, 6])
                            {
                                rng.Value = "可用数量";
                                rng.Merge = true;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            using (var rng = sheet1.Cells[1, 7, 1, 11])
                            {
                                rng.Value = "积压金额";
                                rng.Merge = true;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            using (var rng = sheet1.Cells[1, 12, 1, 16])
                            {
                                rng.Value = "积压天数";
                                rng.Merge = true;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            using (var rng = sheet1.Cells[1, 1, list积压统计_开发.Count + 2, 16])
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
                    //var notCulcPath = Path.Combine(tmp[0], pureFilName + "(结果).xlsx");
                    //var culcPath = Path.Combine(tmp[0], pureFilName + "(详情).xlsx");

                    var _str库存积压处理结果 = Path.Combine(tmp[0], pureFilName + "(结果).xlsx");
                    var len = buffer.Length;
                    using (var fs = File.Create(_str库存积压处理结果, len))
                    {
                        fs.Write(buffer, 0, len);
                    }

                    //var len1 = buffer1.Length;
                    //if (len1 > 0)
                    //{
                    //    using (var fs = File.Create(culcPath, len1))
                    //    {
                    //        fs.Write(buffer1, 0, len1);
                    //    }
                    //}

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

            [ExcelColumn("库存周转天数")]
            public decimal _库存周转天数 { get; set; }

            [ExcelColumn("成本价")]
            public decimal _成本价 { get; set; }

            [ExcelColumn("采购人")]
            public string _采购员 { get; set; }

            [ExcelColumn("业绩归属人2")]
            public string _业绩归属2 { get; set; }
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
            public decimal _积压数量 { get; set; }
            public decimal _积压总金额 { get; set; }
            public int _积压天数 { get; set; }
            public string _采购员 { get; set; }
            public string _业绩归属2 { get; set; }
            public bool _是否停售 { get; set; }
            public DateTime? _开发时间 { get; set; }
        }

        class _库存积压统计Model
        {
            public _Enum积压类别 _统计类别 { get; set; }
            public string _采购员 { get; set; }
            public string _开发 { get; set; }
            public int _可用数量_30以下 { get; set; }
            public int _可用数量_30_50 { get; set; }
            public int _可用数量_50_100 { get; set; }
            public int _可用数量_100_200 { get; set; }
            public int _可用数量_200以上 { get; set; }

            public int _库存金额_30以下 { get; set; }
            public int _库存金额_30_50 { get; set; }
            public int _库存金额_50_100 { get; set; }
            public int _库存金额_100_200 { get; set; }
            public int _库存金额_200以上 { get; set; }

            public int _积压天数_30以下 { get; set; }
            public int _积压天数_30_50 { get; set; }
            public int _积压天数_50_100 { get; set; }
            public int _积压天数_100_200 { get; set; }
            public int _积压天数_200以上 { get; set; }
        }
        enum _Enum统计类型
        {
            _2016年前开发的产品 = 1,
            _2016年开发的产品 = 2,
            _2017年1_6月份开发的产品 = 3,
            _2017年7_12月份开发的产品 = 4,
            _2018年开发的产品 = 5
        }

        enum _Enum积压类别
        {
            _在售 = 1,
            _停售 = 2
        }
    }
}
