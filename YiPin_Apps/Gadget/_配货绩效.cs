﻿using CommonLibs;
using Gadget.Libs;
using LinqToExcel.Attributes;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Gadget
{
    public partial class _配货绩效 : Form
    {
        private const string BaseCacheFolder = "配货绩效缓存";

        private string CacheBasePath
        {
            get
            {
                return AppDomain.CurrentDomain.BaseDirectory;
            }
        }

        private string Folder人员配置
        {
            get
            {
                return Path.Combine(CacheBasePath, BaseCacheFolder, "人员配置");
            }
        }

        private string Folder公共配置
        {
            get
            {
                return Path.Combine(CacheBasePath, BaseCacheFolder, "公共");
            }
        }

        private string Folder拣货绩效
        {
            get
            {
                return Path.Combine(CacheBasePath, BaseCacheFolder, "拣货绩效");
            }
        }

        private string FileName拣货人员配置缓存文件
        {
            get
            {
                return Path.Combine(Folder人员配置, "库位人员配置信息.json");
            }
        }

        private string FileName本月上班时间缓存文件
        {
            get
            {
                return Path.Combine(Folder公共配置, "本月上班时间信息.json");
            }
        }

        private string FileName帮忙拣货时间缓存文件
        {
            get
            {
                return Path.Combine(Folder公共配置, "帮忙拣货时间.json");
            }
        }

        private string MonthFlag
        {
            get
            {
                var currentTime = dtp绩效时间.Value;
                return currentTime.ToString("yyyy-MM");
            }
        }

        private DateTime CulcTime = new DateTime();

        private List<_拣货人员配置信息> _人员负责库位信息 = new List<_拣货人员配置信息>();
        private List<_本月上班时间信息> _本月上班时间 = new List<_本月上班时间信息>();
        private List<_帮忙点货时间> _本月帮忙拣货时间 = new List<_帮忙点货时间>();

        public _配货绩效()
        {
            InitializeComponent();
        }

        private void _配货绩效_Load(object sender, EventArgs e)
        {

            //txt拣货单.Text = @"C:\Users\Leon\Desktop\配货数量\拣货单数据.csv";
            //txt乱单.Text = @"C:\Users\Leon\Desktop\配货数量\乱单原数据.csv";
            //txt拣货时间.Text = @"C:\Users\Leon\Desktop\配货数量\拣货时间表.csv";
            //btn当天绩效.Enabled = true;

            if (!Directory.Exists(Folder人员配置))
                Directory.CreateDirectory(Folder人员配置);
            if (!Directory.Exists(Folder拣货绩效))
                Directory.CreateDirectory(Folder拣货绩效);
            if (!Directory.Exists(Folder公共配置))
                Directory.CreateDirectory(Folder公共配置);

            //加载缓存人员配置文件
            if (File.Exists(FileName拣货人员配置缓存文件))
            {
                using (var fs = new StreamReader(FileName拣货人员配置缓存文件, Encoding.UTF8))
                {
                    var json = fs.ReadToEnd();
                    _人员负责库位信息 = JsonConvert.DeserializeObject<List<_拣货人员配置信息>>(json);
                }
            }
            //加载本月上班时间信息
            if (File.Exists(FileName本月上班时间缓存文件))
            {
                using (var fs = new StreamReader(FileName本月上班时间缓存文件, Encoding.UTF8))
                {
                    var json = fs.ReadToEnd();
                    _本月上班时间 = JsonConvert.DeserializeObject<List<_本月上班时间信息>>(json);
                }
            }
            //加载本月帮忙拣货时间信息
            if (File.Exists(FileName帮忙拣货时间缓存文件))
            {
                using (var fs = new StreamReader(FileName帮忙拣货时间缓存文件, Encoding.UTF8))
                {
                    var json = fs.ReadToEnd();
                    _本月帮忙拣货时间 = JsonConvert.DeserializeObject<List<_帮忙点货时间>>(json);
                }
            }
            RefreshCache();
        }

        /**************** button event ****************/

        #region 上传拣货单
        private void btn上传拣货单_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt拣货单);
        }
        #endregion

        #region 上传乱单
        private void btn上传乱单_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt乱单);
        }
        #endregion

        #region 缓存帮忙点货
        private void btn缓存帮忙拣货时间_Click(object sender, EventArgs e)
        {

            FormHelper.GetCSVPath(txt帮忙点货, (Action)(() =>
            {
                _本月帮忙拣货时间.Clear();
                var strError = string.Empty;
                ShowMsg("开始读取帮忙拣货时间数据");

                #region 读取数据
                var actReadData = new Action(() =>
                {
                    FormHelper.ReadCSVFile(txt帮忙点货.Text, ref _本月帮忙拣货时间, ref strError);
                });
                #endregion

                #region 处理数据
                actReadData.BeginInvoke((AsyncCallback)((obj) =>
                {
                    ShowMsg("帮忙拣货时间数据读取完毕");
                    if (_本月帮忙拣货时间 != null && _本月帮忙拣货时间.Count > 0)
                    {
                        var json = JsonConvert.SerializeObject(_本月帮忙拣货时间);
                        using (var fs = new StreamWriter(FileName帮忙拣货时间缓存文件, false, Encoding.UTF8))
                        {
                            fs.Write(json);
                        }
                        MessageBox.Show("帮忙拣货时间数据存储完毕", "温馨提示");
                    }
                    ShowMsg(strError);
                }), null);
                #endregion
            }));
        }
        #endregion

        #region 缓存拣货人员配置
        private void btn库位人员配置_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt拣货人员配置, () =>
             {
                 var list拣货人员配置 = new List<_拣货人员配置>();
                 var strError = string.Empty;
                 ShowMsg("开始读取拣货人员配置数据");

                 #region 读取数据
                 var actReadData = new Action(() =>
                 {
                     FormHelper.ReadCSVFile(txt拣货人员配置.Text, ref list拣货人员配置, ref strError);
                 });
                 #endregion

                 #region 处理数据
                 actReadData.BeginInvoke((obj) =>
                 {
                     ShowMsg("拣货人员配置数据读取完毕");
                     if (list拣货人员配置 != null && list拣货人员配置.Count > 0)
                     {
                         _人员负责库位信息.Clear();
                         list拣货人员配置.ForEach(x =>
                         {
                             var set = new _拣货人员配置信息();
                             set._姓名 = x._配货人员;
                             set.管理库位 = x._库位;
                             _人员负责库位信息.Add(set);
                         });
                         //var allEmpName = list拣货人员配置.Select(x => x._配货人员).Distinct().ToList();
                         //allEmpName.ForEach(name =>
                         //{
                         //    var md = new _拣货人员配置信息();
                         //    md._姓名 = name;
                         //    md.管理库位 = list拣货人员配置.Where(x => x._配货人员 == name).Select(x => x._库位).ToList();
                         //    _人员负责库位信息.Add(md);
                         //});
                         Cache拣货人员配置();

                         MessageBox.Show("拣货人员配置数据存储完毕", "温馨提示");
                     }
                     ShowMsg(strError);
                 }, null);
                 #endregion
             });
        }
        #endregion

        #region 缓存本月上班时间
        private void btn缓存拣货时间_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt本月上班时间, (Action)(() =>
            {
                _本月上班时间.Clear();
                var strError = string.Empty;
                ShowMsg("开始读取上班时间数据");

                #region 读取数据
                var actReadData = new Action(() =>
                {
                    FormHelper.ReadCSVFile(txt本月上班时间.Text, ref _本月上班时间, ref strError);
                });
                #endregion

                #region 处理数据
                actReadData.BeginInvoke((AsyncCallback)((obj) =>
                {
                    ShowMsg("上班时间数据读取完毕");
                    if (_本月上班时间 != null && _本月上班时间.Count > 0)
                    {
                        var json = JsonConvert.SerializeObject(_本月上班时间);
                        using (var fs = new StreamWriter(FileName本月上班时间缓存文件, false, Encoding.UTF8))
                        {
                            fs.Write(json);
                        }

                        MessageBox.Show("上班时间数据存储完毕", "温馨提示");
                    }
                    ShowMsg(strError);
                }), null);
                #endregion
            }));
        }
        #endregion

        #region 计算当天绩效
        private void btn当天绩效_Click(object sender, EventArgs e)
        {
            if (_人员负责库位信息 != null && _人员负责库位信息.Count > 0)
            {
                try
                {
                    double _d张数定值 = Convert.ToDouble(nup张数定值.Value);
                    double _d张数占比 = Convert.ToDouble(nup张数占比.Value);
                    double _d数量定值 = Convert.ToDouble(nup数量定值.Value);
                    double _d数量占比 = Convert.ToDouble(nup数量占比.Value);

                    CulcTime = dtp绩效时间.Value;
                    btn当天绩效.Enabled = false;
                    btn全月绩效.Enabled = false;
                    var strError = string.Empty;
                    var list拣货单 = new List<_拣货单>();
                    var list乱单 = new List<_乱单>();
                    var list最终绩效 = new List<_配货绩效结果>();
                    #region 读取数据
                    var actReadData = new Action(() =>
                    {
                        ShowMsg("开始读取当天绩效相关信息");
                        FormHelper.ReadCSVFile(txt拣货单.Text, ref list拣货单, ref strError);
                        FormHelper.ReadCSVFile(txt乱单.Text, ref list乱单, ref strError);

                        //将乱单转换正常拣货单
                        foreach (var item乱单 in list乱单)
                        {
                            //var aaa = item乱单.ToData();
                            list拣货单.AddRange(item乱单.ToData());
                        }
                    });
                    #endregion

                    #region 处理数据
                    actReadData.BeginInvoke((obj) =>
                    {
                        ShowMsg("绩效数据读取完毕,即将开始计算");

                        if (list拣货单.Count > 0)
                        {
                            var allEmpNames = _人员负责库位信息.Select(x => x._姓名).Distinct().ToList();
                            allEmpNames.ForEach(name =>
                            {
                                if (!string.IsNullOrEmpty(name))
                                {

                                    //if (name.Trim()== "魏婷")
                                    //{

                                    //}
                                    var md = new _配货绩效结果();
                                    md._d张数占比 = _d张数占比;
                                    md._d张数定值 = _d张数定值;
                                    md._d数量定值 = _d数量定值;
                                    md._d数量占比 = _d数量占比;


                                    md._业绩归属人 = name;
                                    var _订单详情数据 = new List<_订单详情数据>();

                                    #region 抽取详细信息
                                    {
                                        //var aarrr = (from it in list拣货单
                                        //             join s in _人员负责库位信息 on it._库位号 equals s.管理库位
                                        //             where s._姓名 == name && it._乱单 == true
                                        //             select it).ToList();
                                        //if (aarrr.Count > 0)
                                        //{

                                        //}

                                        var refLh = (from it in list拣货单
                                                     join s in _人员负责库位信息 on it._库位号 equals s.管理库位
                                                     where s._姓名 == name
                                                     select it).ToList();
                                        foreach (var deitem in refLh)
                                        {
                                            var item = deitem._拣货明细;
                                            foreach (var it in item)
                                            {
                                                var arr = it.Split(new string[] { "*" }, StringSplitOptions.RemoveEmptyEntries);
                                                if (arr.Length >= 2)
                                                {
                                                    var detail = new _订单详情数据();
                                                    detail.SKU = arr[0].Trim();
                                                    detail.Amount = Convert.ToDouble(arr[1]);
                                                    detail._乱单 = deitem._乱单;
                                                    _订单详情数据.Add(detail);
                                                }

                                            }
                                        }

                                        //foreach (List<string> item in refLh)
                                        //{
                                        //    foreach (var it in item)
                                        //    {
                                        //        var arr = it.Split(new string[] { "*" }, StringSplitOptions.RemoveEmptyEntries);
                                        //        if (arr.Length >= 2)
                                        //        {
                                        //            var detail = new _订单详情数据();
                                        //            detail.SKU = arr[0].Trim();
                                        //            detail.Amount = Convert.ToDouble(arr[1]);
                                        //            //detail._乱单=
                                        //            _订单详情数据.Add(detail);
                                        //        }

                                        //    }
                                        //}
                                    }
                                    #endregion

                                    var list订单详情数据_拣货单 = _订单详情数据.Where(x => x._乱单 == false).ToList();
                                    var list订单详情数据_乱单 = _订单详情数据.Where(x => x._乱单 == true).ToList();

                                    var str_帮忙总时长 = "";
                                    var refTime = calc计算上班时间(CulcTime, name, ref str_帮忙总时长);
                                    md._拣货单张数_正常 = list订单详情数据_拣货单.Select(x => x.SKU).Distinct().Count();
                                    md._购买总数量_正常 = list订单详情数据_拣货单.Select(x => x.Amount).Sum();
                                    md._拣货单张数_乱单 = list订单详情数据_乱单.Select(x => x.SKU).Distinct().Count();
                                    md._购买总数量_乱单 = list订单详情数据_乱单.Select(x => x.Amount).Sum();


                                    //md._拣货单张数 = _订单详情数据.Select(x => x.SKU).Distinct().Count();
                                    //md._购买总数量 = _订单详情数据.Select(x => x.Amount).Sum();
                                    //if (refTime != null)
                                    //{
                                    //refTime.CulcTime = CulcTime;
                                    //var mm = refTime._拣货总时间.TotalMinutes % 60;
                                    //var hh = (refTime._拣货总时间.TotalMinutes - mm) / 60;
                                    md._总时长 = refTime.ToString();
                                    md._帮忙总时长 = str_帮忙总时长;
                                    md._分钟 = Convert.ToDouble(refTime * 60);
                                    //}

                                    list最终绩效.Add(md);
                                }
                            });

                            if (list最终绩效.Count > 0)
                            {
                                ShowMsg("开始存储当天绩效");
                                Cache当天绩效(list最终绩效);
                                ShowMsg("当天绩效存储完毕");
                                ExportExcel(list最终绩效);
                            }
                        }
                        ShowMsg(strError);
                    }, null);
                    #endregion
                }
                catch (Exception ex)
                {
                    ShowMsg(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("库位人员配置为空,请先上传库位人员配置", "温馨提示");
            }
        }
        #endregion

        #region 计算全月绩效
        private void btn全月绩效_Click(object sender, EventArgs e)
        {

            var folder = Path.Combine(Folder拣货绩效, MonthFlag);
            var cacheFiles = Directory.EnumerateFiles(Path.Combine(Folder拣货绩效, MonthFlag));
            if (cacheFiles != null && cacheFiles.Count() > 0)
            {
                btn全月绩效.Enabled = false;
                var list = new List<_配货绩效结果>();
                foreach (var item in cacheFiles)
                {
                    using (var fs = new StreamReader(item, Encoding.UTF8))
                    {
                        var json = fs.ReadToEnd();
                        list.AddRange(JsonConvert.DeserializeObject<List<_配货绩效结果>>(json));
                    }
                }
                var datas = new List<_配货绩效结果>();
                var expNames = list.Select(x => x._业绩归属人).Distinct().ToList();
                expNames.ForEach(n =>
                {
                    if (!string.IsNullOrWhiteSpace(n))
                    {
                        var md = new _配货绩效结果();
                        md._业绩归属人 = n;

                        var defaultItem = list[0];
                        md._d张数占比 = defaultItem._d张数占比;
                        md._d张数定值 = defaultItem._d张数定值;
                        md._d数量占比 = defaultItem._d数量占比;
                        md._d数量定值 = defaultItem._d数量定值;


                        md._购买总数量_正常 = list.Where(x => x._业绩归属人 == n).Select(x => x._购买总数量_正常).Sum();
                        md._拣货单张数_正常 = list.Where(x => x._业绩归属人 == n).Select(x => x._拣货单张数_正常).Sum();
                        md._购买总数量_乱单 = list.Where(x => x._业绩归属人 == n).Select(x => x._购买总数量_乱单).Sum();
                        md._拣货单张数_乱单 = list.Where(x => x._业绩归属人 == n).Select(x => x._拣货单张数_乱单).Sum();

                        md._分钟 = list.Where(x => x._业绩归属人 == n).Select(x => x._分钟).Sum();
                        md._总时长 = list.Where(x => x._业绩归属人 == n).Select(x => x._小时).Sum().ToString();
                        md._帮忙总时长 = list.Where(x => x._业绩归属人 == n).Select(x => Convert.ToDecimal(x._帮忙总时长)).Sum().ToString();
                        //var mm = md._分钟 % 60;
                        //var hh = (md._分钟 - mm) / 60;
                        //md._总时长 = string.Format("{0}:{1}:00", hh > 9 ? "" + hh : "0" + hh, mm > 9 ? "" + mm : "0" + mm);
                        datas.Add(md);
                    }
                });
                ExportExcel(datas);
                btn全月绩效.Enabled = true;
            }
        }
        #endregion

        #region 导出历史绩效
        private void 导出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var fn = lsbCache.SelectedItem.ToString();
            var fileName = Path.Combine(Folder拣货绩效, MonthFlag, fn);
            if (File.Exists(fileName))
            {
                using (var fs = new StreamReader(fileName, Encoding.UTF8))
                {
                    var json = fs.ReadToEnd();
                    var list = JsonConvert.DeserializeObject<List<_配货绩效结果>>(json);
                    ExportExcel(list);
                }
            }
        }
        #endregion

        #region 导出表格说明事件
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormHelper.GenerateTableDes(typeof(_拣货单), typeof(_乱单), typeof(_帮忙点货时间), typeof(_拣货人员配置), typeof(_本月上班时间信息));
        }
        #endregion

        #region 刷新缓存信息
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            RefreshCache();
        }
        #endregion

        /**************** common method ****************/

        #region 计算绩效日期上班和帮忙时间
        private decimal calc计算上班时间(DateTime d绩效日期, string str姓名, ref string str帮忙时间)
        {
            //if (str姓名 == "顾洁")
            //{

            //}
            d绩效日期 = d绩效日期.Date;
            if (_本月上班时间 != null && _本月上班时间.Count > 0)
            {
                var refer工作时间 = _本月上班时间.Where(x => x._姓名 == str姓名).FirstOrDefault();
                var refer帮忙时间 = _本月帮忙拣货时间.Where(x => x._姓名 == str姓名 && x._日期 == d绩效日期).FirstOrDefault();
                if (refer工作时间 != null)
                {
                    decimal d上班时间 = 0;
                    decimal d帮忙时间 = 0;
                    if (refer帮忙时间 != null && refer帮忙时间._帮忙总时间 != null)
                    {
                        var h = (refer帮忙时间._帮忙总时间).Hours;
                        var mh = Math.Round((refer帮忙时间._帮忙总时间).Minutes / 60m, 1);
                        d帮忙时间 = h + mh;
                    }
                    str帮忙时间 = d帮忙时间.ToString();
                    switch (d绩效日期.Day)
                    {
                        case 1:
                            d上班时间 = refer工作时间._1号;
                            break;
                        case 2:
                            d上班时间 = refer工作时间._2号;
                            break;
                        case 3:
                            d上班时间 = refer工作时间._3号;
                            break;
                        case 4:
                            d上班时间 = refer工作时间._4号;
                            break;
                        case 5:
                            d上班时间 = refer工作时间._5号;
                            break;
                        case 6:
                            d上班时间 = refer工作时间._6号;
                            break;
                        case 7:
                            d上班时间 = refer工作时间._7号;
                            break;
                        case 8:
                            d上班时间 = refer工作时间._8号;
                            break;
                        case 9:
                            d上班时间 = refer工作时间._9号;
                            break;
                        case 10:
                            d上班时间 = refer工作时间._10号;
                            break;
                        case 11:
                            d上班时间 = refer工作时间._11号;
                            break;
                        case 12:
                            d上班时间 = refer工作时间._12号;
                            break;
                        case 13:
                            d上班时间 = refer工作时间._13号;
                            break;
                        case 14:
                            d上班时间 = refer工作时间._14号;
                            break;
                        case 15:
                            d上班时间 = refer工作时间._15号;
                            break;
                        case 16:
                            d上班时间 = refer工作时间._16号;
                            break;
                        case 17:
                            d上班时间 = refer工作时间._17号;
                            break;
                        case 18:
                            d上班时间 = refer工作时间._18号;
                            break;
                        case 19:
                            d上班时间 = refer工作时间._19号;
                            break;
                        case 20:
                            d上班时间 = refer工作时间._20号;
                            break;
                        case 21:
                            d上班时间 = refer工作时间._21号;
                            break;
                        case 22:
                            d上班时间 = refer工作时间._22号;
                            break;
                        case 23:
                            d上班时间 = refer工作时间._23号;
                            break;
                        case 24:
                            d上班时间 = refer工作时间._24号;
                            break;
                        case 25:
                            d上班时间 = refer工作时间._25号;
                            break;
                        case 26:
                            d上班时间 = refer工作时间._26号;
                            break;
                        case 27:
                            d上班时间 = refer工作时间._27号;
                            break;
                        case 28:
                            d上班时间 = refer工作时间._28号;
                            break;
                        case 29:
                            d上班时间 = refer工作时间._29号;
                            break;
                        case 30:
                            d上班时间 = refer工作时间._30号;
                            break;
                        case 31:
                            d上班时间 = refer工作时间._31号;
                            break;
                        default:
                            break;
                    }
                    if (d上班时间 <= 0)
                        return 0;
                    return d上班时间 - d帮忙时间;
                }
            }
            return 0;
        }
        #endregion

        #region 导出表格
        private void ExportExcel(List<_配货绩效结果> list拣货绩效)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];

            #region 订单分配
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");


                #region 标题行
                sheet1.Cells[1, 1].Value = "姓名";
                sheet1.Cells[1, 2].Value = "拣货单数量";
                sheet1.Cells[1, 3].Value = "乱单数量";
                sheet1.Cells[1, 4].Value = "总数量";
                sheet1.Cells[1, 5].Value = "拣货单张数";
                sheet1.Cells[1, 6].Value = "乱单张数";
                sheet1.Cells[1, 7].Value = "总张数";
                sheet1.Cells[1, 8].Value = "帮忙总时长";
                sheet1.Cells[1, 9].Value = "工作总时长";
                sheet1.Cells[1, 10].Value = "分钟";
                sheet1.Cells[1, 11].Value = "拣货单效率";
                sheet1.Cells[1, 12].Value = "购买数量效率";
                sheet1.Cells[1, 13].Value = "小时";
                sheet1.Cells[1, 14].Value = "拣货单每小时";
                sheet1.Cells[1, 15].Value = "个数每小时";
                sheet1.Cells[1, 16].Value = "定值倍数";
                sheet1.Cells[1, 17].Value = "工资";

                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = list拣货绩效.Count; idx < len; idx++)
                {
                    var curOrder = list拣货绩效[idx];
                    sheet1.Cells[rowIdx, 1].Value = curOrder._业绩归属人;
                    sheet1.Cells[rowIdx, 2].Value = curOrder._购买总数量_正常;
                    sheet1.Cells[rowIdx, 3].Value = curOrder._购买总数量_乱单;
                    sheet1.Cells[rowIdx, 4].Value = curOrder._购买总数量;
                    sheet1.Cells[rowIdx, 5].Value = curOrder._拣货单张数_正常;
                    sheet1.Cells[rowIdx, 6].Value = curOrder._拣货单张数_乱单;
                    sheet1.Cells[rowIdx, 7].Value = curOrder._拣货单张数;
                    sheet1.Cells[rowIdx, 8].Value = curOrder._帮忙总时长;
                    sheet1.Cells[rowIdx, 9].Value = curOrder._总时长;
                    sheet1.Cells[rowIdx, 10].Value = curOrder._分钟;
                    sheet1.Cells[rowIdx, 11].Value = curOrder._拣货单效率;
                    sheet1.Cells[rowIdx, 12].Value = curOrder._购买数量效率;
                    sheet1.Cells[rowIdx, 13].Value = curOrder._小时;
                    sheet1.Cells[rowIdx, 14].Value = curOrder._拣货单每小时;
                    sheet1.Cells[rowIdx, 15].Value = curOrder._个数每小时;
                    sheet1.Cells[rowIdx, 16].Value = curOrder._定值倍数;
                    sheet1.Cells[rowIdx, 17].Value = curOrder._工资;
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
                    RefreshCache();
                }
                btn当天绩效.Enabled = true;
                btn全月绩效.Enabled = true;
            }, null);
        }
        #endregion

        #region 存储拣货人员配置
        private void Cache拣货人员配置()
        {
            var json = JsonConvert.SerializeObject(_人员负责库位信息);
            using (var fs = new StreamWriter(FileName拣货人员配置缓存文件, false, Encoding.UTF8))
            {
                fs.Write(json);
            }
        }
        #endregion

        #region 存储当天绩效
        private void Cache当天绩效(List<_配货绩效结果> result)
        {
            var ct = CulcTime;
            var dateFlag = string.Format("{0}-{1}", ct.Month, ct.Day);
            var json = JsonConvert.SerializeObject(result);
            var folder = Path.Combine(Folder拣货绩效, MonthFlag);
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            var fileName = Path.Combine(Folder拣货绩效, MonthFlag, dateFlag + ".json");
            using (var fs = new StreamWriter(fileName, false, Encoding.UTF8))
            {
                fs.Write(json);
            }
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

        #region 刷新缓存信息
        private void RefreshCache()
        {
            var list = new List<string>();
            var folder = Path.Combine(Folder拣货绩效, MonthFlag);
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            var cacheFiles = Directory.EnumerateFiles(folder);
            if (cacheFiles != null && cacheFiles.Count() > 0)
            {
                foreach (var item in cacheFiles)
                {
                    var fn = Path.GetFileName(item);
                    list.Add(fn);
                }
                lsbCache.DataSource = list;
            }
        }
        #endregion

        /**************** common class ****************/

        [ExcelTable("拣货单")]
        class _拣货单
        {
            private string _Org商品明细;
            private string _Org完整库位号;

            [ExcelColumn("商品明细")]
            public string _商品明细
            {
                get
                {
                    return _Org商品明细;
                }
                set
                {
                    _Org商品明细 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("库位号")]
            public string _完整库位号
            {
                get
                {
                    return _Org完整库位号;
                }
                set
                {
                    _Org完整库位号 = value != null ? value.ToString().Trim() : "";
                }
            }

            public string _库位号
            {
                get
                {
                    var arr = _完整库位号.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                    if (arr.Count() > 0)
                    {
                        var first = arr[0];
                        if (first.Length < 3)
                            return string.Empty;
                        return first.Substring(0, 3).ToUpper();
                    }
                    return "";
                }
            }

            public List<string> _拣货明细
            {
                get
                {
                    var list = new List<string>();
                    if (!string.IsNullOrEmpty(_商品明细))
                    {
                        var arr = _商品明细.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                        if (arr.Count() > 0)
                            list.AddRange(arr.ToList());
                    }
                    return list;
                }
            }

            public bool _乱单 { get; set; }

        }

        [ExcelTable("_乱单")]
        class _乱单
        {
            private string _Org商品明细;
            private string _Org完整库位号;

            [ExcelColumn("商品明细")]
            public string _商品明细
            {
                get
                {
                    return _Org商品明细;
                }
                set
                {
                    _Org商品明细 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("库位号")]
            public string _完整库位号
            {
                get
                {
                    return _Org完整库位号;
                }
                set
                {
                    _Org完整库位号 = value != null ? value.ToString().Trim() : "";
                }
            }

            public List<_拣货单> ToData()
            {
                var list = new List<_拣货单>();
                var _明细Arr = _商品明细.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                var _库位号Arr = _完整库位号.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                if (_明细Arr.Count == _库位号Arr.Count)
                    for (int idx = 0, len = _明细Arr.Count; idx < len; idx++)
                    {
                        var model = new _拣货单();
                        model._商品明细 = _明细Arr[idx] + ";";
                        model._完整库位号 = _库位号Arr[idx] + ";";
                        model._乱单 = true;
                        list.Add(model);
                    }


                return list;
            }
        }

        [ExcelTable("拣货人员配置表")]
        class _拣货人员配置
        {
            private string _Org库位;
            private string _Org配货人员;

            [ExcelColumn("库位")]
            public string _库位
            {
                get
                {
                    return _Org库位;
                }
                set
                {
                    _Org库位 = value != null ? value.ToString().Trim().ToUpper() : "";
                }
            }

            [ExcelColumn("配货人员")]
            public string _配货人员
            {
                get
                {
                    return _Org配货人员;
                }
                set
                {
                    _Org配货人员 = value != null ? value.ToString().Trim() : "";
                }
            }
        }

        [ExcelTable("本月上班时间信息")]
        class _本月上班时间信息
        {
            [ExcelColumn("姓名")]
            public string _姓名 { get; set; }

            [ExcelColumn("1号")]
            public decimal _1号 { get; set; }

            [ExcelColumn("2号")]
            public decimal _2号 { get; set; }

            [ExcelColumn("3号")]
            public decimal _3号 { get; set; }

            [ExcelColumn("4号")]
            public decimal _4号 { get; set; }

            [ExcelColumn("5号")]
            public decimal _5号 { get; set; }

            [ExcelColumn("6号")]
            public decimal _6号 { get; set; }

            [ExcelColumn("7号")]
            public decimal _7号 { get; set; }

            [ExcelColumn("8号")]
            public decimal _8号 { get; set; }

            [ExcelColumn("9号")]
            public decimal _9号 { get; set; }

            [ExcelColumn("10号")]
            public decimal _10号 { get; set; }

            [ExcelColumn("11号")]
            public decimal _11号 { get; set; }

            [ExcelColumn("12号")]
            public decimal _12号 { get; set; }

            [ExcelColumn("13号")]
            public decimal _13号 { get; set; }

            [ExcelColumn("14号")]
            public decimal _14号 { get; set; }

            [ExcelColumn("15号")]
            public decimal _15号 { get; set; }

            [ExcelColumn("16号")]
            public decimal _16号 { get; set; }

            [ExcelColumn("17号")]
            public decimal _17号 { get; set; }

            [ExcelColumn("18号")]
            public decimal _18号 { get; set; }

            [ExcelColumn("19号")]
            public decimal _19号 { get; set; }

            [ExcelColumn("20号")]
            public decimal _20号 { get; set; }

            [ExcelColumn("21号")]
            public decimal _21号 { get; set; }

            [ExcelColumn("22号")]
            public decimal _22号 { get; set; }

            [ExcelColumn("23号")]
            public decimal _23号 { get; set; }

            [ExcelColumn("24号")]
            public decimal _24号 { get; set; }

            [ExcelColumn("25号")]
            public decimal _25号 { get; set; }

            [ExcelColumn("26号")]
            public decimal _26号 { get; set; }

            [ExcelColumn("27号")]
            public decimal _27号 { get; set; }

            [ExcelColumn("28号")]
            public decimal _28号 { get; set; }

            [ExcelColumn("29号")]
            public decimal _29号 { get; set; }

            [ExcelColumn("30号")]
            public decimal _30号 { get; set; }

            [ExcelColumn("31号")]
            public decimal _31号 { get; set; }

        }

        [ExcelTable("帮忙点货时间")]
        class _帮忙点货时间
        {
            private string _Org姓名;

            [ExcelColumn("姓名")]
            public string _姓名
            {
                get
                {
                    return _Org姓名;
                }
                set
                {
                    _Org姓名 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("日期")]
            public DateTime _日期 { get; set; }

            [ExcelColumn("工作时间")]
            public DateTime _工作时间 { get; set; }

            public TimeSpan _帮忙总时间
            {
                get
                {
                    if (_工作时间 != null)
                    {
                        return new TimeSpan(_工作时间.Hour, _工作时间.Minute, 0);
                    }

                    return new TimeSpan(0, 0, 0);
                }
            }
        }

        /// <summary>
        /// 拣货人员存储格式
        /// </summary>
        class _拣货人员配置信息
        {
            public string _姓名 { get; set; }
            public string 管理库位 { get; set; }
        }

        class _配货绩效结果
        {
            public string _业绩归属人 { get; set; }
            public double _购买总数量
            {
                get
                {
                    return _购买总数量_正常 + _购买总数量_乱单;
                }
            }
            public double _拣货单张数
            {
                get
                {
                    return _拣货单张数_正常 + _拣货单张数_乱单;
                }
            }

            public double _购买总数量_正常 { get; set; }
            public double _购买总数量_乱单 { get; set; }
            public double _拣货单张数_正常 { get; set; }
            public double _拣货单张数_乱单 { get; set; }

            public string _总时长 { get; set; }
            public string _帮忙总时长 { get; set; }
            public double _分钟 { get; set; }
            public double _d张数定值 { get; set; }
            public double _d张数占比 { get; set; }
            public double _d数量定值 { get; set; }
            public double _d数量占比 { get; set; }

            public double _拣货单效率
            {
                get
                {
                    if (_分钟 <= 0)
                        return 0;
                    return Math.Round(_拣货单张数 / _分钟, 4);
                }
            }

            public double _购买数量效率
            {
                get
                {
                    if (_分钟 <= 0)
                        return 0;
                    return Math.Round(_购买总数量 / _分钟, 4);
                }
            }

            public double _小时
            {
                get
                {
                    if (_分钟 <= 0)
                        return 0;
                    var mm = _分钟 % 60;
                    var hh = (_分钟 - mm) / 60;
                    return hh + Math.Round(mm / 60, 4);
                }
            }

            public double _拣货单每小时
            {
                get
                {
                    if (_小时 <= 0)
                        return 0;
                    return Math.Round(_拣货单张数 / _小时, 4);
                }
            }

            public double _个数每小时
            {
                get
                {
                    if (_小时 <= 0)
                        return 0;
                    return Math.Round(_购买总数量 / _小时, 4);
                }
            }

            public double _定值倍数
            {
                get
                {
                    //= 拣货单每小时 / 208 * 0.75 + 个数每小时 / 1186 * 0.25

                    return Math.Round(_拣货单每小时 / _d张数定值 * _d张数占比 + _个数每小时 / _d数量定值 * _d数量占比, 4);
                }
            }

            public double _工资
            {
                get
                {
                    if (_小时 <= 0)
                        return 0;
                    //=IF(定值倍数>1,(定值倍数-1)*3000,0)
                    if (_定值倍数 > 1)
                        return Math.Round((_定值倍数 - 1) * 3000, 2);
                    else
                        return 0;
                }
            }
        }

        class _订单详情数据
        {
            public string SKU { get; set; }
            public double Amount { get; set; }
            public bool _乱单 { get; set; }
        }

        class DateHelper
        {
            /// <summary>
            /// csv读取时间的时候会自动加上日期部分,造成系统异常
            /// </summary>
            /// <param name="str"></param>
            /// <returns></returns>
            public static string GetPureTimeString(string str)
            {
                if (str.IndexOf('/') > 0 || str.IndexOf('-') > 0)
                {
                    var idx = str.IndexOf(' ');
                    return str.Substring(idx, str.Length - idx);
                }
                return str;
            }
        }


    }
}
