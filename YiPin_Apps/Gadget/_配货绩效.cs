using CommonLibs;
using Gadget.Libs;
using LinqToExcel.Attributes;
using Newtonsoft.Json;
using OfficeOpenXml;
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

        private string MonthFlag
        {
            get
            {
                var currentTime = dtp绩效时间.Value;
                return currentTime.ToString("yyyy-MM");
            }
        }

        private DateTime CulcTime = new DateTime();

        private List<_拣货人员配置信息> _人员负责库位信息;

        public _配货绩效()
        {
            InitializeComponent();
        }

        private void _配货绩效_Load(object sender, EventArgs e)
        {

            _人员负责库位信息 = new List<_拣货人员配置信息>();
            if (!Directory.Exists(Folder人员配置))
                Directory.CreateDirectory(Folder人员配置);
            if (!Directory.Exists(Folder拣货绩效))
                Directory.CreateDirectory(Folder拣货绩效);
            //加载缓存人员配置文件
            if (File.Exists(FileName拣货人员配置缓存文件))
            {
                using (var fs = new StreamReader(FileName拣货人员配置缓存文件, Encoding.UTF8))
                {
                    var json = fs.ReadToEnd();
                    _人员负责库位信息 = JsonConvert.DeserializeObject<List<_拣货人员配置信息>>(json);
                }
            }
            //
            RefreshCache();
        }

        /**************** button event ****************/

        #region 上传拣货单
        private void btn上传拣货单_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt拣货单);
        }
        #endregion

        #region 上传拣货时间
        private void btn上传拣货时间_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt拣货时间, () =>
             {
                 btn当天绩效.Enabled = true;
             });
        }
        #endregion

        #region 上传拣货人员配置
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
                         ShowMsg("拣货人员配置数据存储完毕");
                     }
                     ShowMsg(strError);
                 }, null);
                 #endregion
             });
        }
        #endregion

        #region 计算当天绩效
        private void btn当天绩效_Click(object sender, EventArgs e)
        {
            if (_人员负责库位信息 != null && _人员负责库位信息.Count > 0)
            {
                try
                {
                    CulcTime = dtp绩效时间.Value;
                    btn当天绩效.Enabled = false;
                    btn全月绩效.Enabled = false;
                    var strError = string.Empty;
                    var list拣货单 = new List<_拣货单>();
                    var list拣货时间 = new List<_拣货时间>();
                    var list最终绩效 = new List<_配货绩效结果>();
                    #region 读取数据
                    var actReadData = new Action(() =>
                    {
                        ShowMsg("开始读取当天绩效信息");
                        FormHelper.ReadCSVFile(txt拣货单.Text, ref list拣货单, ref strError);
                        FormHelper.ReadCSVFile(txt拣货时间.Text, ref list拣货时间, ref strError);
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
                                    var md = new _配货绩效结果();
                                    md._业绩归属人 = name;
                                    var _订单详情数据 = new List<_订单详情数据>();

                                    #region 抽取详细信息
                                    {
                                        var refLh = (from it in list拣货单
                                                     join s in _人员负责库位信息 on it._库位号 equals s.管理库位
                                                     where s._姓名 == name
                                                     select it._拣货明细).ToList();
                                        foreach (List<string> item in refLh)
                                        {
                                            foreach (var it in item)
                                            {
                                                var arr = it.Split(new string[] { "*" }, StringSplitOptions.RemoveEmptyEntries);
                                                var detail = new _订单详情数据();
                                                detail.SKU = arr[0].Trim();
                                                detail.Amount = Convert.ToDouble(arr[1]);
                                                _订单详情数据.Add(detail);
                                            }
                                        }
                                    }
                                    #endregion

                                    var refTime = list拣货时间.Where(x => x._姓名 == name).FirstOrDefault();
                                    md._拣货单张数 = _订单详情数据.Select(x => x.SKU).Distinct().Count();
                                    md._购买总数量 = _订单详情数据.Select(x => x.Amount).Sum();
                                    if (refTime != null)
                                    {
                                        refTime.CulcTime = CulcTime;
                                        var mm = refTime._拣货总时间.TotalMinutes % 60;
                                        var hh = (refTime._拣货总时间.TotalMinutes - mm) / 60;
                                        md._总时长 = string.Format("{0}:{1}:00", hh > 9 ? "" + hh : "0" + hh, mm > 9 ? "" + mm : "0" + mm);
                                        md._分钟 = refTime._拣货总时间.TotalMinutes;
                                    }

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
                        md._购买总数量 = list.Where(x => x._业绩归属人 == n).Select(x => x._购买总数量).Sum();
                        md._拣货单张数 = list.Where(x => x._业绩归属人 == n).Select(x => x._拣货单张数).Sum();
                        //md._总时长 = list.Where(x => x._业绩归属人 == n).Select(x => x._总时长).Sum();
                        md._分钟 = list.Where(x => x._业绩归属人 == n).Select(x => x._分钟).Sum();
                        var mm = md._分钟 % 60;
                        var hh = (md._分钟 - mm) / 60;
                        md._总时长 = string.Format("{0}:{1}:00", hh > 9 ? "" + hh : "0" + hh, mm > 9 ? "" + mm : "0" + mm);
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
            FormHelper.GenerateTableDes(typeof(_拣货单), typeof(_拣货时间), typeof(_拣货人员配置));
        }
        #endregion

        #region 刷新缓存信息
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            RefreshCache();
        } 
        #endregion

        /**************** common method ****************/

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
                sheet1.Cells[1, 1].Value = "业绩归属人";
                sheet1.Cells[1, 2].Value = "购买总数量";
                sheet1.Cells[1, 3].Value = "拣货单张数";
                sheet1.Cells[1, 4].Value = "总时长";
                sheet1.Cells[1, 5].Value = "分钟";
                sheet1.Cells[1, 6].Value = "拣货单效率";
                sheet1.Cells[1, 7].Value = "购买数量效率";
                sheet1.Cells[1, 8].Value = "小时";
                sheet1.Cells[1, 9].Value = "拣货单每小时";
                sheet1.Cells[1, 10].Value = "个数每小时";
                sheet1.Cells[1, 11].Value = "定值倍数";
                sheet1.Cells[1, 12].Value = "工资";

                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = list拣货绩效.Count; idx < len; idx++)
                {
                    var curOrder = list拣货绩效[idx];
                    sheet1.Cells[rowIdx, 1].Value = curOrder._业绩归属人;
                    sheet1.Cells[rowIdx, 2].Value = curOrder._购买总数量;
                    sheet1.Cells[rowIdx, 3].Value = curOrder._拣货单张数;
                    sheet1.Cells[rowIdx, 4].Value = curOrder._总时长;
                    sheet1.Cells[rowIdx, 5].Value = curOrder._分钟;
                    sheet1.Cells[rowIdx, 6].Value = curOrder._拣货单效率;
                    sheet1.Cells[rowIdx, 7].Value = curOrder._购买数量效率;
                    sheet1.Cells[rowIdx, 8].Value = curOrder._小时;
                    sheet1.Cells[rowIdx, 9].Value = curOrder._拣货单每小时;
                    sheet1.Cells[rowIdx, 10].Value = curOrder._个数每小时;
                    sheet1.Cells[rowIdx, 11].Value = curOrder._定值倍数;
                    sheet1.Cells[rowIdx, 12].Value = curOrder._工资;
                    rowIdx++;
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

        }

        [ExcelTable("拣货时间表")]
        class _拣货时间
        {
            private string _Org姓名;
            private string _Org拣货单开始时间;
            private string _Org拣货单结束时间;

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

            [ExcelColumn("拣货单开始时间")]
            public string _Str拣货单开始时间
            {
                get
                {
                    return _Org拣货单开始时间;
                }
                set
                {
                    _Org拣货单开始时间 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("拣货单结束时间")]
            public string _Str拣货单结束时间
            {
                get
                {
                    return _Org拣货单结束时间;
                }
                set
                {
                    _Org拣货单结束时间 = value != null ? value.ToString().Trim() : "";
                }
            }

            public DateTime _拣货单开始时间
            {
                get
                {
                    var dtString = _Str拣货单开始时间.Replace("：", ":").Replace(";", ":").Replace("；", ":");

                    var arr = DateHelper.GetPureTimeString(dtString).Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                    if (arr.Length > 0)
                    {
                        var ct = CulcTime;
                        return new DateTime(ct.Year, ct.Month, ct.Day, Convert.ToInt32(arr[0]), Convert.ToInt32(arr[1]), 0);
                    }
                    return DateTime.MinValue;
                }
            }

            public DateTime _拣货单结束时间
            {
                get
                {
                    var dtString = _Str拣货单结束时间.Replace("：", ":").Replace(";", ":").Replace("；", ":");
                    var arr = DateHelper.GetPureTimeString(dtString).Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                    if (arr.Length > 0)
                    {
                        var ct = CulcTime;
                        return new DateTime(ct.Year, ct.Month, ct.Day, Convert.ToInt32(arr[0]), Convert.ToInt32(arr[1]), 0);
                    }
                    return DateTime.MinValue;
                }
            }

            public TimeSpan _拣货总时间
            {
                get
                {
                    var _吃饭一小时 = new TimeSpan(1, 0, 0);
                    var _吃饭起始时间 = new DateTime(CulcTime.Year, CulcTime.Month, CulcTime.Day, 12, 0, 0);
                    var _吃饭结束时间 = new DateTime(CulcTime.Year, CulcTime.Month, CulcTime.Day, 13, 0, 0);
                    if (_拣货单开始时间 < _吃饭起始时间 && _拣货单结束时间 > _吃饭结束时间)
                    {
                        return _拣货单结束时间 - _拣货单开始时间 - _吃饭一小时;
                    }
                    return _拣货单结束时间 - _拣货单开始时间;
                }
            }

            public DateTime CulcTime { get; set; }
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
            public double _购买总数量 { get; set; }
            public double _拣货单张数 { get; set; }
            public string _总时长 { get; set; }
            public double _分钟 { get; set; }


            public double _拣货单效率
            {
                get
                {
                    return Math.Round(_拣货单张数 / _分钟, 4);
                }
            }

            public double _购买数量效率
            {
                get
                {
                    return Math.Round(_购买总数量 / _分钟, 4);
                }
            }

            public double _小时
            {
                get
                {
                    var mm = _分钟 % 60;
                    var hh = (_分钟 - mm) / 60;
                    return hh + Math.Round(mm / 60, 4);
                }
            }

            public double _拣货单每小时
            {
                get
                {
                    return Math.Round(_拣货单张数 / _小时, 4);
                }
            }

            public double _个数每小时
            {
                get
                {
                    return Math.Round(_购买总数量 / _小时, 4);
                }
            }

            public double _定值倍数
            {
                get
                {
                    //= 拣货单每小时 / 208 * 0.75 + 个数每小时 / 1186 * 0.25

                    return Math.Round(_拣货单每小时 / 208 * 0.75 + _个数每小时 / 1186 * 0.25, 4);
                }
            }

            public double _工资
            {
                get
                {
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
                if (str.IndexOf('/') > 0|| str.IndexOf('-') > 0)
                {
                    var idx = str.IndexOf(' ');
                    return str.Substring(idx, str.Length - idx);
                }
                return str;
            }
        }


    }
}
