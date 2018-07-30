using CommonLibs;
using Gadget.Libs;
using LinqToExcel.Attributes;
using Newtonsoft.Json;
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
        //private List<_工号记录详细信息> _List近两个月历史点货记录 { get { return _List上个月历史点货记录.AddRange(_List当月历史点货记录); } }
        List<_工号记录详细信息> _List上个月历史点货记录 = new List<_工号记录详细信息>();
        List<_工号记录详细信息> _List当月历史点货记录 = new List<_工号记录详细信息>();
        private string _CacheFolder { get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "缓存信息"); } }
        private string _当月历史点货记录信息Path { get { return Path.Combine(_CacheFolder, DateTime.Now.ToString("yyyy-MM") + ".json"); } }
        private string _上月历史点货记录信息Path { get { return Path.Combine(_CacheFolder, DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ".json"); } }
        public _点货绩效()
        {
            InitializeComponent();
        }

        private void _点货绩效_Load(object sender, EventArgs e)
        {
            //txt入库明细.Text = @"C:\Users\Leon\Desktop\绩效\采购入库单号-对应工号7月份.csv";
            //txt产品订单.Text = @"C:\Users\Leon\Desktop\绩效\采购入库明细表7月份.csv";
            //txt人员代号.Text = @"C:\Users\Leon\Desktop\绩效\人员代号.csv";
            //txt积分参数.Text = @"C:\Users\Leon\Desktop\绩效\积分参数.csv";
            //txt工号记录.Text = @"C:\Users\Leon\Desktop\绩效\工号记录.csv";
            //txt产品等级.Text = @"C:\Users\Leon\Desktop\绩效\产品等级.csv";

            if (File.Exists(_当月历史点货记录信息Path))
                using (var fs = new FileStream(_当月历史点货记录信息Path, FileMode.Open))
                using (var reader = new StreamReader(fs))
                {
                    var str = reader.ReadToEnd();
                    _List当月历史点货记录 = JsonConvert.DeserializeObject<List<_工号记录详细信息>>(str);
                }

            //if (File.Exists(_上月历史点货记录信息Path))
            //    using (var fs = new FileStream(_上月历史点货记录信息Path, FileMode.Open))
            //    using (var reader = new StreamReader(fs))
            //    {
            //        var str = reader.ReadToEnd();
            //        _List上个月历史点货记录.AddRange(JsonConvert.DeserializeObject<List<_工号记录详细信息>>(str));
            //    }

        }

        /**************** button event ****************/

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormHelper.GenerateTableDes(typeof(_入库明细Mapping), typeof(_人员代号Mapping), typeof(_积分参数Mapping), typeof(_工号记录表), typeof(_产品等级), typeof(_订单产品));
        }
        #endregion

        #region 上传入库明细
        private void btn上传入库明细_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt入库明细);
        }
        #endregion

        #region 上传人员代号
        private void btn上传人员代号_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt人员代号);
        }
        #endregion

        #region 上传积分参数
        private void btn上传积分参数_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt积分参数);
        }
        #endregion

        #region 上传热销订单
        private void btn上传订单产品_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt产品订单);
        }
        #endregion

        #region 上传工号记录
        private void btn上传工号记录_Click(object sender, EventArgs e)
        {
            var strError = string.Empty;
            FormHelper.GetCSVPath(txt工号记录, () =>
            {
                var list工号记录详细信息 = new List<_工号记录详细信息>();
                var list工号记录 = new List<_工号记录表>();
                ShowGongMsg("开始读取工号记录数据");
                FormHelper.ReadCSVFile(txt工号记录.Text, ref list工号记录, ref strError);
                for (int idx = 0, len = list工号记录.Count; idx < len; idx += 2)
                {
                    var model = new _工号记录详细信息();
                    model._入库单号 = list工号记录[idx]._工号记录;
                    model._工号 = list工号记录[idx + 1]._工号记录;
                    model._操作日期 = DateTime.Now;
                    list工号记录详细信息.Add(model);
                }
                ShowGongMsg("开始存储工号记录数据,请稍后");
                //var msg = JsonConvert.SerializeObject(list工号记录详细信息);
                #region 记录点货历史信息
                {
                    //var history = _List近两个月历史点货记录.Where(x => !list工号记录详细信息.Any(u => u._入库单号 == x._入库单号)).ToList();
                    //list工号记录详细信息.AddRange(history);

                    _List当月历史点货记录.AddRange(list工号记录详细信息);


                    if (!Directory.Exists(_CacheFolder))
                        Directory.CreateDirectory(_CacheFolder);

                    using (var writer = new StreamWriter(_当月历史点货记录信息Path, false))
                    {
                        _List当月历史点货记录 = list工号记录详细信息.Where(x => !string.IsNullOrWhiteSpace(x._入库单号)).ToList();
                        var str = JsonConvert.SerializeObject(_List当月历史点货记录);
                        writer.Write(str);
                    }
                }
                #endregion

                ShowGongMsg("开始工号记录存储完毕");
            });
        }
        #endregion

        #region 上传产品等级
        private void btn产品等级_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt产品等级);
        }
        #endregion

        #region 处理数据
        private void btn处理_Click(object sender, EventArgs e)
        {
            btn处理.Enabled = false;
            var list入库明细 = new List<_入库明细Mapping>();
            var list订单产品 = new List<_订单产品>();
            var list产品等级 = new List<_产品等级>();
            var list人员代号 = new List<_人员代号Mapping>();
            var list积分参数 = new List<_积分参数Mapping>();
            var list点货绩效 = new List<_点货绩效Model>();
            var list工号记录详细信息 = new List<_工号记录详细信息>();


            #region 读取数据
            var actReadData = new Action(() =>
            {
                var strError = string.Empty;

                ShowMsg("开始读取表格信息");

                #region 读取入库明细
                {
                    ShowMsg("开始读取入库明细数据");
                    FormHelper.ReadCSVFile(txt入库明细.Text, ref list入库明细, ref strError);
                }
                #endregion

                #region 读取人员代号
                {
                    ShowMsg("开始读取人员代号数据");
                    FormHelper.ReadCSVFile(txt人员代号.Text, ref list人员代号, ref strError);
                }
                #endregion

                #region 读取积分参数
                {
                    ShowMsg("开始读取库积分参数数据");
                    FormHelper.ReadCSVFile(txt积分参数.Text, ref list积分参数, ref strError);
                }
                #endregion

                #region 读取订单产品
                {
                    ShowMsg("开始读取订单产品数据");
                    FormHelper.ReadCSVFile(txt产品订单.Text, ref list订单产品, ref strError);
                }
                #endregion

                #region 读取产品等级
                {
                    ShowMsg("开始读取产品等级数据");
                    FormHelper.ReadCSVFile(txt产品等级.Text, ref list产品等级, ref strError);
                }
                #endregion

                #region 读取工号记录
                {
                    var list工号记录 = new List<_工号记录表>();
                    ShowMsg("开始读取工号记录数据");
                    FormHelper.ReadCSVFile(txt工号记录.Text, ref list工号记录, ref strError);
                    for (int idx = 0, len = list工号记录.Count; idx < len; idx += 2)
                    {
                        var model = new _工号记录详细信息();
                        model._入库单号 = list工号记录[idx]._工号记录;
                        model._工号 = list工号记录[idx + 1]._工号记录;
                        var ref人员 = list人员代号.Where(x => x._代号 == model._工号).FirstOrDefault();
                        if (ref人员 != null)
                            model._员工姓名 = ref人员._姓名;
                        list工号记录详细信息.Add(model);
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

                        //if (item._入库单退回单号== "RKD201807210725")
                        //{

                        //}

                        if (!string.IsNullOrEmpty(item._人员代码))
                        {
                            var ref人员 = list人员代号.Where(x => x._代号 == item._人员代码).FirstOrDefault();
                            if (ref人员 != null)
                            {
                                item._人员姓名 = ref人员._姓名;

                                bool bExist = false;
                                for (int ii = list工号记录详细信息.Count - 1; ii >= 0; ii--)
                                {
                                    var referIn工号记录详情 = list工号记录详细信息[ii];
                                    if (referIn工号记录详情 != null)
                                    {
                                        var mm = new _工号记录详细信息();
                                        mm._入库单号 = item._入库单退回单号;
                                        mm._员工姓名 = ref人员._姓名;
                                        mm._工号 = ref人员._代号;
                                        list工号记录详细信息.Add(mm);
                                        bExist = true;
                                        break;
                                    }
                                }
                                if (!bExist)
                                {
                                    var mm = new _工号记录详细信息();
                                    mm._入库单号 = item._入库单退回单号;
                                    mm._员工姓名 = ref人员._姓名;
                                    mm._工号 = ref人员._代号;
                                    list工号记录详细信息.Add(mm);
                                }
                            }
                            else
                                item._人员姓名 = _未指定人员;
                        }
                        else
                        {
                            var ref人员 = list工号记录详细信息.Where(x => x._入库单号 == item._入库单退回单号).FirstOrDefault();
                            if (ref人员 != null)
                                item._人员姓名 = ref人员._员工姓名;
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
                        //找到明细的所有产品,再查产品对应等级的积分
                        var refer产品 = list订单产品.Where(x => x._入库单号 == item._入库单退回单号).ToList();

                        foreach (var referProduct in refer产品)
                        {
                            //查找对应等级
                            var refer等级 = list产品等级.FirstOrDefault(x => x.SKU == referProduct.SKU);
                            if (refer等级 != null)
                            {
                                var ref积分 = list积分参数.Where(x => x._等级.ToLower() == refer等级._等级.ToLower() && x._左区间 < item._总数量 && x._右区间 >= item._总数量).FirstOrDefault();
                                if (ref积分 != null)
                                {
                                    item._盘点积分 += ref积分._积分;
                                    item._积分详情 += string.Format("{0}-数量:{1},积分:{2};", referProduct.SKU, referProduct._采购数量, ref积分._积分);
                                }
                            }
                            else
                            {
                                var ref积分 = list积分参数.Where(x => x._等级.ToLower() == "d" && x._左区间 < item._总数量 && x._右区间 >= item._总数量).FirstOrDefault();
                                if (ref积分 != null)
                                {
                                    item._盘点积分 += ref积分._积分;
                                    item._积分详情 += string.Format("{0}-数量:{1},积分:{2};", referProduct.SKU, referProduct._采购数量, ref积分._积分);
                                }
                            }
                        }




                        //var ref积分 = list积分参数.Where(x => x._左区间 < item._总数量 && x._右区间 >= item._总数量).FirstOrDefault();
                        //if (ref积分 != null)
                        //{
                        //    item._盘点积分 = ref积分._积分;
                        //}

                        //item._是否热销订单 = list热销订单.Where(x => x._热销单号 == item._入库单退回单号).Count() > 0;
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

                #region 记录点货历史信息
                //{
                //    var history = _List近两个月历史点货记录.Where(x => !list工号记录详细信息.Any(u => u._入库单号 == x._入库单号)).ToList();
                //    list工号记录详细信息.AddRange(history);


                //    if (!Directory.Exists(_CacheFolder))
                //        Directory.CreateDirectory(_CacheFolder);

                //    using (var writer = new StreamWriter(_当月历史点货记录信息Path, false))
                //    {
                //        _List近两个月历史点货记录 = list工号记录详细信息.Where(x => !string.IsNullOrWhiteSpace(x._入库单号)).ToList();
                //        var str = JsonConvert.SerializeObject(_List近两个月历史点货记录);
                //        writer.Write(str);
                //    }

                //}
                #endregion

                Export(list点货绩效.OrderByDescending(x => x._总积分).ToList(), list入库明细);
            }, null);
            #endregion
        }
        #endregion

        #region 查询历史点货信息
        private void btn查询_Click(object sender, EventArgs e)
        {
            lb人员名称.Text = "";
            lb人员代号.Text = "";
            var str入库单号 = txt入库单号.Text;
            if (!string.IsNullOrWhiteSpace(str入库单号))
            {
                var _h历史点货记录 = new List<_工号记录详细信息>();
                var qMonth = dtp所在月份.Value.ToString("yyyy-MM");

                #region 读取数据
                var actReadData = new Action(() =>
                {
                    //非当月,需要遍历数据查询
                    if (qMonth != DateTime.Now.ToString("yyyy-MM"))
                    {
                        var path = Path.Combine(_CacheFolder, qMonth + ".json");
                        if (File.Exists(path))
                            using (var fs = new FileStream(path, FileMode.Open))
                            using (var reader = new StreamReader(fs))
                            {
                                var str = reader.ReadToEnd();
                                _h历史点货记录 = JsonConvert.DeserializeObject<List<_工号记录详细信息>>(str);
                            }
                    }
                    else
                    {
                        _h历史点货记录 = _List当月历史点货记录;
                    }
                });
                #endregion

                #region 处理数据
                actReadData.BeginInvoke((obj) =>
                {
                    var refer = _h历史点货记录.FirstOrDefault(x => x._入库单号 == str入库单号);
                    if (refer != null)
                        ShowQueryResult(refer);
                }, null);
                #endregion
            }
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
                    sheet1.Cells[1, 4].Value = "积分";
                    sheet1.Cells[1, 4].Value = "积分详情";

                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = detailList.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = detailList[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._入库单退回单号;
                        sheet1.Cells[rowIdx, 2].Value = info._人员姓名;
                        sheet1.Cells[rowIdx, 3].Value = info._总数量;
                        sheet1.Cells[rowIdx, 4].Value = info._最后积分;
                        sheet1.Cells[rowIdx, 5].Value = info._积分详情;
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
                    btn处理.Enabled = true;
                }
            }, null);
            #endregion
        }
        #endregion

        #region 展示查询结果
        private void ShowQueryResult(_工号记录详细信息 data)
        {
            InvokeMainForm((obj) =>
            {
                lb人员名称.Text = data._员工姓名;
                lb人员代号.Text = data._工号;
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

        #region ShowGongMsg 工号上传消息提示
        /// <summary>
        /// 消息提示
        /// </summary>
        /// <param name="strMsg"></param>
        private void ShowGongMsg(string strMsg)
        {
            if (this.InvokeRequired)
            {
                var act = new Action<string>(ShowMsg);
                this.Invoke(act, strMsg);
            }
            else
            {
                this.lbGongMsg.Text = strMsg;
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
                    //if (_是否热销订单)
                    //    return Math.Round(_盘点积分 / 2, 4);
                    //else
                    return _盘点积分;
                }
            }
            public string _积分详情 { get; set; }
        }

        [ExcelTable("订单对应产品")]
        class _订单产品
        {
            private string orgSKU;
            [ExcelColumn("入库单号")]
            public string _入库单号 { get; set; }
            [ExcelColumn("商品SKU")]
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
            [ExcelColumn("采购数量")]
            public decimal _采购数量 { get; set; }
        }

        [ExcelTable("人员工号")]
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
            [ExcelColumn("等级")]
            public string _等级 { get; set; }

            [ExcelColumn("左区间")]
            public decimal _左区间 { get; set; }

            [ExcelColumn("右区间")]
            public decimal _右区间 { get; set; }

            [ExcelColumn("积分")]
            public decimal _积分 { get; set; }
        }

        [ExcelTable("工号记录")]
        class _工号记录表
        {
            [ExcelColumn("工号记录")]
            public string _工号记录
            {
                get; set;
            }
        }

        [ExcelTable("产品等级")]
        class _产品等级
        {
            private string orgSKU;
            [ExcelColumn("等级")]
            public string _等级 { get; set; }
            [ExcelColumn("SKU")]
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

        class _工号记录详细信息
        {
            public string _入库单号 { get; set; }
            public string _工号 { get; set; }
            public string _员工姓名 { get; set; }
            public DateTime _操作日期 { get; set; }
        }

        class _点货绩效Model
        {
            public string _点货人 { get; set; }
            public decimal _总积分 { get; set; }
            public int _入库单数 { get; set; }
        }

    }
}
