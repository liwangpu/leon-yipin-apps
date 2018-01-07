using CommonLibs;
using Gadget.Libs;
using LinqToExcel;
using LinqToExcel.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Gadget
{
    public partial class _整合供应商人员工资统计 : Form
    {

        public _整合供应商人员工资统计()
        {
            InitializeComponent();
        }

        private void _整合供应商人员工资统计_Load(object sender, EventArgs e)
        {
            //txt停售退货Path.Text = @"C:\Users\Leon\Desktop\工资\停售退货.csv";
            //txt滞销退货Path.Text = @"C:\Users\Leon\Desktop\工资\滞销退货.csv";
            //txt缺货率Path.Text = @"C:\Users\Leon\Desktop\工资\缺货率.csv";
            //txt上月入库时间差Path.Text = @"C:\Users\Leon\Desktop\工资\11月份入库时间差.csv";
            //txt当月入库时间差Path.Text = @"C:\Users\Leon\Desktop\工资\12月份入库时间差.csv";
            //txt压价信息Path.Text = @"C:\Users\Leon\Desktop\工资\压价详情.csv";
            //txt组员分配Path.Text = @"C:\Users\Leon\Desktop\工资\组员分配.csv";


        }

        /**************** property ****************/

        private decimal _dGrade1;
        private decimal _dGrade2;
        private decimal _dGrade3;

        /**************** button event ****************/

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormHelper.GenerateTableDes(typeof(_退货信息), typeof(_缺货率详情), typeof(_入库时间差详情), typeof(_压价详情), typeof(_组员分配));
        }
        #endregion

        #region 上传停售退货
        private void btn上传停售退货_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt停售退货Path);
        }
        #endregion

        #region 上传滞销退货
        private void btn上传滞销退货_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt滞销退货Path);
        }
        #endregion

        #region 上传缺货信息
        private void btn上传缺货信息_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt缺货率Path);
        }
        #endregion

        #region 上传上月入库时间差
        private void btn上传上月入库时间差_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt上月入库时间差Path);
        }
        #endregion

        #region 上传当月入库时间差
        private void btn上传当月入库时间差_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt当月入库时间差Path);
        }
        #endregion

        #region 上传压价信息
        private void btn上传压价信息_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt压价信息Path);
        }
        #endregion

        #region 上传组员分配
        private void btn上传组员分配_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt组员分配Path);

        }
        #endregion

        #region 处理按钮事件
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            var list停售退货信息 = new List<_退货信息>();
            var list滞销退货信息 = new List<_退货信息>();
            var list缺货率详情 = new List<_缺货率详情>();
            var list上月入库时间差详情 = new List<_入库时间差详情>();
            var list当月入库时间差详情 = new List<_入库时间差详情>();
            var list压价信息 = new List<_压价详情>();
            var list组员分配 = new List<_组员分配>();


            var d停售退货奖励比率 = nup停售退货奖励比率.Value / 100;
            var d滞销退货奖励比率 = nup滞销退货奖励比率.Value / 100;
            var d压价奖励比率 = nup压价奖励比率.Value / 100;

            _dGrade1 = nupGrad1.Value;
            _dGrade2 = nupGrad2.Value;
            _dGrade3 = nupGrad3.Value;

            #region 读取数据
            var actReadData = new Action(() =>
            {
                var strError = string.Empty;

                ShowMsg("开始读取停售退货信息");
                FormHelper.ReadCSVFile<_退货信息>(txt停售退货Path.Text, ref list停售退货信息, ref strError);
                if (!string.IsNullOrEmpty(strError))
                    ShowMsg("读取停售退货信息出现异常:" + strError);

                ShowMsg("开始读取滞销退货信息");
                FormHelper.ReadCSVFile<_退货信息>(txt滞销退货Path.Text, ref list滞销退货信息, ref strError);
                if (!string.IsNullOrEmpty(strError))
                    ShowMsg("读取滞销退货信息出现异常:" + strError);

                ShowMsg("开始读取缺货率详情信息");
                FormHelper.ReadCSVFile<_缺货率详情>(txt缺货率Path.Text, ref list缺货率详情, ref strError);
                if (!string.IsNullOrEmpty(strError))
                    ShowMsg("读取缺货率信息出现异常:" + strError);

                ShowMsg("开始读取上月入库时间差信息");
                FormHelper.ReadCSVFile<_入库时间差详情>(txt上月入库时间差Path.Text, ref list上月入库时间差详情, ref strError);
                if (!string.IsNullOrEmpty(strError))
                    ShowMsg("读取上月入库时间差信息出现异常:" + strError);

                ShowMsg("开始读取当月入库时间差信息");
                FormHelper.ReadCSVFile<_入库时间差详情>(txt当月入库时间差Path.Text, ref list当月入库时间差详情, ref strError);
                if (!string.IsNullOrEmpty(strError))
                    ShowMsg("读取当月入库时间差信息出现异常:" + strError);


                ShowMsg("开始读取压价信息");
                FormHelper.ReadCSVFile<_压价详情>(txt压价信息Path.Text, ref list压价信息, ref strError);
                if (!string.IsNullOrEmpty(strError))
                    ShowMsg("读取压价信息出现异常:" + strError);


                ShowMsg("开始读取组员分配信息");
                FormHelper.ReadCSVFile<_组员分配>(txt组员分配Path.Text, ref list组员分配, ref strError);
                if (!string.IsNullOrEmpty(strError))
                    ShowMsg("读取组员分配信息出现异常:" + strError);



            });
            #endregion


            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                var list统计结果 = new List<_工资统计>();
                var list缺货率详细信息 = new List<_缺货率详细信息>();
                var list入库时间差详细信息 = new List<_入库时间差详细信息>();

                var list整合人员姓名 = list停售退货信息.Select(x => x._退货人员).Distinct().ToList();


                list整合人员姓名.ForEach(str整合人员姓名 =>
                {
                    var model = new _工资统计();
                    model._采购员 = str整合人员姓名;


                    #region 计算停售退货奖励
                    {
                        var ref停售退货Item = list停售退货信息.Where(x => x._退货人员 == str整合人员姓名).FirstOrDefault();
                        if (ref停售退货Item != null)
                        {
                            model._停售退货奖励 = d停售退货奖励比率 * ref停售退货Item._退款金额;
                        }
                    }
                    #endregion

                    #region 计算滞销退货奖励
                    {
                        var ref滞销退货Item = list滞销退货信息.Where(x => x._退货人员 == str整合人员姓名).FirstOrDefault();
                        if (ref滞销退货Item != null)
                        {
                            model._滞销退货奖励 = d滞销退货奖励比率 * ref滞销退货Item._退款金额;
                        }
                    }
                    #endregion

                    //计算缺货率和入库时间差共用
                    var ref组员姓名 = new List<string>();

                    #region 获取所有组员信息
                    {
                        var str组员信息Item = list组员分配.Where(x => x._整合人员 == str整合人员姓名).FirstOrDefault();
                        if (str组员信息Item != null && !string.IsNullOrEmpty(str组员信息Item._组员))
                            ref组员姓名.AddRange(str组员信息Item._组员.Split(';').ToList());
                    }
                    #endregion

                    #region 计算组员缺货率奖励
                    {
                        var ref组员缺货率 = new List<decimal>();
                        ref组员姓名.ForEach(str组员姓名 =>
                        {
                            var ref组员缺货率list = list缺货率详情.Where(x => x._采购员 == str组员姓名).ToList();
                            if (ref组员缺货率list != null && ref组员缺货率list.Count > 0)
                            {
                                var _订单总和 = ref组员缺货率list.Select(x => x._交易订单数量).Sum();
                                var _缺货订单 = ref组员缺货率list.Select(x => x._缺货订单数量).Sum();
                                var _缺货率 = _订单总和 != 0 ? _缺货订单 / _订单总和 : 0;
                                ref组员缺货率.Add(_缺货率);

                                var tmp = new _缺货率详细信息();
                                tmp._采购员 = str组员姓名;
                                tmp._组长 = str整合人员姓名;
                                tmp._总订单 = _缺货订单;
                                tmp._总缺货订单 = _缺货订单;
                                tmp._缺货率 = _缺货率;
                                list缺货率详细信息.Add(tmp);
                            }
                        });

                        if (ref组员缺货率.Count > 0)
                        {
                            var _所有组员平均缺货率 = ref组员缺货率.Average();
                            model._缺货率奖励 = Calcu缺货率奖励(_所有组员平均缺货率);
                        }
                    }
                    #endregion

                    #region 计算入库时间差
                    {
                        var ref组员入库时间差 = new List<decimal>();
                        ref组员姓名.ForEach(str组员姓名 =>
                        {
                            var list上月入库时间差 = list上月入库时间差详情.Where(x => x._采购员 == str组员姓名).ToList();
                            var list当月入库时间差 = list当月入库时间差详情.Where(x => x._采购员 == str组员姓名).ToList();

                            var _上月平均入库时间差 = list上月入库时间差.Count > 0 ? list上月入库时间差.Select(x => x._采购入库时间差).Sum() / list上月入库时间差.Count : 0;
                            var _当月平均入库时间差 = list当月入库时间差详情.Count > 0 ? list当月入库时间差详情.Select(x => x._采购入库时间差).Sum() / list当月入库时间差详情.Count : 0;
                            ref组员入库时间差.Add(_当月平均入库时间差 - _上月平均入库时间差);

                            var tmp = new _入库时间差详细信息();
                            tmp._采购员 = str组员姓名;
                            tmp._组长 = str整合人员姓名;
                            tmp._上月入库时间差 = _上月平均入库时间差;
                            tmp._当月入库时间差 = _当月平均入库时间差;
                            list入库时间差详细信息.Add(tmp);
                        });

                        if (ref组员入库时间差.Count > 0)
                        {
                            var _所有组员平均入库时间差 = ref组员入库时间差.Average();
                            model._入库时间差奖励 = Calcu入库时间差奖励(_所有组员平均入库时间差);
                        }
                    }
                    #endregion

                    #region 计算压价奖励
                    {
                        var ref压价Item = list压价信息.Where(x => x._采购员 == str整合人员姓名).FirstOrDefault();
                        if (ref压价Item != null)
                        {
                            model._压价奖励奖励 = d压价奖励比率 * ref压价Item._压价额;
                        }

                    }
                    #endregion


                    list统计结果.Add(model);
                });



                Export(list统计结果, list缺货率详细信息, list入库时间差详细信息);

            }, null);
            #endregion

        }
        #endregion

        /**************** common method ****************/

        #region Calcu缺货率奖励
        private decimal Calcu缺货率奖励(decimal radio)
        {
            if (radio >= 0m && radio < 0.3m)
            {
                return _dGrade1;
            }

            if (radio >= 0.3m && radio < 0.6m)
            {
                return _dGrade2;
            }

            if (radio >= 0.6m && radio < 0.9m)
            {
                return _dGrade3;
            }

            return 0;
        }
        #endregion

        #region Calcu入库时间差奖励
        private decimal Calcu入库时间差奖励(decimal radio)
        {
            return 0;
        }
        #endregion

        private void Export(List<_工资统计> list工资统计
            , List<_缺货率详细信息> list缺货率详细信息
            , List<_入库时间差详细信息> list入库时间差详细信息)
        {
            ShowMsg("计算完毕,开始生成表格");
            var buffer1 = new byte[0];

            #region 生成表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 工资统计
                {
                    var sheet1 = workbox.Worksheets.Add("工资统计");
                    #region 标题行
                    sheet1.Cells[1, 1].Value = "整合人员";
                    sheet1.Cells[1, 2].Value = "停售退货奖励";
                    sheet1.Cells[1, 3].Value = "滞销退货奖励";
                    sheet1.Cells[1, 4].Value = "压价奖励奖励";
                    sheet1.Cells[1, 5].Value = "缺货率奖励";
                    sheet1.Cells[1, 6].Value = "入库时间差奖励";
                    sheet1.Cells[1, 7].Value = "总工资";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list工资统计.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list工资统计[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._采购员;
                        sheet1.Cells[rowIdx, 2].Value = info._停售退货奖励;
                        sheet1.Cells[rowIdx, 3].Value = info._滞销退货奖励;
                        sheet1.Cells[rowIdx, 4].Value = info._压价奖励奖励;
                        sheet1.Cells[rowIdx, 5].Value = info._缺货率奖励;
                        sheet1.Cells[rowIdx, 6].Value = info._入库时间差奖励;
                        sheet1.Cells[rowIdx, 7].Value = info._总工资;
                    }
                    #endregion

                }
                #endregion

                #region 缺货率详细信息
                {
                    var sheet1 = workbox.Worksheets.Add("缺货率详细信息");
                    #region 标题行
                    sheet1.Cells[1, 1].Value = "采购员";
                    sheet1.Cells[1, 2].Value = "总订单";
                    sheet1.Cells[1, 3].Value = "总缺货订单";
                    sheet1.Cells[1, 4].Value = "缺货率";
                    sheet1.Cells[1, 5].Value = "组长";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list缺货率详细信息.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list缺货率详细信息[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._采购员;
                        sheet1.Cells[rowIdx, 2].Value = info._总订单;
                        sheet1.Cells[rowIdx, 3].Value = info._总缺货订单;
                        sheet1.Cells[rowIdx, 4].Value = info._缺货率;
                        sheet1.Cells[rowIdx, 5].Value = info._组长;
                    }
                    #endregion

                }
                #endregion

                #region 缺货率详细信息
                {
                    var sheet1 = workbox.Worksheets.Add("入库时间差详细信息");
                    #region 标题行
                    sheet1.Cells[1, 1].Value = "采购员";
                    sheet1.Cells[1, 2].Value = "上月入库时间差";
                    sheet1.Cells[1, 3].Value = "当月入库时间差";
                    sheet1.Cells[1, 4].Value = "差值情况";
                    sheet1.Cells[1, 5].Value = "组长";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list入库时间差详细信息.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list入库时间差详细信息[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._采购员;
                        sheet1.Cells[rowIdx, 2].Value = info._上月入库时间差;
                        sheet1.Cells[rowIdx, 3].Value = info._当月入库时间差;
                        sheet1.Cells[rowIdx, 4].Value = info._差值情况;
                        sheet1.Cells[rowIdx, 5].Value = info._组长;
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
                    btnAnalyze.Enabled = true;
                }
            }, null);
        }

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

        [ExcelTable("停售/滞销退货信息表")]
        class _退货信息
        {
            [ExcelColumn("退货人员")]
            public string _退货人员 { get; set; }

            [ExcelColumn("退款金额")]
            public decimal _退款金额 { get; set; }
        }

        [ExcelTable("缺货信息表")]
        class _缺货率详情
        {
            private string _org采购员;

            [ExcelColumn("交易订单数量")]
            public decimal _交易订单数量 { get; set; }

            [ExcelColumn("缺货订单数量")]
            public decimal _缺货订单数量 { get; set; }

            [ExcelColumn("采购员")]
            public string _采购员
            {
                get
                {
                    return _org采购员;
                }
                set
                {
                    _org采购员 = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }

            public decimal _缺货率
            {
                get
                {
                    return _交易订单数量 != 0 ? Math.Round(_缺货订单数量 / _交易订单数量, 4) : 0;
                }
            }
        }

        [ExcelTable("入库时间差信息表")]
        class _入库时间差详情
        {
            private string _SKU;

            [ExcelColumn("商品SKU")]
            public string SKU
            {
                get
                {
                    return _SKU;
                }
                set
                {
                    _SKU = !string.IsNullOrEmpty(value) ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("采购入库时间差")]
            public decimal _采购入库时间差 { get; set; }

            [ExcelColumn("制单人")]
            public string _采购员 { get; set; }
        }

        [ExcelTable("压价信息表")]
        class _压价详情
        {
            [ExcelColumn("采购员")]
            public string _采购员 { get; set; }

            [ExcelColumn("压价额")]
            public decimal _压价额 { get; set; }
        }

        [ExcelTable("组员分配(注意:组员信息以分号\";\"分隔)")]
        class _组员分配
        {
            [ExcelColumn("整合人员")]
            public string _整合人员 { get; set; }

            [ExcelColumn("组员")]
            public string _组员 { get; set; }

        }

        class _工资统计
        {
            public string _采购员 { get; set; }

            public decimal _停售退货奖励 { get; set; }

            public decimal _滞销退货奖励 { get; set; }

            public decimal _缺货率奖励 { get; set; }

            public decimal _入库时间差奖励 { get; set; }

            public decimal _压价奖励奖励 { get; set; }

            public decimal _总工资
            {
                get
                {
                    return _停售退货奖励 + _滞销退货奖励 + _缺货率奖励 + _压价奖励奖励;
                }
            }
        }

        class _缺货率详细信息
        {
            public string _采购员 { get; set; }

            public string _组长 { get; set; }

            public decimal _总订单 { get; set; }

            public decimal _总缺货订单 { get; set; }

            public decimal _缺货率 { get; set; }
        }

        class _入库时间差详细信息
        {
            public string _采购员 { get; set; }
            public string _组长 { get; set; }
            public decimal _上月入库时间差 { get; set; }
            public decimal _当月入库时间差 { get; set; }

            public decimal _差值情况
            {
                get
                {
                    return _当月入库时间差 - _上月入库时间差;
                }
            }
        }
    }
}
