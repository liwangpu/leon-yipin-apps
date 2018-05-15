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


namespace Gadget
{
    public partial class _乱单绩效 : Form
    {
        public _乱单绩效()
        {
            InitializeComponent();
        }


        #region Load
        private void _乱单绩效_Load(object sender, EventArgs e)
        {
            //txt商品库位.Text = @"C:\Users\Leon\Desktop\rrr\1111.csv";
            //txt商品明细.Text = @"C:\Users\Leon\Desktop\rrr\乱单5-15.csv";
        }
        #endregion

        #region 上传商品明细
        private void btn上传商品明细_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt商品明细);
        }
        #endregion

        #region 上传商品库位
        private void btn上传商品库位_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt商品库位, () =>
            {
                btn乱单绩效.Enabled = true;
            });
        }
        #endregion

        private void btn乱单绩效_Click(object sender, EventArgs e)
        {
            var list商品明细 = new List<_商品明细信息>();
            var list库位明细 = new List<_库位明细信息>();

            #region 读取数据
            var actReadData = new Action(() =>
            {
                var strError = string.Empty;

                ShowMsg("开始读取商品明细数据");
                FormHelper.ReadCSVFile(txt商品明细.Text, ref list商品明细, ref strError);
                ShowMsg("开始读取库位明细数据");
                FormHelper.ReadCSVFile(txt商品库位.Text, ref list库位明细, ref strError);
                ShowMsg(strError);
            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                double _乱单张数 = 0;
                double _乱单购买数量 = 0;
                double _本区域张数 = 0;
                double _本区域购买数量 = 0;
                double _跨区域张数 = 0;
                double _跨区域购买数量 = 0;
                double _跨楼层张数 = 0;
                double _跨楼层购买数量 = 0;

                var list统计详细信息 = new List<_统计详细信息>();

                #region 取出所有库位
                {
                    var kws = list库位明细.Select(x => x._库位).Distinct().ToList();
                    list统计详细信息.AddRange(kws.Select(x => new _统计详细信息() { _区域 = x }));
                }
                #endregion

                #region 匹配商品明细库位信息
                {
                    for (int idx = list商品明细.Count - 1; idx >= 0; idx--)
                    {
                        var curItem = list商品明细[idx];
                        var list订购信息 = new List<_商品订购信息>();
                        for (int nnn = curItem._详细信息.Count - 1; nnn >= 0; nnn--)
                        {
                            var curDetail = curItem._详细信息[nnn];
                            var refRecord = list库位明细.Where(x => x.SKU == curDetail.SKU).FirstOrDefault();
                            if (refRecord != null)
                            {
                                curDetail._库位 = refRecord._库位;
                            }
                            list订购信息.Add(curDetail);
                        }

                        #region 统计乱单张数
                        if (list订购信息.Count > 1)
                        {
                            _乱单张数++;
                            _乱单购买数量 += list订购信息.Select(x => x._数量).Sum();
                        }
                        #endregion

                        #region 统计本区域/跨区域
                        {
                            if (list订购信息.Select(x => x._库位).Distinct().Count() > 1)
                            {
                                _跨区域张数++;
                                _跨区域购买数量 += list订购信息.Select(x => x._数量).Sum();
                            }
                            else
                            {
                                _本区域张数++;
                                _本区域购买数量 += list订购信息.Select(x => x._数量).Sum();
                            }
                        }
                        #endregion

                        #region 统计详细信息
                        if (list订购信息.Count > 1)
                        {
                            var kys = list订购信息.Select(x => x._库位).Distinct().ToList();

                            foreach (var item in kys)
                            {
                                for (int mmm = list统计详细信息.Count - 1; mmm >= 0; mmm--)
                                {
                                    var model = list统计详细信息[mmm];
                                    if (model._区域 == item)
                                    {
                                        model._乱单张数++;
                                        model._乱单购买数量 += list订购信息.Where(x => x._库位 == item).Select(x => x._数量).Sum();
                                        break;
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                }
                #endregion

                ExportExcel(_乱单张数, _乱单购买数量, _本区域张数, _本区域购买数量, _跨区域张数, _跨区域购买数量, _跨楼层张数, _跨楼层购买数量, list统计详细信息);
            }, null);
            #endregion

        }

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormHelper.GenerateTableDes(typeof(_商品明细信息), typeof(_库位明细信息));
        }
        #endregion

        /**************** common method ****************/

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

        #region ExportExcel 导出表格信息
        private void ExportExcel(double _乱单张数, double _乱单购买数量, double _本区域张数, double _本区域购买数量, double _跨区域张数, double _跨区域购买数量, double _跨楼层张数, double _跨楼层购买数量, List<_统计详细信息> detail)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];


            #region 订单分配
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 统计
                {
                    var sheet1 = workbox.Worksheets.Add("统计");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "乱单张数";
                    sheet1.Cells[1, 2].Value = "乱单购买数量";
                    sheet1.Cells[1, 3].Value = "本区域张数";
                    sheet1.Cells[1, 4].Value = "本区域购买数量";
                    sheet1.Cells[1, 5].Value = "跨区域张数";
                    sheet1.Cells[1, 6].Value = "跨区域购买数量";
                    sheet1.Cells[1, 7].Value = "跨楼层张数";
                    sheet1.Cells[1, 8].Value = "跨楼层购买数量";
                    #endregion

                    #region 数据行
                    sheet1.Cells[2, 1].Value = _乱单张数;
                    sheet1.Cells[2, 2].Value = _乱单购买数量;
                    sheet1.Cells[2, 3].Value = _本区域张数;
                    sheet1.Cells[2, 4].Value = _本区域购买数量;
                    sheet1.Cells[2, 5].Value = _跨区域张数;
                    sheet1.Cells[2, 6].Value = _跨区域购买数量;
                    sheet1.Cells[2, 7].Value = _跨楼层张数;
                    sheet1.Cells[2, 8].Value = _跨楼层购买数量;
                    #endregion
                }
                #endregion

                #region 详细表
                {
                    var sheet1 = workbox.Worksheets.Add("详细表");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "区域";
                    sheet1.Cells[1, 2].Value = "乱单张数";
                    sheet1.Cells[1, 3].Value = "乱单购买数量";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = detail.Count; idx < len; idx++)
                    {
                        var curOrder = detail[idx];
                        sheet1.Cells[rowIdx, 1].Value = curOrder._区域;
                        sheet1.Cells[rowIdx, 2].Value = curOrder._乱单张数;
                        sheet1.Cells[rowIdx, 3].Value = curOrder._乱单购买数量;
                        rowIdx++;
                    }
                    #endregion
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
                        //btnAnalyze.Enabled = true;
                    }
                }, null);
        }
        #endregion

        /**************** common class ****************/

        [ExcelTable("商品明细表")]
        class _商品明细信息
        {
            private string _Org商品明细;

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

            public List<_商品订购信息> _详细信息
            {
                get
                {
                    var list = new List<_商品订购信息>();
                    if (!string.IsNullOrEmpty(_Org商品明细) && _Org商品明细.IndexOf(";") > 0)
                    {
                        var arr = _Org商品明细.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        foreach (var item in arr)
                        {
                            var detailArr = item.Split(new string[] { "*" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                            if (detailArr.Count >= 2)
                            {
                                var model = new _商品订购信息();
                                model.SKU = detailArr[0];
                                model._数量 = Convert.ToDouble(detailArr[1]);
                                list.Add(model);
                            }
                        }
                    }
                    return list;
                }
            }
        }

        [ExcelTable("商品库位表")]
        class _库位明细信息
        {
            private string _Org库位;
            private string _OrgSku;

            [ExcelColumn("库位")]
            public string _完整库位
            {
                get
                {
                    return _Org库位;
                }
                set
                {
                    _Org库位 = value != null ? value.ToString().Trim() : "";
                }
            }

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

            public string _库位
            {
                get
                {
                    if (!string.IsNullOrEmpty(_OrgSku))
                        return _OrgSku.Substring(0, 1).ToUpper();
                    return string.Empty;
                }
            }
        }

        class _商品订购信息
        {
            public string SKU { get; set; }
            public double _数量 { get; set; }
            public string _库位 { get; set; }
        }

        class _统计详细信息
        {
            public string _区域 { get; set; }
            public double _乱单张数 { get; set; }
            public double _乱单购买数量 { get; set; }
        }
    }
}
