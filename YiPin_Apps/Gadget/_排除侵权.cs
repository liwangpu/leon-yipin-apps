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
    public partial class _排除侵权 : Form
    {
        private const string _const侵权情况标记 = "侵权";
        private const int _const建议备货明细 = 0;
        private const int _const侵权详情 = 1;
        public _排除侵权()
        {
            InitializeComponent();
        }

        private void _排除侵权_Load(object sender, EventArgs e)
        {
            //txt侵权产品.Text = @"C:\Users\Leon\Desktop\侵权\侵权ep_product.csv";
            //txt产品信息.Text = @"C:\Users\Leon\Desktop\商品信息所有.csv";
            //txt上上月销量.Text = @"C:\Users\Leon\Desktop\aaa - 副本\10月份销量.csv";
            //txt上月销量.Text = @"C:\Users\Leon\Desktop\aaa - 副本\11月销量.csv";
            //txt当月上半月销量.Text = @"C:\Users\Leon\Desktop\aaa - 副本\12月上半个月.csv";
            //txt当月下半月销量.Text = @"C:\Users\Leon\Desktop\aaa - 副本\12月下半个月.csv";
        }

        /**************** button event ****************/

        #region 上传侵权产品情况
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

        #region 上传上上月销量情况
        private void btn上上月销量_Click(object sender, EventArgs e)
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
                    txt上上月销量.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传上月销量情况
        private void btn上月销量_Click(object sender, EventArgs e)
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
                    txt上月销量.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传当月上半月销量情况
        private void btn当月上半月销量_Click(object sender, EventArgs e)
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
                    txt当月上半月销量.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传当月下半月销量情况
        private void btn当月下半月销量_Click(object sender, EventArgs e)
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
                    txt当月下半月销量.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传商品信息
        private void btn产品信息_Click(object sender, EventArgs e)
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
                    txt产品信息.Text = OpenFileDialog1.FileName;

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
            btnAnalyze.Enabled = false;
            var d销量差 = Convert.ToInt32(nup销量差.Value);
            var d销量天数 = Convert.ToInt32(nup销量天数.Value);
            var d备货天数 = Convert.ToInt32(nup备货天数.Value);
            var list各平台近期销量 = new List<_各平台近期销量>();
            var list侵权情况 = new List<_侵权情况>();
            var list产品信息 = new List<_产品信息>();

            #region 读取数据
            var actReadData = new Action(() =>
            {
                ShowMsg("开始读取侵权数据");
                #region 读取侵权产品
                {
                    var strCsvPath = txt侵权产品.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_侵权情况>()
                                          select c;
                                list侵权情况.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                ShowMsg("开始读取上上月销量数据");
                #region 读取上上月销量
                {
                    var strCsvPath = txt上上月销量.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = (from c in csv.Worksheet<_各平台近期销量>()
                                           select c).ToList();


                                for (int i = 0, len = tmp.Count; i < len; i++)
                                {
                                    var curItem = tmp[i];
                                    curItem._所在月份 = _Enum所在月份._上上月;
                                }

                                list各平台近期销量.AddRange(tmp);

                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }

                }
                #endregion

                ShowMsg("开始读取上月销量数据");
                #region 读取上月销量
                {
                    var strCsvPath = txt上月销量.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = (from c in csv.Worksheet<_各平台近期销量>()
                                           select c).ToList();


                                for (int i = 0, len = tmp.Count; i < len; i++)
                                {
                                    var curItem = tmp[i];
                                    curItem._所在月份 = _Enum所在月份._上月;
                                }

                                list各平台近期销量.AddRange(tmp);

                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                ShowMsg("开始读取当月上半月销量数据");
                #region 读取当月上半月销量
                {
                    var strCsvPath = txt当月上半月销量.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = (from c in csv.Worksheet<_各平台近期销量>()
                                           select c).ToList();


                                for (int i = 0, len = tmp.Count; i < len; i++)
                                {
                                    var curItem = tmp[i];
                                    curItem._所在月份 = _Enum所在月份._当月上半月;
                                }

                                list各平台近期销量.AddRange(tmp);

                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                ShowMsg("开始读取当月下半月销量数据");
                #region 读取当月上半月销量
                {
                    var strCsvPath = txt当月下半月销量.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = (from c in csv.Worksheet<_各平台近期销量>()
                                           select c).ToList();


                                for (int i = 0, len = tmp.Count; i < len; i++)
                                {
                                    var curItem = tmp[i];
                                    curItem._所在月份 = _Enum所在月份._当月下半月;
                                }

                                list各平台近期销量.AddRange(tmp);

                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                ShowMsg("开始读取产品信息数据");
                #region 读取产品信息
                {
                    var strCsvPath = txt产品信息.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_产品信息>()
                                          select c;
                                list产品信息.AddRange(tmp);
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

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                ShowMsg("读取完毕,开始处理数据");
                var list备货信息 = new List<_备货信息>();

                #region 统计各平台销售情况
                var list平台销量SKU = list各平台近期销量.Select(x => x.SKU).Distinct().OrderBy(x => x).ToList();
                list平台销量SKU.ForEach(strSKU =>
                {
                    var model = new _备货信息();
                    model.SKU = strSKU;
                    model._销量天数 = d销量天数;
                    model._备货天数 = d备货天数;

                    var _ref销量信息 = list各平台近期销量.Where(x => x.SKU == strSKU).ToList();

                    model._总销量 = _ref销量信息.Select(x => x._总销量).Sum();

                    #region 判断计算各平台侵权销量
                    var _侵权情况 = list侵权情况.Where(x => x.SKU == strSKU).FirstOrDefault();
                    if (_侵权情况 != null)
                    {
                        if (_侵权情况._wish情况 == (int)_Enum侵权情况._侵权)
                        {
                            model._wish情况 = _const侵权情况标记;
                            model._wish侵权销量 = _ref销量信息.Select(x => x._wish销量).Sum();
                        }

                        if (_侵权情况._Ebay情况 == (int)_Enum侵权情况._侵权)
                        {
                            model._Ebay情况 = _const侵权情况标记;
                            model._Ebay侵权销量 = _ref销量信息.Select(x => x._Ebay销量).Sum();
                        }

                        if (_侵权情况._SMT情况 == (int)_Enum侵权情况._侵权)
                        {
                            model._SMT情况 = _const侵权情况标记;
                            model._SMT侵权销量 = _ref销量信息.Select(x => x._SMT销量).Sum();
                        }

                        if (_侵权情况._Amazon情况 == (int)_Enum侵权情况._侵权)
                        {
                            model._Amazon情况 = _const侵权情况标记;
                            model._Amazon侵权销量 = _ref销量信息.Select(x => x._Amazon销量).Sum();
                        }

                        if (_侵权情况._Shopee情况 == (int)_Enum侵权情况._侵权)
                        {
                            model._Shopee情况 = _const侵权情况标记;
                            model._Shopee侵权销量 = _ref销量信息.Select(x => x._Shopee销量).Sum();
                        }

                        //if (_侵权情况._Joom情况 == (int)_Enum侵权情况._侵权)
                        //{
                        //    model._Joom情况 = _const侵权情况标记;
                        //    model._Joom侵权销量 = _ref销量信息.Select(x => x._Joom销量).Sum();
                        //}

                    }

                    #endregion

                    #region 匹配现有库存
                    {
                        var _ref现有库存Item = list产品信息.Where(x => x.SKU == strSKU).FirstOrDefault();
                        if (_ref现有库存Item != null)
                        {
                            model._现有库存 = _ref现有库存Item._目前库存;
                            model._供应商 = _ref现有库存Item._供应商;
                            model._采购员 = _ref现有库存Item._采购;
                            model._制单员 = _ref现有库存Item._采购;
                            model._商品成本单价 = _ref现有库存Item._商品成本单价;
                        }

                    }
                    #endregion

                    #region 统计该SKU上下半月的销量等的差
                    {
                        var _上上月Item = _ref销量信息.Where(x => x._所在月份 == _Enum所在月份._上上月).FirstOrDefault();
                        var _上月Item = _ref销量信息.Where(x => x._所在月份 == _Enum所在月份._上月).FirstOrDefault();
                        var _上半月Item = _ref销量信息.Where(x => x._所在月份 == _Enum所在月份._当月上半月).FirstOrDefault();
                        var _下半月Item = _ref销量信息.Where(x => x._所在月份 == _Enum所在月份._当月下半月).FirstOrDefault();

                        if (_上半月Item != null && _下半月Item != null)
                        {
                            var dif = _下半月Item._总销量 - _上半月Item._总销量;
                            if (dif > d销量差 || dif < -d销量差)
                            {
                                model._下半月销量 = _下半月Item._总销量;
                                model._上半月销量 = _上半月Item._总销量;
                                model._上升下降情况 = dif > d销量差 ? "上升" : "下降";

                                if (model._总侵权 < d销量差)
                                {
                                    model._表格类型 = _const侵权详情;
                                }
                            }
                        }

                        if (_上上月Item != null)
                        {
                            model._上上月销量 = _上上月Item._总销量;
                        }

                        if (_上月Item != null)
                        {
                            model._上月销量 = _上月Item._总销量;
                        }

                    }
                    #endregion


                    list备货信息.Add(model);


                });
                #endregion

                Export(list备货信息.Where(x => x._总侵权 > 0).OrderByDescending(x => x._总销量).ToList()
                    , list备货信息.Where(x => x._表格类型 == _const侵权详情).OrderByDescending(x => x._销量差).ToList(), list备货信息.OrderByDescending(x => x._总销量).ToList());


            }, null);
            #endregion
        }
        #endregion

        #region 导出表格说明
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_各平台近期销量), typeof(_侵权情况), typeof(_产品信息));

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

        #region Export导出结果
        private void Export(List<_备货信息> list侵权详细信息, List<_备货信息> list销量差详细信息, List<_备货信息> list建议备货)
        {
            ShowMsg("计算完毕,开始生成表格");
            var buffer1 = new byte[0];

            #region 生成表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 侵权详细信息
                {
                    var sheet1 = workbox.Worksheets.Add("侵权详细信息");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "总销量";
                    sheet1.Cells[1, 3].Value = "wish情况";
                    sheet1.Cells[1, 4].Value = "wish侵权销量";
                    sheet1.Cells[1, 5].Value = "Ebay情况";
                    sheet1.Cells[1, 6].Value = "Ebay侵权销量";
                    sheet1.Cells[1, 7].Value = "SMT情况";
                    sheet1.Cells[1, 8].Value = "SMT侵权销量";
                    sheet1.Cells[1, 9].Value = "Amazon情况";
                    sheet1.Cells[1, 10].Value = "Amazon侵权销量";
                    sheet1.Cells[1, 11].Value = "Shopee情况";
                    sheet1.Cells[1, 12].Value = "Shopee侵权销量";
                    sheet1.Cells[1, 13].Value = "Joom情况";
                    sheet1.Cells[1, 14].Value = "Joom侵权销量";
                    sheet1.Cells[1, 15].Value = "侵权总销量";
                    //sheet1.Cells[1, 16].Value = "销量天数";
                    //sheet1.Cells[1, 17].Value = "备货天数";
                    //sheet1.Cells[1, 18].Value = "备货数量";

                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list侵权详细信息.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list侵权详细信息[idx];
                        sheet1.Cells[rowIdx, 1].Value = info.SKU;
                        sheet1.Cells[rowIdx, 2].Value = info._总销量;
                        sheet1.Cells[rowIdx, 3].Value = info._wish情况;
                        sheet1.Cells[rowIdx, 4].Value = info._wish侵权销量;
                        sheet1.Cells[rowIdx, 5].Value = info._Ebay情况;
                        sheet1.Cells[rowIdx, 6].Value = info._Ebay侵权销量;
                        sheet1.Cells[rowIdx, 7].Value = info._SMT情况;
                        sheet1.Cells[rowIdx, 8].Value = info._SMT侵权销量;
                        sheet1.Cells[rowIdx, 9].Value = info._Amazon情况;
                        sheet1.Cells[rowIdx, 10].Value = info._Amazon侵权销量;
                        sheet1.Cells[rowIdx, 11].Value = info._Shopee情况;
                        sheet1.Cells[rowIdx, 12].Value = info._Shopee侵权销量;
                        sheet1.Cells[rowIdx, 13].Value = info._Joom情况;
                        sheet1.Cells[rowIdx, 14].Value = info._Joom侵权销量;
                        sheet1.Cells[rowIdx, 15].Value = info._总侵权;
                        //sheet1.Cells[rowIdx, 16].Value = info._销量天数;
                        //sheet1.Cells[rowIdx, 17].Value = info._备货天数;
                        //sheet1.Cells[rowIdx, 18].Value = info._建议备货;
                    }
                    #endregion

                }
                #endregion

                #region 销量差详细信息
                {
                    var sheet1 = workbox.Worksheets.Add("销量差详细信息");


                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "wish情况";
                    sheet1.Cells[1, 3].Value = "wish侵权销量";
                    sheet1.Cells[1, 4].Value = "Ebay情况";
                    sheet1.Cells[1, 5].Value = "Ebay侵权销量";
                    sheet1.Cells[1, 6].Value = "SMT情况";
                    sheet1.Cells[1, 7].Value = "SMT侵权销量";
                    sheet1.Cells[1, 8].Value = "Amazon情况";
                    sheet1.Cells[1, 9].Value = "Amazon侵权销量";
                    sheet1.Cells[1, 10].Value = "Shopee情况";
                    sheet1.Cells[1, 11].Value = "Shopee侵权销量";
                    sheet1.Cells[1, 12].Value = "Joom情况";
                    sheet1.Cells[1, 13].Value = "Joom侵权销量";
                    sheet1.Cells[1, 14].Value = "侵权总销量";

                    sheet1.Cells[1, 15].Value = "总销量";
                    sheet1.Cells[1, 16].Value = "现有库存";
                    sheet1.Cells[1, 17].Value = "平均日销量";
                    sheet1.Cells[1, 18].Value = "销量天数";
                    sheet1.Cells[1, 19].Value = "备货天数";
                    sheet1.Cells[1, 20].Value = "建议备货数量";

                    sheet1.Cells[1, 21].Value = "下半月销量";
                    sheet1.Cells[1, 22].Value = "上半月销量";
                    sheet1.Cells[1, 23].Value = "变动";
                    sheet1.Cells[1, 24].Value = "销量差";


                    sheet1.Cells[1, 25].Value = "上上月销量";
                    sheet1.Cells[1, 26].Value = "上月销量";
                    sheet1.Cells[1, 27].Value = "当月销量";

                    sheet1.Cells[1, 28].Value = "供应商";
                    sheet1.Cells[1, 29].Value = "采购员";
                    sheet1.Cells[1, 30].Value = "制单员";
                    sheet1.Cells[1, 31].Value = "含税单价";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list销量差详细信息.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list销量差详细信息[idx];
                        sheet1.Cells[rowIdx, 1].Value = info.SKU;
                        sheet1.Cells[rowIdx, 2].Value = info._wish情况;
                        sheet1.Cells[rowIdx, 3].Value = info._wish侵权销量;
                        sheet1.Cells[rowIdx, 4].Value = info._Ebay情况;
                        sheet1.Cells[rowIdx, 5].Value = info._Ebay侵权销量;
                        sheet1.Cells[rowIdx, 6].Value = info._SMT情况;
                        sheet1.Cells[rowIdx, 7].Value = info._SMT侵权销量;
                        sheet1.Cells[rowIdx, 8].Value = info._Amazon情况;
                        sheet1.Cells[rowIdx, 9].Value = info._Amazon侵权销量;
                        sheet1.Cells[rowIdx, 10].Value = info._Shopee情况;
                        sheet1.Cells[rowIdx, 11].Value = info._Shopee侵权销量;
                        sheet1.Cells[rowIdx, 12].Value = info._Joom情况;
                        sheet1.Cells[rowIdx, 13].Value = info._Joom侵权销量;
                        sheet1.Cells[rowIdx, 14].Value = info._总侵权;

                        sheet1.Cells[rowIdx, 15].Value = info._总销量;
                        sheet1.Cells[rowIdx, 16].Value = info._现有库存;
                        sheet1.Cells[rowIdx, 17].Value = info._平均日销量;
                        sheet1.Cells[rowIdx, 18].Value = info._销量天数;
                        sheet1.Cells[rowIdx, 19].Value = info._备货天数;
                        sheet1.Cells[rowIdx, 20].Value = info._建议备货;

                        sheet1.Cells[rowIdx, 21].Value = info._下半月销量;
                        sheet1.Cells[rowIdx, 22].Value = info._上半月销量;
                        sheet1.Cells[rowIdx, 23].Value = info._上升下降情况;
                        sheet1.Cells[rowIdx, 24].Value = info._销量差;

                        sheet1.Cells[rowIdx, 25].Value = info._上上月销量;
                        sheet1.Cells[rowIdx, 26].Value = info._上月销量;
                        sheet1.Cells[rowIdx, 27].Value= info._当月销量;

                        sheet1.Cells[rowIdx, 28].Value = info._供应商;
                        sheet1.Cells[rowIdx, 29].Value = info._采购员;
                        sheet1.Cells[rowIdx, 30].Value = info._制单员;
                        sheet1.Cells[rowIdx, 31].Value = info._商品成本单价;
                    }
                    #endregion

                }
                #endregion

                #region 建议备货明细
                {
                    var sheet1 = workbox.Worksheets.Add("建议备货明细");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "wish情况";
                    sheet1.Cells[1, 3].Value = "wish侵权销量";
                    sheet1.Cells[1, 4].Value = "Ebay情况";
                    sheet1.Cells[1, 5].Value = "Ebay侵权销量";
                    sheet1.Cells[1, 6].Value = "SMT情况";
                    sheet1.Cells[1, 7].Value = "SMT侵权销量";
                    sheet1.Cells[1, 8].Value = "Amazon情况";
                    sheet1.Cells[1, 9].Value = "Amazon侵权销量";
                    sheet1.Cells[1, 10].Value = "Shopee情况";
                    sheet1.Cells[1, 11].Value = "Shopee侵权销量";
                    sheet1.Cells[1, 12].Value = "Joom情况";
                    sheet1.Cells[1, 13].Value = "Joom侵权销量";
                    sheet1.Cells[1, 14].Value = "侵权总销量";

                    sheet1.Cells[1, 15].Value = "总销量";
                    sheet1.Cells[1, 16].Value = "现有库存";
                    sheet1.Cells[1, 17].Value = "平均日销量";
                    sheet1.Cells[1, 18].Value = "销量天数";
                    sheet1.Cells[1, 19].Value = "备货天数";
                    sheet1.Cells[1, 20].Value = "建议备货数量";

                    sheet1.Cells[1, 21].Value = "供应商";
                    sheet1.Cells[1, 22].Value = "采购员";
                    sheet1.Cells[1, 23].Value = "制单员";
                    sheet1.Cells[1, 24].Value = "含税单价";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list建议备货.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list建议备货[idx];
                        sheet1.Cells[rowIdx, 1].Value = info.SKU;
                        sheet1.Cells[rowIdx, 2].Value = info._wish情况;
                        sheet1.Cells[rowIdx, 3].Value = info._wish侵权销量;
                        sheet1.Cells[rowIdx, 4].Value = info._Ebay情况;
                        sheet1.Cells[rowIdx, 5].Value = info._Ebay侵权销量;
                        sheet1.Cells[rowIdx, 6].Value = info._SMT情况;
                        sheet1.Cells[rowIdx, 7].Value = info._SMT侵权销量;
                        sheet1.Cells[rowIdx, 8].Value = info._Amazon情况;
                        sheet1.Cells[rowIdx, 9].Value = info._Amazon侵权销量;
                        sheet1.Cells[rowIdx, 10].Value = info._Shopee情况;
                        sheet1.Cells[rowIdx, 11].Value = info._Shopee侵权销量;
                        sheet1.Cells[rowIdx, 12].Value = info._Joom情况;
                        sheet1.Cells[rowIdx, 13].Value = info._Joom侵权销量;
                        sheet1.Cells[rowIdx, 14].Value = info._总侵权;

                        sheet1.Cells[rowIdx, 15].Value = info._总销量;
                        sheet1.Cells[rowIdx, 16].Value = info._现有库存;
                        sheet1.Cells[rowIdx, 17].Value = info._平均日销量;
                        sheet1.Cells[rowIdx, 18].Value = info._销量天数;
                        sheet1.Cells[rowIdx, 19].Value = info._备货天数;
                        sheet1.Cells[rowIdx, 20].Value = info._建议备货;

                        sheet1.Cells[rowIdx, 21].Value = info._供应商;
                        sheet1.Cells[rowIdx, 22].Value = info._采购员;
                        sheet1.Cells[rowIdx, 23].Value = info._制单员;
                        sheet1.Cells[rowIdx, 24].Value = info._商品成本单价;
                    }
                    #endregion

                }
                #endregion


                buffer1 = package.GetAsByteArray();
            }
            #endregion

            #region 导出表格对话框
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

        [ExcelTable("各平台近期销量表")]
        class _各平台近期销量
        {
            private string orgSKU;

            [ExcelColumn("商品sku")]
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

            [ExcelColumn("业绩归属人")]
            public string _开发 { get; set; }

            [ExcelColumn("销量（已发货并且扣了库存的销量）")]
            public decimal _总销量 { get; set; }

            [ExcelColumn("wish")]
            public decimal _wish销量 { get; set; }

            [ExcelColumn("Ebay")]
            public decimal _Ebay销量 { get; set; }

            [ExcelColumn("SMT")]
            public decimal _SMT销量 { get; set; }

            [ExcelColumn("Amazon")]
            public decimal _Amazon销量 { get; set; }

            [ExcelColumn("Shopee")]
            public decimal _Shopee销量 { get; set; }

            [ExcelColumn("Joom")]
            public decimal _Joom销量 { get; set; }

            public _Enum所在月份 _所在月份 { get; set; }

        }

        [ExcelTable("侵权情况表")]
        class _侵权情况
        {
            private string orgSKU;

            [ExcelColumn("parentsku")]
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

            [ExcelColumn("归属开发")]
            public string _开发 { get; set; }

            [ExcelColumn("wish")]
            public int _wish情况 { get; set; }

            [ExcelColumn("ebay")]
            public int _Ebay情况 { get; set; }

            [ExcelColumn("smt")]
            public int _SMT情况 { get; set; }

            [ExcelColumn("amazon")]
            public int _Amazon情况 { get; set; }

            [ExcelColumn("shopee")]
            public int _Shopee情况 { get; set; }

            [ExcelColumn("joom")]
            public int _Joom情况 { get; set; }

            [ExcelColumn("在售")]
            public int _在售情况 { get; set; }

        }

        [ExcelTable("产品信息表")]
        class _产品信息
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

            [ExcelColumn("业绩归属2")]
            public string _开发 { get; set; }

            [ExcelColumn("供应商")]
            public string _供应商 { get; set; }

            [ExcelColumn("采购员")]
            public string _采购 { get; set; }

            [ExcelColumn("商品成本单价")]
            public decimal _商品成本单价 { get; set; }

            [ExcelColumn("可用数量")]
            public decimal _可用数量 { get; set; }

            [ExcelColumn("缺货及未派单数量")]
            public decimal _缺货及未派单数量 { get; set; }

            [ExcelColumn("采购未入库")]
            public decimal _采购未入库 { get; set; }

            public decimal _目前库存
            {
                get
                {
                    return _可用数量 - _缺货及未派单数量 + _采购未入库;
                }
            }


        }

        class _备货信息
        {
            public string SKU { get; set; }

            public string _供应商 { get; set; }

            public string _采购员 { get; set; }

            public string _制单员 { get; set; }

            public decimal _商品成本单价 { get; set; }

            public decimal _现有库存 { get; set; }

            public decimal _总销量 { get; set; }

            public int _销量天数 { get; set; }

            public int _备货天数 { get; set; }

            public string _wish情况 { get; set; }

            public decimal _wish侵权销量 { get; set; }

            public string _Ebay情况 { get; set; }

            public decimal _Ebay侵权销量 { get; set; }

            public string _SMT情况 { get; set; }

            public decimal _SMT侵权销量 { get; set; }

            public string _Amazon情况 { get; set; }

            public decimal _Amazon侵权销量 { get; set; }
            public string _Shopee情况 { get; set; }

            public decimal _Shopee侵权销量 { get; set; }

            public string _Joom情况 { get; set; }

            public decimal _Joom侵权销量 { get; set; }

            public decimal _总侵权
            {
                get
                {
                    return _wish侵权销量 + _Ebay侵权销量 + _SMT侵权销量 + _Amazon侵权销量 + _Shopee侵权销量 + _Joom侵权销量;
                }
            }

            public decimal _平均日销量
            {
                get
                {
                    decimal _平均日销量 = Math.Round((_总销量 - _总侵权) / _销量天数, 0);
                    return _平均日销量;
                }
            }

            public decimal _建议备货
            {
                get
                {
                    var res = _备货天数 * _平均日销量 - _现有库存;
                    return res;

                }
            }

            public decimal _上半月销量 { get; set; }

            public decimal _下半月销量 { get; set; }

            public decimal _上上月销量 { get; set; }

            public decimal _上月销量 { get; set; }

            public decimal _当月销量
            {
                get
                {
                    return _上半月销量 + _下半月销量;
                }
            }

            public string _上升下降情况 { get; set; }

            public decimal _销量差
            {
                get
                {
                    return _下半月销量 - _上半月销量 > 0 ? _下半月销量 - _上半月销量 : _上半月销量 - _下半月销量;
                }
            }

            public int _表格类型 { get; set; }

        }

        enum _Enum侵权情况
        {
            _侵权 = 0,
            _不侵权 = 1

            //平台下对应的1是不侵权，0是侵权不可刊登
        }

        enum _Enum在售情况
        {
            _全停售 = 0,
            _全在售 = 1,
            _部分子SKU在售 = 2

            //status是商品在售状态，0代表全停售，1代表全在售，2代表部分子sku在售
        }

        enum _Enum所在月份
        {
            _上上月 = 1,
            _上月 = 2,
            _当月上半月 = 3,
            _当月下半月 = 4
        }



    }
}
