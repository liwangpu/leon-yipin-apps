using LinqToExcel;
using OfficeOpenXml;
using OrderAllot.Entities;
using OrderAllot.Maps;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OrderAllot.Libs;
using CommonLibs;


namespace OrderAllot
{
    public partial class _判断供应商个数 : Form
    {
        public _判断供应商个数()
        {
            InitializeComponent();

            txtUpShStore.Text = @"C:\Users\Leon\Desktop\11.19\上海所有库存.xlsx";
            txtUpKsStore.Text = @"C:\Users\Leon\Desktop\11.19\昆山所有库存.xlsx";

        }

        private void _判断供应商个数_Load(object sender, EventArgs e)
        {
            dtpTime.Value = Convert.ToDateTime("2000-01-01");
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

        #region 上传上海库存
        private void btnUpShStore_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpShStore.Text = OpenFileDialog1.FileName;
            }
        }
        #endregion

        #region 上传昆山库存
        private void btnUpKsStore_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpKsStore.Text = OpenFileDialog1.FileName;
            }
        }
        #endregion

        #region 分析计算
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            var dtFlag = dtpTime.Value;

            var _str上海所有库存path = txtUpShStore.Text;
            var _str昆山所有库存path = txtUpKsStore.Text;

            var _上海库存List = new List<_SKU供应商信息>();
            var _昆山库存List = new List<_SKU供应商信息>();
            var _导出结果List = new List<SKU供应商数量信息>();
            var _导出重复链接信息List = new List<重复链接信息>();
            var _导出无链接信息List = new List<重复链接信息>();
            var _开发List = new List<string>();

            #region 读取数据
            var act = new Action(() =>
            {
                ShowMsg("开始读取表格数据");

                #region 读取上海所有库存
                if (!string.IsNullOrEmpty(_str上海所有库存path))
                {
                    using (var excel = new ExcelQueryFactory(_str上海所有库存path))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<_SKU供应商信息>(s)
                                          where c._商品创建时间 >= dtFlag
                                          select c;
                                _上海库存List.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                    }
                }
                #endregion

                #region 读取昆山所有库存
                if (!string.IsNullOrEmpty(_str昆山所有库存path))
                {
                    using (var excel = new ExcelQueryFactory(_str昆山所有库存path))
                    {
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<_SKU供应商信息>(s)
                                          where c._商品创建时间 >= dtFlag
                                          select c;
                                _昆山库存List.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                    }
                }
                #endregion
            });
            #endregion


            #region 解析数据
            act.BeginInvoke((obj) =>
            {
                ShowMsg("开始计算表格数据");

                #region 把昆山库存去重复加入上海库存信息
                {
                    _昆山库存List.ForEach(ks =>
                    {
                        var count = _上海库存List.Count(x => x._SKU码 == ks._SKU码);
                        if (count == 0)
                            _上海库存List.Add(ks);
                    });
                }
                #endregion

                #region 获取所有开发信息
                _开发List = _上海库存List.Select(x => x._开发).Distinct().ToList();
                #endregion

                #region 计算链接等信息
                _开发List.ForEach(dvname =>
                {
                    var _零个链接的sku个数 = 0;
                    var _一个链接的sku个数 = 0;
                    var _两个链接及以上的sku个数 = 0;
                    var _重复的链接sku个数 = 0;

                    var dpInfo = new SKU供应商数量信息();
                    dpInfo._开发 = dvname;
                    var ownSkus = _上海库存List.Where(x => x._开发 == dvname).ToList();
                    dpInfo._SKU个数 = ownSkus.Count;

                    #region 计算链接个数以及重复信息
                    ownSkus.ForEach(ow =>
                    {
                        var nets = new List<string>();
                        var skus = new List<string>();
                        var bMultiple = false;
                        var bZero = false;

                        #region 收集网址信息
                        if (!string.IsNullOrEmpty(ow._网址1))
                            nets.Add(ow._网址1);
                        if (!string.IsNullOrEmpty(ow._网址2))
                            nets.Add(ow._网址2);
                        if (!string.IsNullOrEmpty(ow._网址3))
                            nets.Add(ow._网址3);
                        if (!string.IsNullOrEmpty(ow._网址4))
                            nets.Add(ow._网址4);
                        if (!string.IsNullOrEmpty(ow._网址5))
                            nets.Add(ow._网址5);
                        if (!string.IsNullOrEmpty(ow._网址6))
                            nets.Add(ow._网址6);
                        #endregion

                        #region 计算零个链接的个数
                        if (nets.Count == 0)
                        {
                            _零个链接的sku个数++;
                            bZero = true;
                        }
                        #endregion

                        #region 计算一个链接的个数
                        if (nets.Count == 1)
                        {
                            _一个链接的sku个数++;
                        }
                        #endregion

                        #region 计算两个链接以上的个数
                        if (nets.Count >= 2)
                        {
                            _两个链接及以上的sku个数++;
                            skus = nets.Distinct().ToList();
                        }
                        #endregion

                        #region 计算重复链接的个数
                        if (skus.Count >= 2)
                        {
                            foreach (var sss in skus)
                            {
                                var __count = nets.Count(nt => nt == sss);
                                if (__count >= 2)
                                {
                                    _重复的链接sku个数++;
                                    bMultiple = true;
                                    break;
                                }
                            }
                        }
                        #endregion

                        #region 记录下无网址信息
                        if (bZero)
                        {
                            var mult = new 重复链接信息();
                            mult._开发 = dvname;
                            mult._SKU = ow._SKU码;
                            mult._网址1 = ow._网址1;
                            mult._网址2 = ow._网址2;
                            mult._网址3 = ow._网址3;
                            mult._网址4 = ow._网址4;
                            mult._网址5 = ow._网址5;
                            mult._网址6 = ow._网址6;
                            mult._商品创建时间 = ow._商品创建时间;
                            _导出无链接信息List.Add(mult);
                        }
                        #endregion

                        #region 记录下重复网址信息
                        if (bMultiple)
                        {
                            var mult = new 重复链接信息();
                            mult._开发 = dvname;
                            mult._SKU = ow._SKU码;
                            mult._网址1 = ow._网址1;
                            mult._网址2 = ow._网址2;
                            mult._网址3 = ow._网址3;
                            mult._网址4 = ow._网址4;
                            mult._网址5 = ow._网址5;
                            mult._网址6 = ow._网址6;
                            mult._商品创建时间 = ow._商品创建时间;
                            _导出重复链接信息List.Add(mult);
                        }
                        #endregion
                    });
                    #endregion

                    dpInfo._零个链接 = _零个链接的sku个数;
                    dpInfo._一个链接 = _一个链接的sku个数;
                    dpInfo._两个链接 = _两个链接及以上的sku个数;
                    dpInfo._重复链接 = _重复的链接sku个数;
                    _导出结果List.Add(dpInfo);
                });
                #endregion




                ExportExcel(_导出结果List, _导出重复链接信息List, _导出无链接信息List);

            }, null);
            #endregion
        }
        #endregion

        #region ExportExcel 导出Excel表格
        private void ExportExcel(List<SKU供应商数量信息> list开发产品信息, List<重复链接信息> list重复链接信息, List<重复链接信息> list无链接信息)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];
            var buffer2 = new byte[0];
            var buffer3 = new byte[0];

            #region 汇总
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 信息汇总
                {
                    var sheet1 = workbox.Worksheets.Add("汇总");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "开发";
                    sheet1.Cells[1, 2].Value = "SKU个数";
                    sheet1.Cells[1, 3].Value = "无链接";
                    sheet1.Cells[1, 4].Value = "一个链接";
                    sheet1.Cells[1, 5].Value = "两个及以上链接";
                    sheet1.Cells[1, 6].Value = "重复链接";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list开发产品信息.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list开发产品信息[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._开发;
                        sheet1.Cells[rowIdx, 2].Value = info._SKU个数;
                        sheet1.Cells[rowIdx, 3].Value = info._零个链接;
                        sheet1.Cells[rowIdx, 4].Value = info._一个链接;
                        sheet1.Cells[rowIdx, 5].Value = info._两个链接;
                        sheet1.Cells[rowIdx, 6].Value = info._重复链接;
                    }
                    #endregion
                }
                #endregion

                #region 重复链接
                {
                    var sheet2 = workbox.Worksheets.Add("重复链接详情");

                    #region 标题行
                    sheet2.Cells[1, 1].Value = "开发";
                    sheet2.Cells[1, 2].Value = "SKU";
                    sheet2.Cells[1, 3].Value = "网址";
                    sheet2.Cells[1, 4].Value = "网址2";
                    sheet2.Cells[1, 5].Value = "网址3";
                    sheet2.Cells[1, 6].Value = "网址4";
                    sheet2.Cells[1, 7].Value = "网址5";
                    sheet2.Cells[1, 8].Value = "网址6";
                    sheet2.Cells[1, 9].Value = "开发时间";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list重复链接信息.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list重复链接信息[idx];
                        sheet2.Cells[rowIdx, 1].Value = info._开发;
                        sheet2.Cells[rowIdx, 2].Value = info._SKU;
                        sheet2.Cells[rowIdx, 3].Value = info._网址1;
                        sheet2.Cells[rowIdx, 4].Value = info._网址2;
                        sheet2.Cells[rowIdx, 5].Value = info._网址3;
                        sheet2.Cells[rowIdx, 6].Value = info._网址4;
                        sheet2.Cells[rowIdx, 7].Value = info._网址5;
                        sheet2.Cells[rowIdx, 8].Value = info._网址6;
                        sheet2.Cells[rowIdx, 9].Value = info._商品创建时间.ToString("yyyy-MM-dd");
                    }
                    #endregion
                }
                #endregion

                #region 无链接
                {
                    var sheet3 = workbox.Worksheets.Add("无链接详情");

                    #region 标题行
                    sheet3.Cells[1, 1].Value = "开发";
                    sheet3.Cells[1, 2].Value = "SKU";
                    sheet3.Cells[1, 3].Value = "网址";
                    sheet3.Cells[1, 4].Value = "网址2";
                    sheet3.Cells[1, 5].Value = "网址3";
                    sheet3.Cells[1, 6].Value = "网址4";
                    sheet3.Cells[1, 7].Value = "网址5";
                    sheet3.Cells[1, 8].Value = "网址6";
                    sheet3.Cells[1, 9].Value = "开发时间";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list无链接信息.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list无链接信息[idx];
                        sheet3.Cells[rowIdx, 1].Value = info._开发;
                        sheet3.Cells[rowIdx, 2].Value = info._SKU;
                        sheet3.Cells[rowIdx, 3].Value = info._网址1;
                        sheet3.Cells[rowIdx, 4].Value = info._网址2;
                        sheet3.Cells[rowIdx, 5].Value = info._网址3;
                        sheet3.Cells[rowIdx, 6].Value = info._网址4;
                        sheet3.Cells[rowIdx, 7].Value = info._网址5;
                        sheet3.Cells[rowIdx, 8].Value = info._网址6;
                        sheet3.Cells[rowIdx, 9].Value = info._商品创建时间.ToString("yyyy-MM-dd");
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

                    txtExport.Text = FileName;
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
                    btnAnalyze.Enabled = false;
                }
            }, null);
        }
        #endregion


    }
}
