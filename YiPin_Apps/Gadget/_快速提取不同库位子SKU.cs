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
    public partial class _快速提取不同库位子SKU : Form
    {
        public _快速提取不同库位子SKU()
        {
            InitializeComponent();
        }

        private void _快速提取不同库位子SKU_Load(object sender, EventArgs e)
        {
            //txtFile.Text = @"C:\Users\Leon\Desktop\在售SKU5月24号.csv";
            //txtFloor.Text = @"C:\Users\Leon\Desktop\区域对应楼层.csv";
        }

        /**************** button event ****************/

        #region 上传SKU数据
        private void BtnUpload_Click(object sender, EventArgs e)
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
                    txtFile.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 上传楼层数据
        private void BtnUploadFloor_Click(object sender, EventArgs e)
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
                    txtFloor.Text = OpenFileDialog1.FileName;

                }
                else
                {
                    MessageBox.Show("csv文件名称不规范,请去掉文件名称中的特殊字符如\".\"等", "温馨提示");
                }
            }
        }
        #endregion

        #region 处理按钮事件
        private void BtnCalcu_Click(object sender, EventArgs e)
        {
            var _list楼层区域 = new List<_楼层区域表Mapping>();
            var _list在售SKU = new List<_在售SKUMapping>();
            ShowMsg("开始读取数据");

            #region 读取数据
            var actReadData = new Action(() =>
            {
                #region 读取SKU
                {
                    var strCsvPath = txtFile.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_在售SKUMapping>()
                                          select c;
                                _list在售SKU.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        }
                    }
                }
                #endregion

                #region 读取楼层
                {
                    var strCsvPath = txtFloor.Text;
                    if (!string.IsNullOrEmpty(strCsvPath))
                    {
                        using (var csv = new ExcelQueryFactory(strCsvPath))
                        {
                            try
                            {
                                var tmp = from c in csv.Worksheet<_楼层区域表Mapping>()
                                          select c;
                                _list楼层区域.AddRange(tmp);
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

            ShowMsg("开始计算数据");

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                var _list异常在售SKU = new List<_在售SKUMapping>();

                var _list楼层异常SKU = _list楼层区域.Select(x => x.Floor).Distinct().Select(x => new _楼层异常SKU { _Floor = x }).ToList();

                var _list异常SKU组 = new List<_SKU组>();

                var allGroup = _list在售SKU.Select(x => x.ParentSKU).Distinct().Select(x => new _SKU组 { ParentSKU = x }).ToList();

                //sku分组
                for (var idx = _list在售SKU.Count - 1; idx >= 0; idx--)
                {
                    var item = _list在售SKU[idx];
                    var refGroup = allGroup.Where(x => x.ParentSKU == item.ParentSKU).First();
                    var floorItem = _list楼层区域.Where(x => x.Area == item.Area).FirstOrDefault();
                    if (floorItem != null)
                        item.Floor = floorItem.Floor;
                    refGroup.SKUs.Add(item);
                    //添加区域信息
                    var hasThisItemArea = refGroup.Areas.Any(x => x == item.Area);
                    if (!hasThisItemArea)
                        refGroup.Areas.Add(item.Area);
                    //添加楼层信息
                    if (floorItem != null)
                    {
                        var hasThisItemFloor = refGroup.Floors.Any(x => x == floorItem.Floor);
                        if (!hasThisItemFloor)
                            refGroup.Floors.Add(floorItem.Floor);
                    }
                }

                //移除正常sku组
                for (var idx = allGroup.Count - 1; idx >= 0; idx--)
                {
                    //判断是否是异常sku
                    var group = allGroup[idx];
                    var onlyOneArea = group.Areas.Count <= 1;
                    if (onlyOneArea)
                    {
                        allGroup.RemoveAt(idx);
                    }
                    else
                    {
                        _list异常在售SKU.AddRange(group.SKUs);
                    }

                }

                //现在allGroup留下的都是异常的sku了,sku单楼层标记
                for (int idx = allGroup.Count - 1; idx >= 0; idx--)
                {
                    var group = allGroup[idx];
                    for (int iix = group.SKUs.Count - 1; iix >= 0; iix--)
                    {
                        var skuItem = group.SKUs[iix];
                        skuItem.SingleFloor = group.SKUs.Count(x => x.Floor == skuItem.Floor) == 1;
                    }
                }


                Export(_list异常在售SKU, allGroup, _list楼层区域.Select(x => x.Floor).Distinct().ToList());
            }, null);
            #endregion
        }
        #endregion


        #region 导出表格说明
        private void LkDecs_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_在售SKUMapping), typeof(_楼层区域表Mapping));

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

        #region Export 导出表
        private void Export(List<_在售SKUMapping> _list异常在售SKU, List<_SKU组> _list异常SKU组, List<string> _listFloors)
        {
            ShowMsg("开始生成表格");
            var buffer1 = new byte[0];

            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;
                var _单楼层异常SKUs = new List<_在售SKUMapping>();


                #region 在售异常SKU
                {
                    var sheet1 = workbox.Worksheets.Add("所有异常SKU库位信息");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "父SKU";
                    sheet1.Cells[1, 2].Value = "子SKU";
                    sheet1.Cells[1, 3].Value = "库位";
                    sheet1.Cells[1, 4].Value = "区域";
                    sheet1.Cells[1, 5].Value = "楼层";
                    sheet1.Column(1).Width = 16;
                    sheet1.Column(2).Width = 16;
                    sheet1.Column(3).Width = 16;
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = _list异常在售SKU.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = _list异常在售SKU[idx];
                        sheet1.Cells[rowIdx, 1].Value = info.ParentSKU;
                        sheet1.Cells[rowIdx, 2].Value = info.ChildSKU;
                        sheet1.Cells[rowIdx, 3].Value = info.Store;
                        sheet1.Cells[rowIdx, 4].Value = info.Area;
                        sheet1.Cells[rowIdx, 5].Value = info.Floor;
                    }
                    #endregion

                }
                #endregion

                #region 异常sku组
                {
                    var sheet1 = workbox.Worksheets.Add("异常SKU概况");

                    sheet1.Cells[1, 1].Value = "父SKU";
                    sheet1.Cells[1, 2].Value = "子SKU数";
                    sheet1.Cells[1, 3].Value = "区域数";
                    sheet1.Cells[1, 4].Value = "楼层数";
                    sheet1.Cells[1, 5].Value = "子SKU";
                    sheet1.Cells[1, 6].Value = "区域";
                    sheet1.Cells[1, 7].Value = "楼层";
                    using (var rng = sheet1.Cells[1, 1, 1, 7])
                    {
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                    }

                    sheet1.Column(1).Width = 16;
                    sheet1.Column(5).Width = 18;

                    for (int i = 0, currentRow = 2, len = _list异常SKU组.Count; i < len; i++)
                    {
                        var grooup = _list异常SKU组[i];
                        var skuCount = grooup.SKUs.Count;
                        //父sku
                        using (var rng = sheet1.Cells[currentRow, 1, currentRow + skuCount - 1, 1])
                        {
                            rng.Merge = true;
                            rng.Value = grooup.ParentSKU;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        //子SKU数
                        using (var rng = sheet1.Cells[currentRow, 2, currentRow + skuCount - 1, 2])
                        {
                            rng.Merge = true;
                            rng.Value = skuCount;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        //区域数
                        using (var rng = sheet1.Cells[currentRow, 3, currentRow + skuCount - 1, 3])
                        {
                            rng.Merge = true;
                            rng.Value = grooup.Areas.Count;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        //楼层数
                        using (var rng = sheet1.Cells[currentRow, 4, currentRow + skuCount - 1, 4])
                        {
                            rng.Merge = true;
                            rng.Value = grooup.Floors.Count;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        //子SKU
                        var childSkus = grooup.SKUs.OrderBy(x => x.Area).ToList();
                        for (int ii = 0, iiLen = childSkus.Count; ii < iiLen; ii++)
                        {
                            var iit = childSkus[ii];
                            sheet1.Cells[currentRow + ii, 5].Value = iit.ChildSKU;
                            sheet1.Cells[currentRow + ii, 6].Value = iit.Area;
                            sheet1.Cells[currentRow + ii, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            sheet1.Cells[currentRow + ii, 7].Value = iit.Floor;
                            sheet1.Cells[currentRow + ii, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            //如果该sku是单独一个楼层,标记一下颜色
                            if (iit.SingleFloor)
                            {
                                using (var iRng = sheet1.Cells[currentRow + ii, 5, currentRow + ii, 7])
                                {
                                    iRng.Style.Font.Color.SetColor(Color.Red);
                                    _单楼层异常SKUs.Add(iit);
                                }

                            }
                        }


                        currentRow += skuCount;

                    }

                    using (var allRng = sheet1.Cells[1, 1, sheet1.Dimension.End.Row, 7])
                    {
                        allRng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        allRng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        allRng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        allRng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    }
                }
                #endregion

                #region 单独楼层SKU
                {
                    var sheet1 = workbox.Worksheets.Add("单独楼层异常SKU");
                    sheet1.Cells[1, 1].Value = "父SKU";
                    sheet1.Cells[1, 2].Value = "子SKU";
                    sheet1.Cells[1, 3].Value = "库位";
                    sheet1.Cells[1, 4].Value = "区域";
                    sheet1.Cells[1, 5].Value = "楼层";
                    sheet1.Column(1).Width = 16;
                    sheet1.Column(2).Width = 16;
                    sheet1.Column(3).Width = 16;

                    _单楼层异常SKUs = _单楼层异常SKUs.OrderBy(x => x.Area).ToList();
                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = _单楼层异常SKUs.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = _单楼层异常SKUs[idx];
                        sheet1.Cells[rowIdx, 1].Value = info.ParentSKU;
                        sheet1.Cells[rowIdx, 2].Value = info.ChildSKU;
                        sheet1.Cells[rowIdx, 3].Value = info.Store;
                        sheet1.Cells[rowIdx, 4].Value = info.Area;
                        sheet1.Cells[rowIdx, 5].Value = info.Floor;
                    }
                    #endregion
                }
                #endregion

                #region 异常sku组分表
                {
                    foreach (var floorName in _listFloors)
                    {
                        var sheet1 = workbox.Worksheets.Add(floorName + "楼");
                        sheet1.Cells[1, 1].Value = "父SKU";
                        sheet1.Cells[1, 2].Value = "该楼层子SKU数(排除单楼层的)";
                        sheet1.Cells[1, 3].Value = "区域数";
                        sheet1.Cells[1, 4].Value = "子SKU";
                        sheet1.Cells[1, 5].Value = "区域";
                        sheet1.Column(1).Width = 16;
                        sheet1.Column(2).Width = 32;
                        sheet1.Column(4).Width = 18;
                        var currentRow = 2;
                        for (int gdx = _list异常SKU组.Count - 1; gdx >= 0; gdx--)
                        {
                            var group = _list异常SKU组[gdx];
                            var errorSkus = new List<_在售SKUMapping>();
                            for (int iix = group.SKUs.Count - 1; iix >= 0; iix--)
                            {
                                var item = group.SKUs[iix];
                                //单楼层就移除了
                                if (item.SingleFloor)
                                {
                                    group.SKUs.RemoveAt(iix);
                                }
                                else
                                {
                                    if (item.Floor == floorName)
                                    {
                                        errorSkus.Add(item);
                                        group.SKUs.RemoveAt(iix);
                                    }
                                }
                            }
                            if (errorSkus.Count == 0)
                                continue;


                            if (errorSkus.Select(x=>x.Area).Distinct().Count()==1)
                                continue;

                            //父SKU
                            using (var rng = sheet1.Cells[currentRow, 1, currentRow + errorSkus.Count - 1, 1])
                            {
                                rng.Merge = true;
                                rng.Value = group.ParentSKU;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                                rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            }
                            //sku详情
                            var orderSkus = errorSkus.OrderBy(x => x.Area).ToList();
                            for (int idx = 0; idx < orderSkus.Count; idx++)
                            {
                                var item = orderSkus[idx];
                                sheet1.Cells[currentRow + idx, 4].Value = item.ChildSKU;
                                sheet1.Cells[currentRow + idx, 5].Value = item.Area;
                            }
                            //sku数量
                            using (var rng = sheet1.Cells[currentRow, 2, currentRow + errorSkus.Count - 1, 2])
                            {
                                rng.Merge = true;
                                rng.Value = errorSkus.Count;
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                                rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            }
                            //区域数量
                            using (var rng = sheet1.Cells[currentRow, 3, currentRow + errorSkus.Count - 1, 3])
                            {
                                rng.Merge = true;
                                rng.Value = errorSkus.Select(x => x.Area).Distinct().Count();
                                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                                rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            }

                            currentRow += errorSkus.Count;
                        }

                        using (var allRng = sheet1.Cells[1, 1, sheet1.Dimension.End.Row, 5])
                        {
                            allRng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            allRng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            allRng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            allRng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        }
                    }
                }
                #endregion

                buffer1 = package.GetAsByteArray();
            }


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
                }
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

        [ExcelTable("在售SKU")]
        class _在售SKUMapping
        {
            [ExcelColumn("子SKU")]
            public string ChildSKU { get; set; }

            [ExcelColumn("父SKU")]
            public string ParentSKU { get; set; }

            [ExcelColumn("库位")]
            public string Store { get; set; }

            public string Area
            {
                get
                {
                    if (!string.IsNullOrWhiteSpace(Store))
                        return Store.Substring(0, 1);
                    return string.Empty;
                }
            }

            public string Floor { get; set; }

            public bool SingleFloor { get; set; }
        }

        class _SKU组
        {
            public string ParentSKU { get; set; }
            public List<string> Areas = new List<string>();
            public List<string> Floors = new List<string>();
            public List<_在售SKUMapping> SKUs = new List<_在售SKUMapping>();
        }

        [ExcelTable("楼层区域表")]
        class _楼层区域表Mapping
        {
            [ExcelColumn("区域")]
            public string Area { get; set; }
            [ExcelColumn("楼层")]
            public string Floor { get; set; }
        }

        class _楼层异常SKU
        {
            public string _Floor { get; set; }
            public List<_SKU组> Groups { get; set; }
        }
    }
}
