using LinqToExcel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CommonLibs;
using LinqToExcel.Attributes;

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
        }

        /**************** button event ****************/

        #region 上传数据
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

        #region 处理按钮事件
        private void BtnCalcu_Click(object sender, EventArgs e)
        {
            var _list在售SKU = new List<_在售SKUMapping>();
            var _list异常在售SKU = new List<_在售SKUMapping>();

            ShowMsg("开始读取数据");

            #region 读取数据
            var actReadData = new Action(() =>
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
            });
            #endregion

            ShowMsg("开始计算数据");

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
                var parentSKUs = _list在售SKU.Select(x => x.ParentSKU).Distinct().ToList();
                foreach (var parentSKU in parentSKUs)
                {
                    var childSKUs = _list在售SKU.Where(x => x.ParentSKU == parentSKU).ToList();
                    if (childSKUs.Select(x => x.Area).Distinct().Count() > 1)
                        _list异常在售SKU.AddRange(childSKUs);
                }
                Export(_list异常在售SKU);
            }, null);
            #endregion
        }
        #endregion


        #region 导出表格说明
        private void LkDecs_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var strDesc = XlsxHelper.GetDecsipt(typeof(_在售SKUMapping));

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
        private void Export(List<_在售SKUMapping> _list异常在售SKU)
        {
            ShowMsg("开始生成表格");
            var buffer1 = new byte[0];

            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 在售
                {
                    var sheet1 = workbox.Worksheets.Add("异常SKU库位信息");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "子SKU";
                    sheet1.Cells[1, 2].Value = "父SKU";
                    sheet1.Cells[1, 3].Value = "库位";
                    sheet1.Cells[1, 4].Value = "区域";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = _list异常在售SKU.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = _list异常在售SKU[idx];
                        sheet1.Cells[rowIdx, 1].Value = info.ChildSKU;
                        sheet1.Cells[rowIdx, 2].Value = info.ParentSKU;
                        sheet1.Cells[rowIdx, 3].Value = info.Store;
                        sheet1.Cells[rowIdx, 4].Value = info.Area;
                    }
                    #endregion

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
        //private void Export(List<_在售SKUMapping> _在售List)
        //{
        //    ShowMsg("开始生成表格");
        //    var buffer1 = new byte[0];

        //    #region 生成表
        //    using (ExcelPackage package = new ExcelPackage())
        //    {
        //        var workbox = package.Workbook;

        //        #region 在售
        //        {
        //            var sheet1 = workbox.Worksheets.Add("在售");

        //            #region 标题行
        //            sheet1.Cells[1, 1].Value = "SKU";
        //            sheet1.Cells[1, 2].Value = "商品名称";
        //            sheet1.Cells[1, 3].Value = "上海库位";
        //            sheet1.Cells[1, 4].Value = "上海仓库";
        //            sheet1.Cells[1, 5].Value = "上海货架";
        //            sheet1.Cells[1, 6].Value = "昆山库位";
        //            sheet1.Cells[1, 7].Value = "昆山区域";
        //            sheet1.Cells[1, 8].Value = "可用数量";
        //            sheet1.Cells[1, 9].Value = "备注";
        //            #endregion

        //            #region 数据行
        //            for (int idx = 0, rowIdx = 2, len = _在售List.Count; idx < len; idx++, rowIdx++)
        //            {
        //                var info = _在售List[idx];
        //                sheet1.Cells[rowIdx, 1].Value = info._SKU;
        //                sheet1.Cells[rowIdx, 2].Value = info._商品名称;
        //                sheet1.Cells[rowIdx, 3].Value = info._上海库位;
        //                sheet1.Cells[rowIdx, 4].Value = info._上海仓库;
        //                sheet1.Cells[rowIdx, 5].Value = info._上海货架;
        //                sheet1.Cells[rowIdx, 6].Value = info._昆山库位;
        //                sheet1.Cells[rowIdx, 7].Value = info._昆山区域;
        //                sheet1.Cells[rowIdx, 8].Value = info._可用数量;
        //                sheet1.Cells[rowIdx, 9].Value = info._备注;
        //            }
        //            #endregion

        //        }
        //        #endregion

        //        buffer1 = package.GetAsByteArray();
        //    }
        //    #endregion

        //    InvokeMainForm((obj) =>
        //    {
        //        SaveFileDialog saveFile = new SaveFileDialog();
        //        saveFile.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
        //        saveFile.Title = "导出数据";//设置标题
        //        saveFile.AddExtension = true;//是否自动增加所辍名
        //        saveFile.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
        //        if (saveFile.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
        //        {
        //            var FileName = saveFile.FileName;//得到文件路径   
        //            var saveFilName = Path.GetFileNameWithoutExtension(FileName);
        //            var len = buffer1.Length;
        //            using (var fs = File.Create(FileName, len))
        //            {
        //                fs.Write(buffer1, 0, len);
        //            }
        //            ShowMsg("表格生成完毕");
        //        }
        //    }, null);
        //}
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
        }

    }
}
