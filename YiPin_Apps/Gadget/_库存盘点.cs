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
    public partial class _库存盘点 : Form
    {
        public _库存盘点()
        {
            InitializeComponent();
        }

        private void _库存盘点_Load(object sender, EventArgs e)
        {
            txtUpJiaoHuo.Text = @"C:\Users\Leon\Desktop\aaa\拣货表.xlsx";
            txtUpKucun.Text = @"C:\Users\Leon\Desktop\aaa\上海所有库存.xlsx";
        }

        /**************** button event ****************/

        #region 上传拣货表
        private void btnUpJiaoHuo_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpJiaoHuo.Text = OpenFileDialog1.FileName;
            }
        }
        #endregion

        #region 上传库存表
        private void btnUpKucun_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txtUpKucun.Text = OpenFileDialog1.FileName;
            }
        }
        #endregion

        #region 数据处理
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            var list拣货信息 = new List<_拣货表>();
            var list库存信息 = new List<_库存表>();
            var list结果信息 = new List<_导出表>();

            #region 读取数据
            var actReadData = new Action(() =>
            {
                ShowMsg("开始读取表格数据");

                var str拣货表Path = txtUpJiaoHuo.Text;
                var str库存表Path = txtUpKucun.Text;

                #region 读取拣货表
                if (!string.IsNullOrEmpty(str拣货表Path))
                {
                    using (var excel = new ExcelQueryFactory(str拣货表Path))
                    {
                        //excel.StrictMapping = LinqToExcel.Query.StrictMappingType.ClassStrict;
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<_拣货表>(s)
                                          where c._SKU != null
                                          select c;
                                list拣货信息.AddRange(tmp);
                            }
                            catch (Exception ex)
                            {
                                ShowMsg(ex.Message);
                            }
                        });
                    }
                }
                #endregion

                #region 读取库存表
                if (!string.IsNullOrEmpty(str库存表Path))
                {
                    using (var excel = new ExcelQueryFactory(str库存表Path))
                    {
                        //excel.StrictMapping = LinqToExcel.Query.StrictMappingType.ClassStrict;
                        var sheetNames = excel.GetWorksheetNames().ToList();
                        sheetNames.ForEach(s =>
                        {
                            try
                            {
                                var tmp = from c in excel.Worksheet<_库存表>(s)
                                          where c._SKU != null
                                          select c;
                                list库存信息.AddRange(tmp);
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

            actReadData.BeginInvoke((ob) =>
            {
                ShowMsg("开始计算数据");
                var list拣货SKU = list拣货信息.Select(x => x._SKU).Distinct().ToList();

                list拣货SKU.ForEach(curSKU =>
                {
                    /*
                     * 两类sku,一类为CLBA10A94-B,CLBA10A94,其中第一个是第二个的子sku,我们称第一个是父sku
                     * 父sku一般不含中间横线,且一般以数值结尾
                     * 另一类为CLBA12A2F,CLBA12A2A,CLBA12A2B,不含中间横线,一般这个最后的字母是代表子类
                     * 所以截取这个字母前的字符作为父sku
                     */
                    var stParentPart = string.Empty;

                    #region 匹配第一类sku
                    if (curSKU.Contains('-'))
                    {
                        var indx = curSKU.IndexOf('-');
                        if (indx != -1)
                            stParentPart = curSKU.Substring(0, indx);
                    }
                    #endregion

                    #region 匹配第二类sku
                    if (string.IsNullOrEmpty(stParentPart))
                    {
                        var lastChar = curSKU.Substring(curSKU.Length - 1, 1);
                        try
                        {
                            var a = Convert.ToInt32(lastChar);
                        }
                        catch (Exception)
                        {
                            stParentPart = curSKU.Substring(0, curSKU.Length - 1);
                        }
                    }
                    #endregion

                    stParentPart = !string.IsNullOrEmpty(stParentPart) ? stParentPart : curSKU;

                    var refStoreItems = list库存信息.Where(sto => sto._SKU.Contains(stParentPart)).ToList();
                    if (refStoreItems != null && refStoreItems.Count > 0)
                    {
                        #region 第一个默认是拣货单sku
                        {
                            var defaultItem = refStoreItems.Where(x => x._SKU == stParentPart).FirstOrDefault();
                            if (defaultItem != null)
                            {
                                var bExist = list结果信息.Count(x => x._SKU == defaultItem._SKU) > 0;
                                if (!bExist)
                                {
                                    var data = new _导出表();
                                    data._SKU = defaultItem._SKU;
                                    data._可用数量 = defaultItem._可用数量;
                                    data._库存数量 = defaultItem._库存数量;
                                    data._占用数量 = defaultItem._占用数量;
                                    data._库位 = defaultItem._库位;
                                    if (curSKU == stParentPart)
                                    {
                                        var ref拣货Item = list拣货信息.Where(x => x._SKU == curSKU).FirstOrDefault();
                                        if (ref拣货Item != null)
                                        {
                                            data._拣货单数量 = ref拣货Item._拣货单数量;
                                            data._实际仓库数量 = ref拣货Item._实际仓库数量;
                                        }
                                    }
                                    list结果信息.Add(data);
                                }
                            }
                        }
                        #endregion

                        #region 余下的子sku
                        {
                            var remainItems = refStoreItems.Where(x => x._SKU != stParentPart).ToList();
                            if (remainItems != null && remainItems.Count > 0)
                            {
                                foreach (var item in remainItems)
                                {
                                    var bExist = list结果信息.Count(x => x._SKU == item._SKU) > 0;
                                    if (!bExist)
                                    {
                                        var data = new _导出表();
                                        data._SKU = item._SKU;
                                        data._可用数量 = item._可用数量;
                                        data._库存数量 = item._库存数量;
                                        data._占用数量 = item._占用数量;
                                        data._库位 = item._库位;
                                        if (curSKU == stParentPart)
                                        {
                                            var ref拣货Item = list拣货信息.Where(x => x._SKU == curSKU).FirstOrDefault();
                                            if (ref拣货Item != null)
                                            {
                                                data._拣货单数量 = ref拣货Item._拣货单数量;
                                                data._实际仓库数量 = ref拣货Item._实际仓库数量;
                                            }
                                        }
                                        list结果信息.Add(data);
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                });

                Export(list结果信息);
            }, null);
        }
        #endregion

        /**************** common method ****************/

        #region Export 导出结果表格
        private void Export(List<_导出表> list)
        {
            ShowMsg("开始生成表格");
            var buffer = new byte[0];

            #region 数据表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                var sheet1 = workbox.Worksheets.Add("Sheet1");

                #region 标题行
                sheet1.Cells[1, 1].Value = "SKU";
                sheet1.Cells[1, 2].Value = "库存数量";
                sheet1.Cells[1, 3].Value = "占用数量";
                sheet1.Cells[1, 4].Value = "可用数量";
                sheet1.Cells[1, 5].Value = "拣货单数量";
                sheet1.Cells[1, 6].Value = "实际仓库数量";
                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = list.Count; idx < len; idx++, rowIdx++)
                {
                    var info = list[idx];
                    sheet1.Cells[rowIdx, 1].Value = info._SKU;
                    sheet1.Cells[rowIdx, 2].Value = info._库存数量;
                    sheet1.Cells[rowIdx, 3].Value = info._占用数量;
                    sheet1.Cells[rowIdx, 4].Value = info._可用数量;
                    sheet1.Cells[rowIdx, 5].Value = info._拣货单数量;
                    sheet1.Cells[rowIdx, 6].Value = info._实际仓库数量;
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
                    var len = buffer.Length;
                    using (var fs = File.Create(FileName, len))
                    {
                        fs.Write(buffer, 0, len);
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

        class _拣货表
        {
            [ExcelColumn("SKU")]
            public string _SKU { get; set; }
            [ExcelColumn("拣货单数量")]
            public decimal _拣货单数量 { get; set; }
            [ExcelColumn("实际仓库数量")]
            public decimal _实际仓库数量 { get; set; }
        }

        class _库存表
        {
            [ExcelColumn("SKU码")]
            public string _SKU { get; set; }
            [ExcelColumn("库存数量")]
            public decimal _库存数量 { get; set; }
            [ExcelColumn("占用数量")]
            public decimal _占用数量 { get; set; }
            [ExcelColumn("可用数量")]
            public decimal _可用数量 { get; set; }
            [ExcelColumn("库位")]
            public string _库位 { get; set; }
        }

        class _导出表
        {
            public string _SKU { get; set; }
            public decimal _库存数量 { get; set; }
            public decimal _占用数量 { get; set; }
            public decimal _可用数量 { get; set; }
            public decimal _拣货单数量 { get; set; }
            public decimal _实际仓库数量 { get; set; }
            public string _库位 { get; set; }
        }

    }
}
