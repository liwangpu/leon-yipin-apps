using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Gadget
{
    public partial class _库位号订单提取 : Form
    {
        public _库位号订单提取()
        {
            InitializeComponent();
        }

        private void _库位号订单提取_Load(object sender, EventArgs e)
        {
            //txt订单表.Text = @"C:\Users\Leon\Desktop\11111.xlsx";
            btnCalcu.Enabled = false;
        }

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

        /**************** common class ****************/

        private void btnCalcu_Click(object sender, EventArgs e)
        {
            #region 读取数据
            var actReadData = new Action(() =>
            {
                ShowMsg("开始读取订单信息");
                using (var package = new ExcelPackage(new FileInfo(txt订单表.Text)))
                {
                    var worksheet = package.Workbook.Worksheets[1];
                    var endRow = worksheet.Dimension.End.Row;
                    var endColumn = worksheet.Dimension.End.Column;
                    var _need库位 = txt库位号.Text.ToUpper();
                    for (int idx = endRow; idx >= 2; idx--)
                    {
                        var _库位obj = worksheet.Cells[idx, 8].Value;
                        if (_库位obj != null)
                        {
                            var _库位str = _库位obj.ToString();

                            var notneed = _库位str.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Substring(0, 1).ToUpper()).Any(x => x != _need库位);
                            if (notneed)
                                worksheet.DeleteRow(idx, 1);

                        }
                        else
                            worksheet.DeleteRow(idx, 1);

                        //Debugger.Log(1, "", idx.ToString() + Environment.NewLine);

                    }
                    //package.SaveAs();
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
                            package.SaveAs(new FileInfo(FileName));
                            ShowMsg("表格生成完毕");
                        }
                    }, null);
                }
            });
            #endregion

            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {
            }, null);
            #endregion
        }

        private void btn上传订单表_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.Filter = "Execl 97-2003工作簿|*.xls|Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Filter = "Excel 工作簿|*.xlsx";//设置文件类型
            OpenFileDialog1.Title = "表格信息";//设置标题
            OpenFileDialog1.Multiselect = false;
            OpenFileDialog1.AutoUpgradeEnabled = true;//是否随系统升级而升级外观
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)//如果点的是确定就得到文件路径
            {
                txt订单表.Text = OpenFileDialog1.FileName;
                btnCalcu.Enabled = true;
            }
        }

        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }
    }
}
