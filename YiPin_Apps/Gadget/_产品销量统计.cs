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
    public partial class _产品销量统计 : Form
    {
        public _产品销量统计()
        {
            InitializeComponent();
        }

        private void _产品销量统计_Load(object sender, EventArgs e)
        {
            //dtp开发起始时间.Value = Convert.ToDateTime(string.Format("{0}-01-01", DateTime.Now.Year));
            //dtp开发截止时间.Value = Convert.ToDateTime(string.Format("{0}-{1}-01", DateTime.Now.Year, DateTime.Now.Month - 3 > 0 ? DateTime.Now.Month - 3 : 1));



            dtp开发起始时间.Value = Convert.ToDateTime(string.Format("2017-01-01"));
            dtp开发截止时间.Value = Convert.ToDateTime(string.Format("2018-02-01"));
            txt在售商品信息.Text = @"C:\Users\pulw\Desktop\产品统计\在售.csv";
            txt停售商品信息.Text = @"C:\Users\pulw\Desktop\产品统计\停售.csv";
            txt各平台销量一览表.Text = @"C:\Users\pulw\Desktop\产品统计\10到12月销量.csv";




        }

        /**************** button event ****************/

        #region 上传在售商品信息按钮事件
        private void btn在售商品信息_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt在售商品信息);
        }
        #endregion

        #region 上传停售商品信息按钮事件
        private void btn停售商品信息_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt停售商品信息);
        }
        #endregion

        #region 上传各平台销量一览表按钮事件
        private void btn各平台销量一览表_Click(object sender, EventArgs e)
        {
            FormHelper.GetCSVPath(txt各平台销量一览表);
        }
        #endregion

        #region 处理数据按钮事件
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            btnAnalyze.Enabled = false;
            var dt开发起始时间 = dtp开发起始时间.Value;
            var dt开发截止时间 = dtp开发截止时间.Value.AddDays(1);
            var d滞销量 = nud滞销量.Value;
            var d爆款量 = nud爆款量.Value;

            var list在售商品信息 = new List<_产品信息>();
            var list停售商品信息 = new List<_产品信息>();
            var list各平台销量信息 = new List<_各平台销量>();


            var list统计范围内的在售商品SKU = new List<string>();
            var list统计范围内的停售商品SKU = new List<string>();

            var volumeStore = new _SKU统计仓库();

            ShowMsg("开始读取数据");

            #region 读取数据
            var actReadData = new Action(() =>
            {
                var strMsg = string.Empty;
                ShowMsg("开始读取在售产品数据");
                FormHelper.ReadCSVFile<_产品信息>(txt在售商品信息.Text, ref list在售商品信息, ref strMsg);
                ShowMsg(strMsg);

                ShowMsg("开始读取停售产品数据");
                FormHelper.ReadCSVFile<_产品信息>(txt停售商品信息.Text, ref list停售商品信息, ref strMsg);
                ShowMsg(strMsg);

                ShowMsg("开始读取各平台销量一览数据");
                FormHelper.ReadCSVFile<_各平台销量>(txt各平台销量一览表.Text, ref list各平台销量信息, ref strMsg);
                ShowMsg(strMsg);
            });
            #endregion

            ShowMsg("读取完毕,开始处理数据");
            #region 处理数据
            actReadData.BeginInvoke((obj) =>
            {

                #region 将统计范围内的在售/停售产品加入 =>list统计范围内的在售商品SKU/list统计范围内的停售商品SKU
                {
                    //list统计范围内的在售商品SKU = list在售商品信息.Where(x => x._开发时间 >= dt开发起始时间 && x._开发时间 < dt开发截止时间).Select(x => x.SKU).Distinct().ToList();
                    //list统计范围内的停售商品SKU = list停售商品信息.Where(x => x._开发时间 >= dt开发起始时间 && x._开发时间 < dt开发截止时间).Select(x => x.SKU).Distinct().ToList();


                    /*
                     *之所以没有用上面的方式是需要删去不用的产品信息,以便在从商品信息统计金额数据量没有那么大
                     */
                    for (int idx = list在售商品信息.Count - 1; idx >= 0; idx--)
                    {
                        var curItem = list在售商品信息[idx];


                        //if (curItem.SKU.Trim() == "DNFA5A25-L")
                        //{

                        //}

                        if (curItem._开发时间 >= dt开发起始时间 && curItem._开发时间 < dt开发截止时间)
                        {
                            list统计范围内的在售商品SKU.Add(curItem.SKU);
                        }
                        else
                        {
                            list在售商品信息.RemoveAt(idx);
                        }
                    }

                    for (int idx = list停售商品信息.Count - 1; idx >= 0; idx--)
                    {
                        var curItem = list停售商品信息[idx];
                        //if (curItem.SKU.Trim() == "DNFA5A25-L")
                        //{

                        //}

                        if (curItem._开发时间 >= dt开发起始时间 && curItem._开发时间 < dt开发截止时间)
                        {
                            list统计范围内的停售商品SKU.Add(curItem.SKU);
                        }
                        else
                        {
                            list停售商品信息.RemoveAt(idx);
                        }
                    }
                }
                #endregion

                #region 遍历各平台销量,标记记录是否是爆款/滞销/停售 同时也删去不在统计范围的销量纪录
                {

                    for (int idx = list各平台销量信息.Count - 1; idx >= 0; idx--)
                    {
                        var bRemove = false;
                        var curItem = list各平台销量信息[idx];


                        //if (curItem.SKU == "DNFA5A25-L")
                        //{

                        //}


                        if (!list统计范围内的停售商品SKU.Exists(new Predicate<string>((t) => { return t == curItem.SKU; })))
                        {
                            if (list统计范围内的在售商品SKU.Exists(new Predicate<string>((t) => { return t == curItem.SKU; })))
                            {
                                var saleAmount = list各平台销量信息.Where(x => x.SKU == curItem.SKU).Sum(x => x._总销量);
                                if (saleAmount <= d滞销量)
                                {
                                    curItem._产品类型 = _Enum产品类型._滞销;
                                }
                                else
                                {
                                    if (saleAmount >= d爆款量)
                                    {
                                        curItem._产品类型 = _Enum产品类型._爆款;
                                    }
                                    else
                                    {
                                        curItem._产品类型 = _Enum产品类型._普通款;
                                    }
                                }
                            }
                            else
                            {
                                bRemove = true;
                                list各平台销量信息.RemoveAt(idx);
                            }
                        }
                        else
                        {
                            curItem._产品类型 = _Enum产品类型._停售;
                        }

                        //if (curItem.SKU=="LGDH9H17")
                        //{

                        //}
                        if (!bRemove)
                        {

                            //if (curItem.SKU == "LGDP6P39-B-6SP")
                            //{

                            //}

                            //if (curItem.SKU == "LGDP6P09-R")
                            //{

                            //}

                            volumeStore.AddSaleVolume(curItem._开发, curItem.SKU, curItem._总销量);
                        }

                    }
                }
                #endregion


                var list开发销量统计 = new List<_滞销爆款统计>();
                var list供应商统计 = new List<_供应商统计>();
                var list滞销爆款停售详情 = new List<_滞销爆款停售详情>();
                var list爆款滞销统计的详情 = new List<_爆款滞销统计详情>();

                var list开发人员 = list在售商品信息.Select(x => x._开发).Distinct().ToList();
                #region 获取开发人员姓名信息
                if (list停售商品信息.Count > 0)
                {
                    var tmp = list停售商品信息.Select(x => x._开发).Distinct().ToList();
                    foreach (var item in tmp)
                    {
                        if (list开发人员.Count(x => x == item) == 0)
                            list开发人员.Add(item);
                    }
                }
                #endregion

                #region 统计开发人员的爆款/滞销信息
                list开发人员.ForEach(str开发姓名 =>
                {
                    if (str开发姓名 == "张浩")
                    {

                    }

                    var listBao = new List<_爆款滞销统计详情>();
                    var listZhi = new List<_爆款滞销统计详情>();
                    var model = new _滞销爆款统计();
                    model._开发人员 = str开发姓名;
                    model._爆款主SKU个数 = volumeStore.CalcuGoodVolume(str开发姓名, d爆款量, ref listBao);
                    model._滞销主SKU个数 = volumeStore.CalcuBadVolume(str开发姓名, d滞销量, ref listZhi);
                    list爆款滞销统计的详情.AddRange(listBao);
                    list爆款滞销统计的详情.AddRange(listZhi);

                    #region 统计爆款
                    {
                        var refList销售记录 = list各平台销量信息.Where(x => x._开发 == str开发姓名 && x._产品类型 == _Enum产品类型._爆款).ToList();
                        model._爆款SKU个数 = refList销售记录.Select(x => x.SKU).Distinct().Count();

                        var amount = from rr in refList销售记录
                                     join zz in list在售商品信息 on rr.SKU equals zz.SKU
                                     select zz._库存金额;
                        model._爆款总金额 = amount.Sum();

                        list滞销爆款停售详情.AddRange(refList销售记录.Select(x => new _滞销爆款停售详情()
                        {
                            Refer产品信息 = list在售商品信息.Where(zz => zz.SKU == x.SKU).First(),
                            _SKU = x.SKU,
                            _Amazon销量 = x._Amazon销量,
                            _Ebay销量 = x._Ebay销量,
                            _Joom销量 = x._Joom销量,
                            _Shopee销量 = x._Shopee销量,
                            _SMT销量 = x._SMT销量,
                            _wish销量 = x._wish销量,
                            _产品类型 = _Enum产品类型._爆款
                        }).ToList());

                    }
                    #endregion

                    #region 统计滞销
                    {
                        var refList销售记录 = list各平台销量信息.Where(x => x._开发 == str开发姓名 && x._产品类型 == _Enum产品类型._滞销).ToList();
                        model._滞销SKU个数 = refList销售记录.Select(x => x.SKU).Distinct().Count();

                        var amount = from rr in refList销售记录
                                     join zz in list在售商品信息 on rr.SKU equals zz.SKU
                                     select zz._库存金额;
                        model._滞销总金额 = amount.Sum();

                        list滞销爆款停售详情.AddRange(refList销售记录.Select(x => new _滞销爆款停售详情()
                        {
                            Refer产品信息 = list在售商品信息.Where(zz => zz.SKU == x.SKU).FirstOrDefault(),
                            _SKU = x.SKU,
                            _Amazon销量 = x._Amazon销量,
                            _Ebay销量 = x._Ebay销量,
                            _Joom销量 = x._Joom销量,
                            _Shopee销量 = x._Shopee销量,
                            _SMT销量 = x._SMT销量,
                            _wish销量 = x._wish销量,
                            _产品类型 = _Enum产品类型._滞销
                        }).ToList());
                    }
                    #endregion

                    #region 统计停售
                    {
                        var refList销售记录 = list各平台销量信息.Where(x => x._开发 == str开发姓名 && x._产品类型 == _Enum产品类型._停售).ToList();
                        model._停售SKU个数 = refList销售记录.Select(x => x.SKU).Distinct().Count();

                        var amount = from rr in refList销售记录
                                     join zz in list停售商品信息 on rr.SKU equals zz.SKU
                                     select zz._库存金额;
                        model._停售总金额 = amount.Sum();

                        list滞销爆款停售详情.AddRange(refList销售记录.Select(x => new _滞销爆款停售详情()
                        {
                            Refer产品信息 = list停售商品信息.Where(zz => zz.SKU == x.SKU).First(),
                            _SKU = x.SKU,
                            _Amazon销量 = x._Amazon销量,
                            _Ebay销量 = x._Ebay销量,
                            _Joom销量 = x._Joom销量,
                            _Shopee销量 = x._Shopee销量,
                            _SMT销量 = x._SMT销量,
                            _wish销量 = x._wish销量,
                            _产品类型 = _Enum产品类型._停售
                        }).ToList());
                    }
                    #endregion

                    #region 统计普通款
                    {
                        var refList销售记录 = list各平台销量信息.Where(x => x._开发 == str开发姓名 && x._产品类型 == _Enum产品类型._普通款).ToList();
                        model._普通款SKU个数 = refList销售记录.Select(x => x.SKU).Distinct().Count();

                        var amount = from rr in refList销售记录
                                     join zz in list在售商品信息 on rr.SKU equals zz.SKU
                                     select zz._库存金额;
                        model._普通款总金额 = amount.Sum();
                    }
                    #endregion

                    if (model._所有SKU个数 > 0)
                    {
                        list开发销量统计.Add(model);
                    }

                });
                #endregion

                #region 计算供应商统计
                {
                    var list供应商名称 = list滞销爆款停售详情.Select(x => x._供应商).Distinct().ToList();
                    list供应商名称.ForEach(str供应商名称 =>
                    {
                        var model = new _供应商统计();
                        model._供应商 = str供应商名称;
                        var refDetails = list滞销爆款停售详情.Where(x => x._供应商 == str供应商名称).ToList();

                        {
                            var refSpecDetail = refDetails.Where(x => x._产品类型 == _Enum产品类型._爆款).ToList();
                            model._爆款SKU个数 = refSpecDetail.Select(x => x._SKU).Distinct().Count();
                            model._爆款库存金额 = refSpecDetail.Select(x => x._库存金额).Sum();
                        }

                        {
                            var refSpecDetail = refDetails.Where(x => x._产品类型 == _Enum产品类型._滞销).ToList();
                            model._滞销SKU个数 = refSpecDetail.Select(x => x._SKU).Distinct().Count();
                            model._滞销库存金额 = refSpecDetail.Select(x => x._库存金额).Sum();
                        }

                        {
                            var refSpecDetail = refDetails.Where(x => x._产品类型 == _Enum产品类型._普通款).ToList();
                            model._普通款SKU个数 = refSpecDetail.Select(x => x._SKU).Distinct().Count();
                            model._普通款库存金额 = refSpecDetail.Select(x => x._库存金额).Sum();
                        }

                        {
                            var refSpecDetail = refDetails.Where(x => x._产品类型 == _Enum产品类型._停售).ToList();
                            model._停售SKU个数 = refSpecDetail.Select(x => x._SKU).Distinct().Count();
                            model._停售库存金额 = refSpecDetail.Select(x => x._库存金额).Sum();
                        }

                        list供应商统计.Add(model);
                    });
                }
                #endregion

                Export(list开发销量统计.OrderByDescending(x => x._滞销总金额).ToList(), list滞销爆款停售详情.OrderByDescending(x => x._库存金额).ToList()
                    , list供应商统计.OrderByDescending(x => x._滞销库存金额).ToList(), list爆款滞销统计的详情.OrderByDescending(x => x._开发).ToList());

            }, null);
            #endregion

        }
        #endregion

        #region 导出表格说明按钮事件
        private void lkDecs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormHelper.GenerateTableDes(typeof(_产品信息), typeof(_各平台销量));
        }
        #endregion

        /**************** common method ****************/



        #region Export 导出表格
        private void Export(List<_滞销爆款统计> list滞销爆款统计, List<_滞销爆款停售详情> list滞销爆款停售详情
            , List<_供应商统计> list供应商统计, List<_爆款滞销统计详情> list_爆款滞销统计详情)
        {
            ShowMsg("计算完毕,开始生成表格");
            var buffer1 = new byte[0];

            #region 生成表
            using (ExcelPackage package = new ExcelPackage())
            {
                var workbox = package.Workbook;

                #region 开发销量统计
                {
                    var sheet1 = workbox.Worksheets.Add("开发销量统计");
                    #region 标题行
                    sheet1.Cells[1, 1].Value = "开发";
                    sheet1.Cells[1, 2].Value = "滞销SKU个数";
                    sheet1.Cells[1, 3].Value = "滞销主SKU个数";
                    sheet1.Cells[1, 4].Value = "滞销总金额";
                    sheet1.Cells[1, 5].Value = "爆款SKU个数";
                    sheet1.Cells[1, 6].Value = "爆款主SKU个数";
                    sheet1.Cells[1, 7].Value = "爆款总金额";
                    sheet1.Cells[1, 8].Value = "停售SKU个数";
                    sheet1.Cells[1, 9].Value = "停售总金额";
                    sheet1.Cells[1, 10].Value = "普通款SKU个数";
                    sheet1.Cells[1, 11].Value = "普通款总金额";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list滞销爆款统计.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list滞销爆款统计[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._开发人员;
                        sheet1.Cells[rowIdx, 2].Value = info._滞销SKU个数;
                        sheet1.Cells[rowIdx, 3].Value = info._滞销主SKU个数;
                        sheet1.Cells[rowIdx, 4].Value = info._滞销总金额;
                        sheet1.Cells[rowIdx, 5].Value = info._爆款SKU个数;
                        sheet1.Cells[rowIdx, 6].Value = info._爆款主SKU个数;
                        sheet1.Cells[rowIdx, 7].Value = info._爆款总金额;
                        sheet1.Cells[rowIdx, 8].Value = info._停售SKU个数;
                        sheet1.Cells[rowIdx, 9].Value = info._停售总金额;
                        sheet1.Cells[rowIdx, 10].Value = info._普通款SKU个数;
                        sheet1.Cells[rowIdx, 11].Value = info._普通款总金额;
                    }
                    #endregion

                }
                #endregion

                #region 供应商统计
                {
                    var sheet1 = workbox.Worksheets.Add("供应商统计");
                    #region 标题行
                    sheet1.Cells[1, 1].Value = "供应商";
                    sheet1.Cells[1, 2].Value = "滞销SKU个数";
                    sheet1.Cells[1, 3].Value = "滞销总金额";
                    sheet1.Cells[1, 4].Value = "爆款SKU个数";
                    sheet1.Cells[1, 5].Value = "爆款总金额";
                    sheet1.Cells[1, 6].Value = "停售SKU个数";
                    sheet1.Cells[1, 7].Value = "停售总金额";
                    sheet1.Cells[1, 8].Value = "普通款SKU个数";
                    sheet1.Cells[1, 9].Value = "普通款总金额";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list供应商统计.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list供应商统计[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._供应商;
                        sheet1.Cells[rowIdx, 2].Value = info._滞销SKU个数;
                        sheet1.Cells[rowIdx, 3].Value = info._滞销库存金额;
                        sheet1.Cells[rowIdx, 4].Value = info._爆款SKU个数;
                        sheet1.Cells[rowIdx, 5].Value = info._爆款库存金额;
                        sheet1.Cells[rowIdx, 6].Value = info._停售SKU个数;
                        sheet1.Cells[rowIdx, 7].Value = info._停售库存金额;
                        sheet1.Cells[rowIdx, 8].Value = info._普通款SKU个数;
                        sheet1.Cells[rowIdx, 9].Value = info._普通款库存金额;
                    }
                    #endregion
                }
                #endregion

                #region 爆款销量详情
                {
                    var sheet1 = workbox.Worksheets.Add("爆款销量详情");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "供应商";
                    sheet1.Cells[1, 3].Value = "可用数量";
                    sheet1.Cells[1, 4].Value = "总销量";
                    sheet1.Cells[1, 5].Value = "wish";
                    sheet1.Cells[1, 6].Value = "Ebay";
                    sheet1.Cells[1, 7].Value = "SMT";
                    sheet1.Cells[1, 8].Value = "Amazon";
                    sheet1.Cells[1, 9].Value = "Shopee";
                    sheet1.Cells[1, 10].Value = "Joom";
                    #endregion

                    #region 数据行
                    var tmp = list滞销爆款停售详情.Where(x => x._产品类型 == _Enum产品类型._爆款).OrderByDescending(x => x._总销量).ToList();
                    for (int idx = 0, rowIdx = 2, len = tmp.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = tmp[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._SKU;
                        sheet1.Cells[rowIdx, 2].Value = info._供应商;
                        sheet1.Cells[rowIdx, 3].Value = info._可用数量;
                        sheet1.Cells[rowIdx, 4].Value = info._总销量;
                        sheet1.Cells[rowIdx, 5].Value = info._wish销量;
                        sheet1.Cells[rowIdx, 6].Value = info._Ebay销量;
                        sheet1.Cells[rowIdx, 7].Value = info._SMT销量;
                        sheet1.Cells[rowIdx, 8].Value = info._Amazon销量;
                        sheet1.Cells[rowIdx, 9].Value = info._Shopee销量;
                        sheet1.Cells[rowIdx, 10].Value = info._Joom销量;
                    }
                    #endregion
                }
                #endregion

                #region 滞销销量详情
                {
                    var sheet1 = workbox.Worksheets.Add("滞销销量详情");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "供应商";
                    sheet1.Cells[1, 3].Value = "可用数量";
                    sheet1.Cells[1, 4].Value = "总销量";
                    sheet1.Cells[1, 5].Value = "wish";
                    sheet1.Cells[1, 6].Value = "Ebay";
                    sheet1.Cells[1, 7].Value = "SMT";
                    sheet1.Cells[1, 8].Value = "Amazon";
                    sheet1.Cells[1, 9].Value = "Shopee";
                    sheet1.Cells[1, 10].Value = "Joom";
                    #endregion

                    #region 数据行
                    var tmp = list滞销爆款停售详情.Where(x => x._产品类型 == _Enum产品类型._滞销).OrderByDescending(x => x._总销量).ToList();
                    for (int idx = 0, rowIdx = 2, len = tmp.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = tmp[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._SKU;
                        sheet1.Cells[rowIdx, 2].Value = info._供应商;
                        sheet1.Cells[rowIdx, 3].Value = info._可用数量;
                        sheet1.Cells[rowIdx, 4].Value = info._总销量;
                        sheet1.Cells[rowIdx, 5].Value = info._wish销量;
                        sheet1.Cells[rowIdx, 6].Value = info._Ebay销量;
                        sheet1.Cells[rowIdx, 7].Value = info._SMT销量;
                        sheet1.Cells[rowIdx, 8].Value = info._Amazon销量;
                        sheet1.Cells[rowIdx, 9].Value = info._Shopee销量;
                        sheet1.Cells[rowIdx, 10].Value = info._Joom销量;
                    }
                    #endregion
                }
                #endregion

                #region 停售销量详情
                {
                    var sheet1 = workbox.Worksheets.Add("停售销量详情");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "SKU";
                    sheet1.Cells[1, 2].Value = "供应商";
                    sheet1.Cells[1, 3].Value = "可用数量";
                    sheet1.Cells[1, 4].Value = "总销量";
                    sheet1.Cells[1, 5].Value = "wish";
                    sheet1.Cells[1, 6].Value = "Ebay";
                    sheet1.Cells[1, 7].Value = "SMT";
                    sheet1.Cells[1, 8].Value = "Amazon";
                    sheet1.Cells[1, 9].Value = "Shopee";
                    sheet1.Cells[1, 10].Value = "Joom";
                    #endregion

                    #region 数据行
                    var tmp = list滞销爆款停售详情.Where(x => x._产品类型 == _Enum产品类型._停售).OrderByDescending(x => x._总销量).ToList();
                    for (int idx = 0, rowIdx = 2, len = tmp.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = tmp[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._SKU;
                        sheet1.Cells[rowIdx, 2].Value = info._供应商;
                        sheet1.Cells[rowIdx, 3].Value = info._可用数量;
                        sheet1.Cells[rowIdx, 4].Value = info._总销量;
                        sheet1.Cells[rowIdx, 5].Value = info._wish销量;
                        sheet1.Cells[rowIdx, 6].Value = info._Ebay销量;
                        sheet1.Cells[rowIdx, 7].Value = info._SMT销量;
                        sheet1.Cells[rowIdx, 8].Value = info._Amazon销量;
                        sheet1.Cells[rowIdx, 9].Value = info._Shopee销量;
                        sheet1.Cells[rowIdx, 10].Value = info._Joom销量;
                    }
                    #endregion
                }
                #endregion

                #region 爆款滞销详情
                {
                    var sheet1 = workbox.Worksheets.Add("爆款滞销详情");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "开发";
                    sheet1.Cells[1, 2].Value = "主SKU";
                    sheet1.Cells[1, 3].Value = "子SKU个数";
                    sheet1.Cells[1, 4].Value = "子SKU所有销量";
                    sheet1.Cells[1, 5].Value = "所有子SKU销量信息";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = list_爆款滞销统计详情.Count; idx < len; idx++, rowIdx++)
                    {
                        var info = list_爆款滞销统计详情[idx];
                        sheet1.Cells[rowIdx, 1].Value = info._开发;
                        sheet1.Cells[rowIdx, 2].Value = info._主SKU;
                        sheet1.Cells[rowIdx, 3].Value = info._子SKU个数;
                        sheet1.Cells[rowIdx, 4].Value = info._子SKU所有销量;
                        sheet1.Cells[rowIdx, 5].Value = info._所有子SKU销量信息;
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

        [ExcelTable("在售/停售商品信息表")]
        class _产品信息
        {
            private string orgSKU;
            private string org商品名称;

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

            [ExcelColumn("商品名称")]
            public string _商品名称
            {
                get
                {
                    return org商品名称;
                }
                set
                {
                    org商品名称 = value != null ? value.ToString().Trim() : "";
                }
            }

            [ExcelColumn("商品创建时间")]
            public DateTime _开发时间 { get; set; }

            [ExcelColumn("供应商")]
            public string _供应商 { get; set; }

            [ExcelColumn("业绩归属2")]
            public string _开发 { get; set; }

            [ExcelColumn("采购员")]
            public string _采购员 { get; set; }

            [ExcelColumn("可用数量")]
            public decimal _可用数量 { get; set; }

            [ExcelColumn("库存金额")]
            public decimal _库存金额 { get; set; }
        }

        [ExcelTable("各平台销量一览表")]
        class _各平台销量
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

            public _Enum产品类型 _产品类型 { get; set; }

        }

        class _滞销爆款统计
        {
            public string _开发人员 { get; set; }
            public string _SKU { get; set; }

            public int _滞销SKU个数 { get; set; }

            public int _滞销主SKU个数 { get; set; }

            public decimal _滞销总金额 { get; set; }

            public int _爆款SKU个数 { get; set; }
            public int _爆款主SKU个数 { get; set; }

            public decimal _爆款总金额 { get; set; }

            public int _停售SKU个数 { get; set; }

            public decimal _停售总金额 { get; set; }

            public int _普通款SKU个数 { get; set; }

            public decimal _普通款总金额 { get; set; }

            public int _所有SKU个数
            {
                get
                {
                    return _滞销SKU个数 + _爆款SKU个数 + _停售SKU个数 + _普通款SKU个数;
                }
            }

        }

        class _滞销爆款停售详情
        {
            private _产品信息 _Refer产品信息;
            public string _SKU { get; set; }
            public decimal _可用数量 { get { return Refer产品信息._可用数量; } }
            public decimal _库存金额 { get { return Refer产品信息._库存金额; } }
            public string _供应商 { get { return Refer产品信息._供应商; } }
            public decimal _总销量
            {
                get
                {
                    return _wish销量 + _Ebay销量 + _SMT销量 + _Amazon销量 + _Shopee销量 + _Joom销量;
                }
            }
            public decimal _wish销量 { get; set; }
            public decimal _Ebay销量 { get; set; }
            public decimal _SMT销量 { get; set; }
            public decimal _Amazon销量 { get; set; }
            public decimal _Shopee销量 { get; set; }
            public decimal _Joom销量 { get; set; }
            public _Enum产品类型 _产品类型 { get; set; }

            public _产品信息 Refer产品信息
            {
                get
                {
                    return _Refer产品信息;
                }
                set
                {
                    _Refer产品信息 = value != null ? value : new _产品信息();
                }
            }
        }

        class _爆款滞销统计详情
        {
            public string _开发 { get; set; }
            public string _主SKU { get; set; }

            public int _子SKU个数 { get; set; }

            public decimal _子SKU所有销量 { get; set; }

            public string _所有子SKU销量信息 { get; set; }

        }

        class _供应商统计
        {
            public string _供应商 { get; set; }

            public int _爆款SKU个数 { get; set; }

            public decimal _爆款库存金额 { get; set; }

            public int _滞销SKU个数 { get; set; }

            public decimal _滞销库存金额 { get; set; }

            public int _普通款SKU个数 { get; set; }

            public decimal _普通款库存金额 { get; set; }

            public int _停售SKU个数 { get; set; }

            public decimal _停售库存金额 { get; set; }
        }

        class _SKU统计仓库
        {
            /// <summary>
            /// 所有主SKU
            /// </summary>
            List<string> _MainSKUS;
            List<_主子SKU销售详情> _VolumeDetail;

            public List<string> MainSKUs { get { return _MainSKUS; } }

            #region 构造函数
            public _SKU统计仓库()
            {
                _MainSKUS = new List<string>();
                _VolumeDetail = new List<_主子SKU销售详情>();
            }
            #endregion

            #region AddSaleVolume
            public void AddSaleVolume(string strDeveloper, string strSeedSKU, decimal dSaleVolume)
            {
                var strMainSKU = CutParent(strSeedSKU);
                if (_MainSKUS.Where(x => x == strMainSKU).FirstOrDefault() == null)
                    _MainSKUS.Add(strMainSKU);

                var ref主子详情 = _VolumeDetail.Where(x => x._主SKU == strMainSKU).FirstOrDefault(); ;
                if (ref主子详情 == null)
                {
                    var model = new _主子SKU销售详情();
                    model._主SKU = strMainSKU;
                    model._开发 = strDeveloper;
                    model.AddDetail(strDeveloper, strSeedSKU, dSaleVolume);
                    _VolumeDetail.Add(model);
                }
                else
                {
                    for (int idx = _VolumeDetail.Count - 1; idx >= 0; idx--)
                    {
                        var curItem = _VolumeDetail[idx];
                        if (curItem._主SKU == strMainSKU)
                        {
                            curItem.AddDetail(strDeveloper, strSeedSKU, dSaleVolume);
                        }
                    }
                }

            }
            #endregion

            #region CalcuTotalVolume
            public int CalcuGoodVolume(string strDeveloper, decimal upLimit, ref List<_爆款滞销统计详情> details)
            {
                //details = new List<_爆款滞销统计详情>();
                var res = 0;
                var refListDetails = _VolumeDetail.Where(x => x._开发 == strDeveloper).ToList();
                if (refListDetails != null && refListDetails.Count > 0)
                {
                    var tmp = refListDetails.Where(x => x._所有销量合计 >= upLimit).ToList();
                    foreach (var item in tmp)
                    {
                        //if (item._主SKU=="LGDH9H17")
                        //{

                        //}

                        var model = new _爆款滞销统计详情();
                        model._主SKU = item._主SKU;
                        model._开发 = item._开发;
                        model._子SKU所有销量 = item._所有销量合计;
                        model._子SKU个数 = item._子SKU个数;
                        model._所有子SKU销量信息 = item._所有销量信息;
                        details.Add(model);
                    }
                    res = refListDetails.Where(x => x._所有销量合计 >= upLimit).Count();

                }
                return res;
            }
            #endregion

            #region CalcuBadVolume
            public int CalcuBadVolume(string strDeveloper, decimal lowLimit, ref List<_爆款滞销统计详情> details)
            {
                var res = 0;
                var refListDetails = _VolumeDetail.Where(x => x._开发 == strDeveloper).ToList();
                if (refListDetails != null && refListDetails.Count > 0)
                {
                    var tmp = refListDetails.Where(x => x._所有销量合计 <= lowLimit).ToList();
                    foreach (var item in tmp)
                    {
                        var model = new _爆款滞销统计详情();
                        model._主SKU = item._主SKU;
                        model._开发 = item._开发;
                        model._子SKU所有销量 = item._所有销量合计;
                        model._子SKU个数 = item._子SKU个数;
                        model._所有子SKU销量信息 = item._所有销量信息;
                        details.Add(model);
                    }
                    res = refListDetails.Where(x => x._所有销量合计 <= lowLimit).Count();
                }
                return res;
            }
            #endregion

            /**************** static method ****************/

            /// <summary>
            /// 获取主SKU
            /// </summary>
            /// <param name="strSKU"></param>
            /// <returns></returns>
            static string CutParent(string strSKU)
            {
                if (strSKU.Contains("-"))
                {
                    //sku含有 "-" 第一个"-"前的都是主sku
                    var idx = strSKU.IndexOf("-");
                    if (idx != -1)
                    {
                        return strSKU.Substring(0, idx);
                    }
                }
                else
                {
                    var msg = 0;
                    var iCutLength = 0;
                    if (!string.IsNullOrEmpty(strSKU))
                    {
                        for (int idx = strSKU.Length - 1; idx >= 0; idx--)
                        {
                            var cschar = strSKU.Substring(idx, 1);
                            var bIsNumber = int.TryParse(cschar, out msg);
                            if (bIsNumber)
                            {
                                break;
                            }
                            iCutLength++;
                        }
                    }
                    if (iCutLength > 0)
                    {
                        return strSKU.Substring(0, strSKU.Length - iCutLength);
                    }
                }

                return strSKU;
            }

            class _主子SKU销售详情
            {
                List<KeyValuePair<string, decimal>> _p相关子SKU销售详情;

                public string _主SKU { get; set; }

                public string _开发 { get; set; }
                public decimal _所有销量合计 { get { return GetAllVolume(); } }
                public string _所有销量信息 { get { return GetAllVolumeStr(); } }

                public int _子SKU个数 { get { return GetAllSeedSKUCount(); } }
                public List<KeyValuePair<string, decimal>> _相关子SKU销售详情 { get { return _p相关子SKU销售详情; } }

                #region 构造函数
                public _主子SKU销售详情()
                {
                    _p相关子SKU销售详情 = new List<KeyValuePair<string, decimal>>();
                }
                #endregion

                #region AddDetail
                public void AddDetail(string strDeveloper, string strSeedSKU, decimal dVolume)
                {
                    //_开发 = strDeveloper;
                    _p相关子SKU销售详情.Add(new KeyValuePair<string, decimal>(strSeedSKU, dVolume));
                }
                #endregion

                #region GetAllVolume
                decimal GetAllVolume()
                {
                    return _p相关子SKU销售详情.Select(x => x.Value).Sum();
                }
                #endregion

                #region GetAllVolumeStr
                string GetAllVolumeStr()
                {
                    var tmp = _p相关子SKU销售详情.Select(x => string.Format("{0}*{1}", x.Key, x.Value)).ToArray();
                    return string.Join(";", tmp);
                }
                #endregion

                #region GetAllSeedSKUCount
                int GetAllSeedSKUCount()
                {
                    return _p相关子SKU销售详情.Select(x => x.Key).Distinct().Count();
                }
                #endregion
            }
        }

        enum _Enum产品类型
        {
            _滞销 = 1,
            _爆款 = 2,
            _普通款 = 3,
            _停售 = 4
        }
    }
}
