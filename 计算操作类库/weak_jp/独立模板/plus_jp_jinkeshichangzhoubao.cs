using Aspose.Slides;
using Calculation.Base;
using Calculation.Dal;
using Calculation.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    public class plus_jp_jinkeshichangzhoubao : plus_jp_base
    {
        private static DataTable bn;

        public class Base_Config_Cjba_Qn
        {

            public const string 全年_累计成交建面 = "qn_jzmj";
            public const string 全年_累计成交金额 = "qn_cjje";
            public const string 全年_累计套内均价 = "qn_tnjj";



            public static string[] _全年备案数据 = { "qn_jzmj", "qn_cjje", "qn_tnjj" };
        }

        public plus_jp_jinkeshichangzhoubao()
        {
            DateTime start = new DateTime(Base_date.bn, 1, 1);
            bn = ZB_Data_CJBA_DataProvider.GET_ZB(start, Base_date.bz_Last);
        }

        public ISlideCollection _plus_jp_jinkeshichangzhoubao_1(string str, int cjbh)
        {
            try
            {

                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);
                foreach (var item in param)
                {
                    var tp = new Presentation(str);
                    var temp = tp.Slides;
                    if (item.ytcs.IsNotNull()&&item.ytcs.Contains("商铺"))
                    {
                        var page1 = temp[1];
                        IAutoShape text1 = (IAutoShape)page1.Shapes[0];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc);
                        DataTable dt = new DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);
                        dt.Columns.Add(Base_Config_Cjba_Qn.全年_累计成交金额);
                        dt.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.上周_套内面积);
                        dt.Columns.Add(Base_Config_Cjba.上周_套内均价);
                        dt.Columns.Add(Base_Config_Cjba.上周_成交金额);

                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.本周_套内面积);
                        dt.Columns.Add(Base_Config_Cjba.本周_套内均价);
                        dt.Columns.Add(Base_Config_Cjba.本周_成交金额);
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            //获取竞品项目数据
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_JINKESHICHANGZHOUBAO_2_Table(page1, dt, 1, null, null);
                            t.AddClone(page1);
                        }
                    }
                    else
                    {
                        var page1 = temp[0];
                        IAutoShape text1 = (IAutoShape)page1.Shapes[0];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc);
                        DataTable dt = new DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);
                        dt.Columns.Add("可售存量");
                        dt.Columns.Add(Base_Config_Cjba_Qn.全年_累计成交建面);
                        dt.Columns.Add(Base_Config_Cjba_Qn.全年_累计成交金额);
                        dt.Columns.Add(Base_Config_Cjba_Qn.全年_累计套内均价);

                        dt.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.上周_套内面积);
                        dt.Columns.Add(Base_Config_Cjba.上周_套内均价);
                        dt.Columns.Add(Base_Config_Cjba.上周_成交金额);

                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.本周_套内面积);
                        dt.Columns.Add(Base_Config_Cjba.本周_套内均价);
                        dt.Columns.Add(Base_Config_Cjba.本周_成交金额);
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            //获取竞品项目数据
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_JINKESHICHANGZHOUBAO_1_Table(page1, dt, 1, null, null);
                            t.AddClone(page1);
                        }
                    }
                }
                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log(e.StackTrace);
                return null;
            }
        }

        /// <summary>
        /// 获取竞品项目，数据维度为备案
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="jpxm"></param>
        /// <returns></returns>
        public DataTable GET_JPXM_BX(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm.OrderByDescending(m=>m.id))
            {

                if (item.ytcs[0] == "别墅")
                {
                    if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {

                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_bn = bn.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据

                            #endregion

                            dt.Rows.Add(GET_ROW_BA1(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, null, null, temp_ba_bn, item));

                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_bn = bn.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        #endregion

                        dt.Rows.Add(GET_ROW_BA1(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, temp_ba_bn, item));
                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    if (item.hxcs != null & item.hxcs.Length > 0)
                    {
                        for (int i = 0; i < item.hxcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                             var temp_ba_bn = bn.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            #endregion

                            dt.Rows.Add(GET_ROW_BA1(item.hxcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, temp_ba_bn, item));
                        }
                    }
                    else if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_bn = bn.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                            #endregion

                            dt.Rows.Add(GET_ROW_BA1(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, temp_ba_bn, item));
                        }

                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_bn = bn.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        //本周本案认购数据
                        #endregion

                        dt.Rows.Add(GET_ROW_BA1(string.Join(",", item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, temp_ba_bn, item));
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    //竞品业态
                    var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_ba_bn = bn.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));

                    #endregion

                    dt.Rows.Add(GET_ROW_BA1(string.Join(",", item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, temp_ba_bn, item));
                }


            }


            return dt;
        }

        public  DataRow GET_ROW_BA1(string yt, DataRow dr1, System.Data.DataTable dt,
                             EnumerableRowCollection<DataRow> temp_ba_bz,
                             EnumerableRowCollection<DataRow> temp_ba_sz,
                             EnumerableRowCollection<DataRow> temp_ba_ssz,
                             EnumerableRowCollection<DataRow> temp_ba_sssz,
                             EnumerableRowCollection<DataRow> temp_ba_bn,
                             JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {

                try
                {
                    #region _备案数据


                    if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Cjba.本周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null ? temp_ba_bz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.本周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null ? temp_ba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()).je_wy() : 0; }; break;
                            case Base_Config_Cjba.本周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null ? temp_ba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()).mj() : 0; }; break;
                            case Base_Config_Cjba.本周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null ? temp_ba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls()).mj() : 0; }; break;
                            case Base_Config_Cjba.本周_建面均价:
                                {

                                    if ((temp_ba_bz != null && temp_ba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_ba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_ba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Cjba.本周_套内均价:
                                {
                                    if ((temp_ba_bz != null && temp_ba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_ba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_ba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Cjba.上周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_ba_sz != null ? temp_ba_sz.Sum(m => m[Base_Config_Cjba.上周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.上周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_ba_sz != null ? temp_ba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()).je_wy() : 0; }; break;
                            case Base_Config_Cjba.上周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_ba_sz != null ? temp_ba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()).mj(): 0; }; break;
                            case Base_Config_Cjba.上周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_ba_sz != null ? temp_ba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls()).mj() : 0; }; break;
                            case Base_Config_Cjba.上周_建面均价:
                                {
                                    if ((temp_ba_sz != null && temp_ba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_ba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) / temp_ba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Cjba.上周_套内均价:
                                {
                                    if ((temp_ba_sz != null && temp_ba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_ba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) / temp_ba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                           
                            case Base_Config_Cjba.本周_备案套数环比:
                                {
                                    dr1[dt.Columns[j].ColumnName] = ((temp_ba_bz.Sum(m => m["ts"].ints()) - temp_ba_sz.Sum(m => m["ts"].ints())) / temp_ba_sz.Sum(m => m["ts"].ints())).doubls().ss_bfb();
                                }; break;
                            case Base_Config_Cjba.本周_套内均价环比:
                                {
                                    long bz_cjje = temp_ba_bz.Sum(m => m["cjje"].ints());
                                    long bz_tnmj = temp_ba_bz.Sum(m => m["tnmj"].ints());
                                    long sz_cjje = temp_ba_sz.Sum(m => m["cjje"].ints());
                                    long sz_tnmj = temp_ba_sz.Sum(m => m["tnmj"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_cjje / bz_tnmj - sz_cjje / sz_tnmj) / (sz_cjje / sz_tnmj)).doubls().ss_bfb();
                                }; break;
                            case Base_Config_Cjba.本周_套内面积环比:
                                {
                                    long bz_tnmj = temp_ba_bz.Sum(m => m["tnmj"].ints());
                                    long sz_tnmj = temp_ba_sz.Sum(m => m["tnmj"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_tnmj - sz_tnmj) / (sz_tnmj)).doubls().ss_bfb();
                                }; break;
                            case Base_Config_Cjba.本周_套均总价环比:
                                {
                                    long bz_cjje = temp_ba_bz.Sum(m => m["cjje"].ints());
                                    long bz_ts = temp_ba_bz.Sum(m => m["ts"].ints());
                                    long sz_cjje = temp_ba_sz.Sum(m => m["cjje"].ints());
                                    long sz_ts = temp_ba_sz.Sum(m => m["ts"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_cjje / bz_ts - sz_cjje / sz_ts) / (sz_cjje / sz_ts)).doubls().ss_bfb();
                                }; break;

                            case Base_Config_Cjba.本周_建筑面积环比:
                                {
                                    long bz_cjje = temp_ba_bz.Sum(m => m["cjje"].ints());
                                    long bz_jzmj = temp_ba_bz.Sum(m => m["jzmj"].ints());
                                    long sz_cjje = temp_ba_sz.Sum(m => m["cjje"].ints());
                                    long sz_jzmj = temp_ba_sz.Sum(m => m["jzmj"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_cjje / bz_jzmj - sz_cjje / sz_jzmj) / (sz_cjje / sz_jzmj)).doubls().ss_bfb();
                                }; break;
                            case Base_Config_Cjba.本周_建面均价环比:
                                {
                                    long bz_cjje = temp_ba_bz.Sum(m => m["cjje"].ints());
                                    long bz_jzmj = temp_ba_bz.Sum(m => m["jzmj"].ints());
                                    long sz_cjje = temp_ba_sz.Sum(m => m["cjje"].ints());
                                    long sz_jzmj = temp_ba_sz.Sum(m => m["jzmj"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_cjje / bz_jzmj - sz_cjje / sz_jzmj) / (sz_cjje / sz_jzmj)).doubls().ss_bfb();
                                }; break;
                            case Base_Config_Cjba.本周_成交金额环比:
                                {
                                    long bz_cjje = temp_ba_bz.Sum(m => m["cjje"].ints());
                                    long sz_cjje = temp_ba_sz.Sum(m => m["cjje"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_cjje - sz_cjje) / sz_cjje).doubls().ss_bfb();
                                }; break;



                            default: { dr1[dt.Columns[j].ColumnName] = "0"; }; break;
                        }


                    }
                    #endregion
                    #region _竞争格局参数名称


                    else if (Base_Config_Jzgj._竞争格局参数名称.Contains(dt.Columns[j].ColumnName))
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Jzgj.项目名称:
                                {
                                    dr1[dt.Columns[j].ColumnName] = item.lpcs[0];
                                }; break;
                            case Base_Config_Jzgj.业态:
                                {
                                    dr1[dt.Columns[j].ColumnName] = yt;
                                }; break;
                            case Base_Config_Jzgj.组团:
                                {
                                    dr1[dt.Columns[j].ColumnName] = item != null && item.ztcs != null ? string.Join(",", item.ztcs) : "";
                                }; break;
                            case Base_Config_Jzgj.竞争格局_主力面积区间:
                                {
                                    dr1[dt.Columns[j].ColumnName] = item.zlmjqj;
                                }; break;
                            case Base_Config_Jzgj.竞争格局名称:
                                {
                                    dr1[dt.Columns[j].ColumnName] = item.jzgjmc;
                                }; break;
                        }

                    }
                    #endregion

                    else if (Base_Config_Cjba_Qn._全年备案数据.Contains(dt.Columns[j].ColumnName))
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Cjba_Qn.全年_累计套内均价:
                                {
                                    double tnmj = temp_ba_bn.Sum(m => m["tnmj"].doubls());
                                    dr1[dt.Columns[j].ColumnName] = tnmj != 0 ? (temp_ba_bn.Sum(m => m["cjje"].longs()) / tnmj).je_y().ToString() : "-";
                                }; break;
                            case Base_Config_Cjba_Qn.全年_累计成交建面:
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_bn.Sum(m => m["jzmj"].doubls()).mj_wf();
                                }; break;
                            case Base_Config_Cjba_Qn.全年_累计成交金额:
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_bn.Sum(m => m["cjje"].longs()).je_wy();
                                }; break;
                        }
                    } 
                }
                catch (Exception e)
                {

                    throw e;
                }
            }

            return dr1;
        }
    }
}
