using Aspose.Slides;
using Calculation.Base;
using Calculation.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    /// <summary>
    /// 按组团、楼盘名称作为分段
    /// </summary>
    public class plus_jp_baoyi : plus_jp_base
    {
        //竞品分布
        //组团周度排名
        //组团业态周度排名
        //组团近八周排名
        //组团业态近近八周排名


        public ISlideCollection _plus_jp_baoyi_1(string str, int cjbh)
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
                    var page2 = temp[1];

                    #region 本周主团项目排名
                    DataTable dt2 = new DataTable();
                    dt2.Columns.Add("排名");
                    dt2.Columns.Add("项目名称");
                    dt2.Columns.Add("成交套数");
                    dt2.Columns.Add("成交金额");
                    dt2.Columns.Add("建面体量");
                    dt2.Columns.Add("套内体量");
                    dt2.Columns.Add("建面均价");
                    dt2.Columns.Add("套内均价");
                    var data1 = (from a in Cache_data_cjjl.bz.AsEnumerable()
                                 where item.ztcs.Contains(a["zt"])
                                 group a by new { lpmc = a["lpmc"] } into s
                                 select new
                                 {
                                     lpmc = s.Key.lpmc,
                                     ts = s.Sum(a => a["ts"].doubls()),
                                     cjje = s.Sum(a => a["cjje"].longs()),
                                     jmtl = s.Sum(a => a["jzmj"].doubls()),
                                     tntl = s.Sum(a => a["tnmj"].doubls()),
                                 }).OrderByDescending(m => m.cjje).Take(10).ToList();

                    for (int i = 0; i < data1.Count; i++)
                    {
                        DataRow dr2 = dt2.NewRow();
                        dr2["排名"] = i + 1;
                        dr2["项目名称"] = data1[i].lpmc;
                        dr2["成交套数"] = data1[i].ts;
                        dr2["成交金额"] = data1[i].cjje.je_wy();
                        dr2["建面体量"] = data1[i].jmtl.mj_wf();
                        dr2["套内体量"] = data1[i].tntl.mj_wf();
                        dr2["建面均价"] = (data1[i].cjje / data1[i].jmtl).je_y();
                        dr2["套内均价"] = (data1[i].cjje / data1[i].tntl).je_y();
                        dt2.Rows.Add(dr2);
                    }
                    Office_Tables.SetTable(page2, dt2, 1, null, null);
                    t.AddClone(page2);



                    foreach (var yt in item.ytcs)
                    {
                        var tp1 = new Presentation(str);
                        var temp1 = tp1.Slides;
                        var page3 = temp1[1];

                        DataTable dt2_1 = new DataTable();
                        dt2_1.Columns.Add("排名");
                        dt2_1.Columns.Add("项目名称");
                        dt2_1.Columns.Add("成交套数");
                        dt2_1.Columns.Add("成交金额");
                        dt2_1.Columns.Add("建面体量");
                        dt2_1.Columns.Add("套内体量");
                        dt2_1.Columns.Add("建面均价");
                        dt2_1.Columns.Add("套内均价");
                        var data2_1 = (from a in Cache_data_cjjl.bz.AsEnumerable()
                                       where item.ztcs.Contains(a["zt"]) && a["yt"].ToString() == yt
                                       group a by new { lpmc = a["lpmc"] } into s
                                       select new
                                       {
                                           lpmc = s.Key.lpmc,
                                           ts = s.Sum(a => a["ts"].doubls()),
                                           cjje = s.Sum(a => a["cjje"].longs()),
                                           jmtl = s.Sum(a => a["jzmj"].doubls()),
                                           tntl = s.Sum(a => a["tnmj"].doubls()),
                                       }).OrderByDescending(m => m.cjje).ToList();
                        for (int i = 0; i < data2_1.Count; i++)
                        {
                            DataRow dr2_1 = dt2_1.NewRow();
                            dr2_1["排名"] = i + 1;
                            dr2_1["项目名称"] = data2_1[i].lpmc;
                            dr2_1["成交套数"] = data2_1[i].ts;
                            dr2_1["成交金额"] = data2_1[i].cjje.je_wy();
                            dr2_1["建面体量"] = data2_1[i].jmtl.mj_wf();
                            dr2_1["套内体量"] = data2_1[i].tntl.mj_wf();
                            dr2_1["建面均价"] = (data2_1[i].cjje / data2_1[i].jmtl).je_y();
                            dr2_1["套内均价"] = (data2_1[i].cjje / data2_1[i].tntl).je_y();
                            dt2_1.Rows.Add(dr2_1);
                        }
                        Office_Tables.SetTable(page3, dt2_1, 1, null, null);
                        t.AddClone(page3);
                    }


                    #endregion

                    #region 组团近八周排名


                    int[] index1 = { 2, 3, 0, 1, 2 };
                    foreach (var ztpmitem in this.JBZ_ZT_PM(str, index1, item.ztcs))
                    {
                        t.AddClone(ztpmitem);
                    }
                    int[] index2 = { 3, 3, 0, 1, 2 };
                    foreach (var ztytpmitem in this.JBZ_ZT_YT_PM(str, index2, item.ztcs, item.ytcs))
                    {
                        t.AddClone(ztytpmitem);
                    }
                    #endregion

                    #region 竞品表现
                    #region 认购
                    var page5 = temp[4];

                    DataTable dt5_0 = new DataTable();

                    dt5_0.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt5_0.Columns.Add(Base_Config_Jzgj.业态);
                    dt5_0.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                    dt5_0.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                    dt5_0.Columns.Add(Base_Config_Rgsj.本周_新开建面均价);
                    dt5_0.Columns.Add(Base_Config_Rgsj.本周_本周来电);
                    dt5_0.Columns.Add(Base_Config_Rgsj.本周_本周到访量);
                    dt5_0.Columns.Add(Base_Config_Rgsj.上上上周_认购套数);
                    dt5_0.Columns.Add(Base_Config_Rgsj.上上上周_认购建面均价);
                    dt5_0.Columns.Add(Base_Config_Rgsj.上上周_认购套数);
                    dt5_0.Columns.Add(Base_Config_Rgsj.上上周_认购建面均价);
                    dt5_0.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                    dt5_0.Columns.Add(Base_Config_Rgsj.上周_认购建面均价);
                    dt5_0.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                    dt5_0.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);
                    dt5_0.Columns.Add("合计认购套数");
                    dt5_0.Columns.Add(Base_Config_Rgsj.本周_变化原因);
                    IAutoShape text5_0 = (IAutoShape)page5.Shapes[0];
                    text5_0.TextFrame.Text = string.Format(text5_0.TextFrame.Text, item.bamc);
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt5_0 = GET_JPXM_BX_RG(dt5_0, item.jpxmlb);
                        Office_Tables.SetJP_baoyi_1_Table(page5, dt5_0, 1, null, null);
                        t.AddClone(page5);
                    }
                    #endregion


                    #region 备案
                    var page6 = temp[5];
                    DataTable dt6_0 = new DataTable();

                    dt6_0.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt6_0.Columns.Add(Base_Config_Jzgj.业态);
                    dt6_0.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间);
                    dt6_0.Columns.Add("主力房型");
                    dt6_0.Columns.Add(Base_Config_Cjba.上上上周_备案套数);
                    dt6_0.Columns.Add(Base_Config_Cjba.上上上周_建面均价);
                    dt6_0.Columns.Add(Base_Config_Cjba.上上周_备案套数);
                    dt6_0.Columns.Add(Base_Config_Cjba.上上周_建面均价);
                    dt6_0.Columns.Add(Base_Config_Cjba.上周_备案套数);
                    dt6_0.Columns.Add(Base_Config_Cjba.上周_建面均价);
                    dt6_0.Columns.Add(Base_Config_Cjba.本周_备案套数);
                    dt6_0.Columns.Add(Base_Config_Cjba.本周_建面均价);
                    dt6_0.Columns.Add("合计认购套数");
                    dt6_0.Columns.Add("合计建面均价");
                    IAutoShape text6 = (IAutoShape)page6.Shapes[0];
                    text6.TextFrame.Text = string.Format(text6.TextFrame.Text, item.bamc);
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt6_0 = GET_JPXM_BX_BA(dt6_0, item.jpxmlb);
                        Office_Tables.SetJP_baoyi_2_Table(page6, dt6_0, 1, null, null);
                        t.AddClone(page6);
                    }
                    #endregion
                    #endregion
                }

                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log(e.Message);
                return null;
            }
        }

        /// <summary>
        /// 认购表现
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="jpxm"></param>
        /// <returns></returns>

        public DataTable GET_JPXM_BX_RG(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
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
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                            var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                            var temp_ba_ssz = temp_rgsj_ssz.FirstOrDefault();
                            var temp_ba_sssz = temp_rgsj_sssz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));

                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        var temp_ba_ssz = temp_rgsj_ssz.FirstOrDefault();
                        var temp_ba_sssz = temp_rgsj_sssz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
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
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            //本周本案认购数据
                            var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                            var temp_ba_ssz = temp_rgsj_ssz.FirstOrDefault();
                            var temp_ba_sssz = temp_rgsj_sssz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                        }
                    }
                    else if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                            var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                            var temp_ba_ssz = temp_rgsj_ssz.FirstOrDefault();
                            var temp_ba_sssz = temp_rgsj_sssz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                        }

                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        var temp_ba_ssz = temp_rgsj_ssz != null && temp_rgsj_ssz.Count() > 0 ? temp_rgsj_ssz.FirstOrDefault() : null;
                        var temp_ba_sssz = temp_rgsj_sssz != null && temp_rgsj_sssz.Count() > 0 ? temp_rgsj_sssz.FirstOrDefault() : null;
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                    } 

                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                    var temp_ba_ssz = temp_rgsj_ssz != null && temp_rgsj_ssz.Count() > 0 ? temp_rgsj_ssz.FirstOrDefault() : null;
                    var temp_ba_sssz = temp_rgsj_sssz != null && temp_rgsj_sssz.Count() > 0 ? temp_rgsj_sssz.FirstOrDefault() : null;
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                }


            }


            return dt;
        }

        /// <summary>
        /// 备案表现
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="jpxm"></param>
        /// <returns></returns>
        public DataTable GET_JPXM_BX_BA(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
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
                            var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据

                            #endregion

                            dt.Rows.Add(GET_ROW_BA_SZ(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));

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
                        #endregion

                        dt.Rows.Add(GET_ROW_BA_SZ(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
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

                            #endregion

                            dt.Rows.Add(GET_ROW_BA_SZ(item.hxcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
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
                            //本周本案认购数据
                            #endregion

                            dt.Rows.Add(GET_ROW_BA_SZ(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
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
                        //本周本案认购数据
                        #endregion

                        dt.Rows.Add(GET_ROW_BA_SZ(string.Join(",", item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
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

                    #endregion

                    dt.Rows.Add(GET_ROW_BA_SZ(string.Join(",", item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                }


            }


            return dt;
        }

   

        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
              DataRow temp_ba_bz,
              DataRow temp_ba_sz,
              DataRow temp_ba_ssz,
              DataRow temp_ba_sssz,
              JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {


                if (Base_Config_Rgsj._认购数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Rgsj.本周_新开销售套数:
                        case Base_Config_Rgsj.本周_新开套数:
                        case Base_Config_Rgsj.本周_认购套数:
                        case Base_Config_Rgsj.本周_认购套内均价:
                        case Base_Config_Rgsj.本周_认购建面均价:
                        case Base_Config_Rgsj.本周_认购套内体量:
                        case Base_Config_Rgsj.本周_认购建面体量:
                        case Base_Config_Rgsj.本周_认购金额:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null ? temp_ba_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.本周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null && temp_ba_bz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_ba_bz[Base_Config_Rgsj.本周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_bz[Base_Config_Rgsj.本周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_bz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
                            }; break;
                        case Base_Config_Rgsj.上周_新开销售套数:
                        case Base_Config_Rgsj.上周_新开套数:
                        case Base_Config_Rgsj.上周_认购套数:
                        case Base_Config_Rgsj.上周_认购套内均价:
                        case Base_Config_Rgsj.上周_认购建面均价:
                        case Base_Config_Rgsj.上周_认购套内体量:
                        case Base_Config_Rgsj.上周_认购建面体量:
                        case Base_Config_Rgsj.上周_认购金额:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_sz != null ? temp_ba_sz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.上周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_sz != null && temp_ba_sz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_ba_sz[Base_Config_Rgsj.本周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_sz[Base_Config_Rgsj.本周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_sz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
                            }; break;
                        case Base_Config_Rgsj.上上周_新开销售套数:
                        case Base_Config_Rgsj.上上周_新开套数:
                        case Base_Config_Rgsj.上上周_认购套数:
                        case Base_Config_Rgsj.上上周_认购套内均价:
                        case Base_Config_Rgsj.上上周_认购建面均价:
                        case Base_Config_Rgsj.上上周_认购套内体量:
                        case Base_Config_Rgsj.上上周_认购建面体量:
                        case Base_Config_Rgsj.上上周_认购金额:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_ssz != null ? temp_ba_ssz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.上上周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_ssz != null && temp_ba_ssz[Base_Config_Rgsj.上上周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_ba_ssz[Base_Config_Rgsj.上上周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_ssz[Base_Config_Rgsj.上上周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_ssz[Base_Config_Rgsj.上上周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
                            }; break;
                        case Base_Config_Rgsj.上上上周_新开销售套数:
                        case Base_Config_Rgsj.上上上周_新开套数:
                        case Base_Config_Rgsj.上上上周_认购套数:
                        case Base_Config_Rgsj.上上上周_认购套内均价:
                        case Base_Config_Rgsj.上上上周_认购建面均价:
                        case Base_Config_Rgsj.上上上周_认购套内体量:
                        case Base_Config_Rgsj.上上上周_认购建面体量:
                        case Base_Config_Rgsj.上上上周_认购金额:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_sssz != null ? temp_ba_sssz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.上上上周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_sssz != null && temp_ba_sssz[Base_Config_Rgsj.上上上周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_ba_sssz[Base_Config_Rgsj.上上上周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_sssz[Base_Config_Rgsj.上上上周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_sssz[Base_Config_Rgsj.上上上周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
                            }; break;
                        default:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null ? temp_ba_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : "-";
                            }; break;
                    }
                }
                else if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                {

                }
                else if (Base_Config_Jzgj._竞争格局参数名称.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Jzgj.组团: { dr1[dt.Columns[j].ColumnName] = item.ztcs[0]; }; break;
                        case Base_Config_Jzgj.项目名称: { dr1[dt.Columns[j].ColumnName] = item.lpcs[0]; }; break;
                        case Base_Config_Jzgj.业态: { dr1[dt.Columns[j].ColumnName] = yt; }; break;
                        case Base_Config_Jzgj.竞争格局名称: { dr1[dt.Columns[j].ColumnName] = item.jzgjmc; }; break;
                        case Base_Config_Jzgj.竞争格局_主力面积区间: { dr1[dt.Columns[j].ColumnName] = item.zlmjqj; }; break;
                        default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                    }

                }

            }

            return dr1;
        }

    }
}
