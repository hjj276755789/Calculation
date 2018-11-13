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
    /// 贝蒙置地
    /// </summary>
    public class plus_jp_beimengzhidi:plus_jp_base
    {
        public ISlideCollection _plus_jp_beimengzhidi_1(string str, int cjbh)
        {
            try
            {

                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);
                foreach (var item in param)
                {

                    
                    #region 市场量价

                    
                    if (item.qtcs== "市场量价")
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page1 = temp[0];

                      
                        #region 商品房
                        var jbz_cjba_spf = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                                           where a["zt"].ToString() == item.ztcs[0] 
                                           group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                                           select new
                                           {
                                               zc = s.Key.zc,
                                               zcmc = s.Key.zcmc,
                                               cjje = s.Sum(a => a["cjje"].longs()),
                                               jzmj = s.Sum(a => a["jzmj"].doubls()),
                                           }).OrderBy(m => m.zc).ToList();
                        var jbz_xzys_spf = (from a in Cache_data_xzys.jbz.AsEnumerable()
                                           where a["zt"].ToString() == item.ztcs[0]
                                           group a by new { zc = a["zc"] } into s
                                           select new
                                           {
                                               zc = s.Key.zc,
                                               xzgy = s.Sum(a => a["jzmj"].doubls())+s.Sum(a => a["fzzmj"].doubls()),
                                           }).OrderBy(m => m.zc).ToList();
                        var temp_spf = (from a in jbz_cjba_spf
                                        join b in jbz_xzys_spf on a.zc equals b.zc into tempdata
                                       from tt in tempdata.DefaultIfEmpty()
                                       select new
                                       {
                                           zcmc = a.zcmc,
                                           xzgyl = tt == null ? 0 : tt.xzgy,//这里主要第二个集合有可能为空。需要判断
                                           cjmj = a.jzmj,
                                           jmjj = a.cjje / a.jzmj
                                       }).ToList();
                        DataTable dt_spf = new DataTable();
                        dt_spf.Columns.Add("周次名称");
                        dt_spf.Columns.Add("供应体量(万方）");
                        dt_spf.Columns.Add("成交体量(万方）");
                        dt_spf.Columns.Add("建面均价");
                        foreach (var item_spf in temp_spf)
                        {
                            DataRow dr = dt_spf.NewRow();
                            dr["周次名称"] = item_spf.zcmc;
                            dr["供应体量(万方）"] = item_spf.xzgyl.mj_wf();
                            dr["成交体量(万方）"] = item_spf.cjmj.mj_wf();
                            dr["建面均价"] = item_spf.jmjj.je_y();
                            dt_spf.Rows.Add(dr);
                        }
                       
                        Office_Charts.Chart_gxfx(page1, dt_spf, 3);
                        #endregion

                        #region 商品住宅
                        var jbz_cjba_zz = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                                        where a["zt"].ToString() == item.ztcs[0] && (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                                        select new
                                        {
                                            zc = s.Key.zc,
                                            zcmc = s.Key.zcmc,
                                            cjje = s.Sum(a => a["cjje"].longs()),
                                            jzmj = s.Sum(a => a["jzmj"].doubls()),
                                        }).OrderBy(m => m.zc).ToList();
                        var jbz_xzys_zz = (from a in Cache_data_xzys.jbz.AsEnumerable()
                                        where a["zt"].ToString() == item.ztcs[0] && (a["tyyt"].ToString() == "别墅" || a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "小高层" || a["tyyt"].ToString() == "洋房" || a["tyyt"].ToString() == "洋楼")
                                        group a by new { zc = a["zc"] } into s
                                        select new
                                        {
                                            zc = s.Key.zc,
                                            xzgy = s.Sum(a => a["jzmj"].doubls()),
                                        }).OrderBy(m => m.zc).ToList();
                        var temp_zz = (from a in jbz_cjba_zz
                                     join b in jbz_xzys_zz on a.zc equals b.zc into tempdata
                                     from tt in tempdata.DefaultIfEmpty()
                                     select new
                                     {
                                         zcmc = a.zcmc,
                                         xzgyl = tt == null ? 0 : tt.xzgy,//这里主要第二个集合有可能为空。需要判断
                                         cjmj = a.jzmj,
                                         jmjj = a.cjje / a.jzmj
                                     }).ToList();
                        DataTable dt_zz = new DataTable();
                        dt_zz.Columns.Add("周次名称");
                        dt_zz.Columns.Add("供应体量(万方）");
                        dt_zz.Columns.Add("成交体量(万方）");
                        dt_zz.Columns.Add("建面均价");
                        foreach (var item_zz in temp_zz)
                        {
                            DataRow dr = dt_zz.NewRow();
                            dr["周次名称"] = item_zz.zcmc;
                            dr["供应体量(万方）"] = item_zz.xzgyl.mj_wf();
                            dr["成交体量(万方）"] = item_zz.cjmj.mj_wf();
                            dr["建面均价"] = item_zz.jmjj.je_y();
                            dt_zz.Rows.Add(dr);
                        }
                      
                        Office_Charts.Chart_gxfx(page1, dt_zz, 4);
                        #endregion

                        IAutoShape text0_1 = (IAutoShape)page1.Shapes[0];
                        text0_1.TextFrame.Text = string.Format(text0_1.TextFrame.Text, item.ztcs[0], temp_spf[7].xzgyl.mj_wf(), temp_spf[7].cjmj.mj_wf(), temp_spf[7].jmjj.je_y());
                        IAutoShape text0_2 = (IAutoShape)page1.Shapes[1];
                        text0_2.TextFrame.Text = string.Format(text0_2.TextFrame.Text, item.ztcs[0], temp_zz[7].xzgyl.mj_wf(), temp_zz[7].cjmj.mj_wf(), temp_zz[7].jmjj.je_y());
                        IAutoShape text0_3 = (IAutoShape)page1.Shapes[2];
                        text0_3.TextFrame.Text = string.Format(text0_3.TextFrame.Text, item.bamc);
                        t.AddClone(page1);
                    }
                    #endregion

                 


                    else
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page2 = temp[1];
                        IAutoShape text1_1 = (IAutoShape)page2.Shapes[0];
                        text1_1.TextFrame.Text = string.Format(text1_1.TextFrame.Text, item.bamc);
                        IAutoShape text1_2 = (IAutoShape)page2.Shapes[2];
                        text1_2.TextFrame.Text = string.Format(text1_2.TextFrame.Text, item.bamc);
                        #region 市场成交走势
                        #region 商务
                        if (item.ytcs[0] == "商务")
                        {
                            var jbz_cjba_sw = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                                                where  (a["yt"].ToString() == "商务"&& a["xfyt"].ToString()=="商务公寓" && a["hx"].ToString()=="LOFT" )
                                                group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                                                select new
                                                {
                                                    zc = s.Key.zc,
                                                    zcmc = s.Key.zcmc,
                                                    cjje = s.Sum(a => a["cjje"].longs()),
                                                    jzmj = s.Sum(a => a["jzmj"].doubls()),
                                                }).OrderBy(m => m.zc).ToList();
                            var jbz_xzys_sw = (from a in Cache_data_xzys.jbz.AsEnumerable()
                                                where a["wylx"].ToString()=="SOHO" || a["wylx"].ToString()=="LOFT"
                                                group a by new { zc = a["zc"] } into s
                                                select new
                                                {
                                                    zc = s.Key.zc,
                                                    xzgy = s.Sum(a => a["jzmj"].doubls()) + s.Sum(a => a["fzzmj"].doubls()),
                                                }).OrderBy(m => m.zc).ToList();
                            var temp_spf = (from a in jbz_cjba_sw
                                            join b in jbz_xzys_sw on a.zc equals b.zc into tempdata
                                            from tt in tempdata.DefaultIfEmpty()
                                            select new
                                            {
                                                zcmc = a.zcmc,
                                                xzgyl = tt == null ? 0 : tt.xzgy,//这里主要第二个集合有可能为空。需要判断
                                                cjmj = a.jzmj,
                                                jmjj = a.cjje / a.jzmj
                                            }).ToList();
                            DataTable dt_spf = new DataTable();
                            dt_spf.Columns.Add("周次名称");
                            dt_spf.Columns.Add("供应体量（㎡）");
                            dt_spf.Columns.Add("成交体量（㎡）");
                            dt_spf.Columns.Add("建面均价（元/㎡）");
                            foreach (var item_spf in temp_spf)
                            {
                                DataRow dr = dt_spf.NewRow();
                                dr["周次名称"] = item_spf.zcmc;
                                dr["供应体量（㎡）"] = item_spf.xzgyl.mj();
                                dr["成交体量（㎡）"] = item_spf.cjmj.mj();
                                dr["建面均价（元/㎡）"] = item_spf.jmjj.je_y();
                                dt_spf.Rows.Add(dr);
                            }

                            Office_Charts.Chart_gxfx(page2, dt_spf, 1);

                            var bz_cjba_sw_pm = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                                               where (a["yt"].ToString() == "商务" && a["xfyt"].ToString() == "商务公寓" && a["hx"].ToString() == "LOFT" &&  a["zc"].ints() == Base_date.bz && a["nf"].ints() == Base_date.bn) 
                                               group a by new { lpmc = a["lpmc"],qy=a["qy"] } into s
                                               select new
                                               {
                                                   lpmc = s.Key.lpmc,
                                                    qy =s.Key.qy,
                                                   cjts = s.Count(),
                                                   jzmj = s.Sum(a => a["jzmj"].doubls()).mj(),
                                                   cjje = s.Sum(a=>a["cjje"].doubls()).je_wy(),
                                                   jmjj = (s.Sum(a => a["cjje"].doubls())/ s.Sum(a => a["jzmj"].doubls())).je_y()
                                               }).OrderByDescending(m => m.cjje).Take(10).ToList();
                            DataTable pm_tb = new DataTable();
                            pm_tb.Columns.Add("排名");
                            pm_tb.Columns.Add("项目名称");
                            pm_tb.Columns.Add("区域");
                            pm_tb.Columns.Add("成交套数");
                            pm_tb.Columns.Add("成交面积");
                            pm_tb.Columns.Add("成交金额");
                            pm_tb.Columns.Add("成交建均");

                            for (int i = 0; i < bz_cjba_sw_pm.Count(); i++)
                            {
                                DataRow drss = pm_tb.NewRow();
                                drss["排名"] = i+1;
                                drss["项目名称"] = bz_cjba_sw_pm[i].lpmc;
                                drss["区域"] = bz_cjba_sw_pm[i].qy;
                                drss["成交套数"] = bz_cjba_sw_pm[i].cjts;
                                drss["成交面积"] = bz_cjba_sw_pm[i].jzmj;
                                drss["成交金额"] = bz_cjba_sw_pm[i].cjje;
                                drss["成交建均"] = bz_cjba_sw_pm[i].jmjj;
                                pm_tb.Rows.Add(drss);
                            }
                            Office_Tables.SetTable(page2, pm_tb, 3,null,null);
                            t.AddClone(page2);
                        }
                        #endregion
                        #region 商铺
                        else if (item.ytcs[0]=="商铺")
                        {

                            var jbz_cjba_spf = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                                                where a["qy"].ToString() == item.qycs[0] && a["yt"].ToString() == item.ytcs[0] 
                                                group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                                                select new
                                                {
                                                    zc = s.Key.zc, 
                                                    zcmc = s.Key.zcmc,
                                                    cjje = s.Sum(a => a["cjje"].longs()),
                                                    jzmj = s.Sum(a => a["jzmj"].doubls()),
                                                }).OrderBy(m => m.zc).ToList();
                            var jbz_xzys_spf = (from a in Cache_data_xzys.jbz.AsEnumerable()
                                                where a["tyyt"].ToString() == item.ytcs[0] && a["qy"].ToString() == item.qycs[0]
                                                group a by new { zc = a["zc"] } into s
                                                select new
                                                {
                                                    zc = s.Key.zc,
                                                    xzgy = s.Sum(a => a["jzmj"].doubls()) + s.Sum(a => a["fzzmj"].doubls()),
                                                }).OrderBy(m => m.zc).ToList();
                            var temp_spf = (from a in jbz_cjba_spf
                                            join b in jbz_xzys_spf on a.zc equals b.zc into tempdata
                                            from tt in tempdata.DefaultIfEmpty()
                                            select new
                                            {
                                                zcmc = a.zcmc,
                                                xzgyl = tt == null ? 0 : tt.xzgy,//这里主要第二个集合有可能为空。需要判断
                                                cjmj = a.jzmj,
                                                jmjj = a.cjje / a.jzmj
                                            }).ToList();
                            DataTable dt_spf = new DataTable();
                            dt_spf.Columns.Add("周次名称");
                            dt_spf.Columns.Add("供应体量(万方）");
                            dt_spf.Columns.Add("成交体量(万方）");
                            dt_spf.Columns.Add("建面均价");
                            foreach (var item_spf in temp_spf)
                            {
                                DataRow dr = dt_spf.NewRow();
                                dr["周次名称"] = item_spf.zcmc;
                                dr["供应体量(万方）"] = item_spf.xzgyl.mj_wf();
                                dr["成交体量(万方）"] = item_spf.cjmj.mj_wf();
                                dr["建面均价"] = item_spf.jmjj.je_y();
                                dt_spf.Rows.Add(dr);
                            }

                            Office_Charts.Chart_gxfx(page2, dt_spf, 1);
                            var bz_cjba_sw_pm = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                                                 where a["qy"].ToString() == item.qycs[0] && a["yt"].ToString() == item.ytcs[0] && a["zc"].ints() == Base_date.bz && a["nf"].ints() == Base_date.bn
                                                 group a by new { lpmc = a["lpmc"], zt = a["zt"] } into s
                                                 select new
                                                 {
                                                     lpmc = s.Key.lpmc,
                                                     zt = s.Key.zt,
                                                     cjts = s.Count(),
                                                     jzmj = s.Sum(a => a["jzmj"].doubls()).mj(),
                                                     cjje = s.Sum(a => a["cjje"].doubls()).je_wy(),
                                                     jmjj = (s.Sum(a => a["cjje"].doubls()) / s.Sum(a => a["jzmj"].doubls())).je_y()
                                                 }).OrderByDescending(m => m.cjje).Take(10).ToList();
                            DataTable pm_tb = new DataTable();
                            pm_tb.Columns.Add("排名");
                            pm_tb.Columns.Add("项目名称");
                            pm_tb.Columns.Add("区域");
                            pm_tb.Columns.Add("成交套数");
                            pm_tb.Columns.Add("成交面积");
                            pm_tb.Columns.Add("成交金额");
                            pm_tb.Columns.Add("成交建均");

                            for (int i = 0; i < bz_cjba_sw_pm.Count(); i++)
                            {
                                DataRow drss = pm_tb.NewRow();
                                drss["排名"] = i + 1;
                                drss["项目名称"] = bz_cjba_sw_pm[i].lpmc;
                                drss["区域"] = bz_cjba_sw_pm[i].zt;
                                drss["成交套数"] = bz_cjba_sw_pm[i].cjts;
                                drss["成交面积"] = bz_cjba_sw_pm[i].jzmj;
                                drss["成交金额"] = bz_cjba_sw_pm[i].cjje;
                                drss["成交建均"] = bz_cjba_sw_pm[i].jmjj;
                                pm_tb.Rows.Add(drss);
                            }
                            Office_Tables.SetTable(page2, pm_tb, 3, null, null);
                            t.AddClone(page2);
                        }
                        #endregion

                        #endregion
                      

                        #region 典型竞争项目
                        var page3 = temp[2];
                        DataTable dt_jzxm = new DataTable();
                        dt_jzxm.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt_jzxm.Columns.Add(Base_Config_Jzgj.业态);
                        dt_jzxm.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间);
                        dt_jzxm.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt_jzxm.Columns.Add(Base_Config_Cjba.本周_建面均价);
                        dt_jzxm.Columns.Add(Base_Config_Cjba.本周_套均总价);
                        dt_jzxm.Columns.Add("产品、配置及营销");

                       
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            IAutoShape text2 = (IAutoShape)page3.Shapes[0];
                            text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc);
                            dt_jzxm = GET_JPXM_BX(dt_jzxm, item.jpxmlb);
                            Office_Tables.SetJP_BEIMENGZHIDI_JINGZHENGXIANGMU_Table(page3, dt_jzxm, 1, null, null);
                            t.AddClone(page3);
                        }
                        #endregion
                    }






                }
                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log(e.Message);
                return null;
            }
        }

        public System.Data.DataTable GET_JPXM_BX(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {

                if (item.ytcs[0] == "别墅")
                {
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {

                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));

                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    for (int i = 0; i < item.hxcs.Length; i++)
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.hxcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));
                }


            }


            return dt;
        }
    }
}
