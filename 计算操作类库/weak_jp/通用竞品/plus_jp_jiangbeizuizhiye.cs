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
    /// 江北嘴置业住宅竞品
    /// </summary>
    public class plus_jp_jiangbeizuizhiye : plus_jp_base
    {
        public string qy = "江北区";
        public ISlideCollection _plus_jp_jiangbeizuizhiye_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);


  
                var pages = new Presentation(str).Slides;
                var jbz = pages[0];

                #region 近8周江北区住宅市场环境
                DataTable zzsc = new DataTable();
                zzsc.Columns.Add("时间");
                zzsc.Columns.Add("预售新增供应量（单位: 万㎡）");
                zzsc.Columns.Add("成交量（单位: 万㎡）");
                zzsc.Columns.Add("建面均价（元 /㎡）");
                var jbz_cjba = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                                where a["qy"].ToString() == qy && (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                                group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                               select new
                               {
                                   zc = s.Key.zc,
                                   zcmc = s.Key.zcmc,
                                   cjje = s.Sum(a => a["cjje"].longs()),
                                   jzmj = s.Sum(a => a["jzmj"].doubls()),
                               }).OrderBy(m=>m.zc).ToList();
                var jbz_xzys = (from a in Cache_data_xzys.jbz.AsEnumerable()
                               where a["qx1"].ToString() == qy && (a["tyyt"].ToString() == "别墅" || a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "小高层" || a["tyyt"].ToString() == "洋房" || a["tyyt"].ToString() == "洋楼")
                               group a by new { zc = a["zc"] } into s
                               select new
                               {
                                   zc = s.Key.zc,
                                   xzgy = s.Sum(a => a["jzmj"].doubls()),
                               }).OrderBy(m => m.zc).ToList();
                var temp6 = (from a in jbz_cjba
                             join b in jbz_xzys on a.zc equals b.zc into temp
                             from tt in temp.DefaultIfEmpty()
                             select new
                             {
                                 zcmc = a.zcmc,
                                 xzgyl = tt == null ? 0 : tt.xzgy,//这里主要第二个集合有可能为空。需要判断
                                 cjmj = a.jzmj,
                                 jmjj = a.cjje / a.jzmj
                             }).ToList();
                for (int i = 0; i < temp6.Count(); i++)
                {
                    DataRow dr = zzsc.NewRow();
                    dr[0] = temp6[i].zcmc;
                    dr[1] = temp6[i].xzgyl.mj_wf();
                    dr[2] = temp6[i].cjmj.mj_wf();
                    dr[3] = temp6[i].jmjj.je_y();
                    zzsc.Rows.Add(dr);
                }
                Office_Charts.Chart_gxfx(jbz, zzsc, 1);
                t.AddClone(jbz);
                #endregion

                #region 江北区周度住宅排名
                var temp_data_cj = from a in Cache_data_cjjl.bz.AsEnumerable()
                                 where a["qy"].ToString()== qy && (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                                 group a by new { lpmc= a["lpmc"] } into d
                                 select new
                                 {
                                     lpmc = d.Key.lpmc,
                                     cjts = d.Sum(m=>m["ts"].ints()),
                                     cjtl = d.Sum(m => m["jzmj"].doubls()),
                                     cjje = d.Sum(m => m["cjje"].doubls())
                                 };
                var cjpm_ts = temp_data_cj.OrderByDescending(m => m.cjts).Take(10).ToList();
                var cjpm_mj = temp_data_cj.OrderByDescending(m => m.cjtl).Take(10).ToList();
                var cjpm_je = temp_data_cj.OrderByDescending(m => m.cjje).Take(10).ToList();
                DataTable cjpm = new DataTable();
                cjpm.Columns.Add("序号");
                cjpm.Columns.Add("项目名称1");
                cjpm.Columns.Add("套数");
                cjpm.Columns.Add("项目名称2");
                cjpm.Columns.Add("成交面积");
                cjpm.Columns.Add("项目名称3");
                cjpm.Columns.Add("成交金额");
                for (int i = 0; i < 10; i++)
                {
                    DataRow dr = cjpm.NewRow();
                    dr["序号"] = i + 1;
                    if (cjpm_ts.Count()>i)
                    {
                        dr["项目名称1"] = cjpm_ts[i].lpmc;
                        dr["套数"] = cjpm_ts[i].cjts;
                    }
                    else
                    {
                        dr["项目名称1"] = "";
                        dr["套数"] = "" ;
                    }

                    if (cjpm_mj.Count() > i)
                    {
                        dr["项目名称2"] = cjpm_ts[i].lpmc;
                        dr["成交面积"] = cjpm_mj[i].cjtl.ints();
                    }
                    else
                    {
                        dr["项目名称2"] = "";
                        dr["成交面积"] = "";
                    }

                    if (cjpm_je.Count() > i)
                    {
                        dr["项目名称3"] = cjpm_ts[i].lpmc;
                        dr["成交金额"] = cjpm_je[i].cjje.je_wy();
                    }
                    else
                    {
                        dr["项目名称3"] = "";
                        dr["成交金额"] = "";
                    }
                    cjpm.Rows.Add(dr);
                }
                var cjpmp_temp = pages;
                var cjpmp_page = cjpmp_temp[1];

                IAutoShape cjpmwz = (IAutoShape)cjpmp_page.Shapes[2];
                cjpmwz.TextFrame.Text = string.Format(cjpmwz.TextFrame.Text, Base_date.GET_ZCMC(Base_date.bn, Base_date.bz));
                Office_Tables.SetChart(cjpmp_page, cjpm, 4, null,null);
                t.AddClone(cjpmp_page);
                #endregion

                #region 竞品



                foreach (var item in param)
                {
                    var tp = new Presentation(str);
                    var temp = tp.Slides;
                    #region 格局统计


                    var page = temp[2];
                    IAutoShape text = (IAutoShape)page.Shapes[1];
                    text.TextFrame.Text = string.Format(text.TextFrame.Text, item.bamc);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add(Base_Config_Jzgj.竞争格局名称);
                    dt.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt.Columns.Add(Base_Config_Rgsj.本周到访量);
                    dt.Columns.Add(Base_Config_Rgsj.本周来电);
                    dt.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                    dt.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                    dt.Columns.Add(Base_Config_Rgsj.新开套内均价);

                    dt.Columns.Add(Base_Config_Cjba.上周_备案套数);
                    dt.Columns.Add(Base_Config_Cjba.上周_套内均价);
                    dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                    dt.Columns.Add(Base_Config_Rgsj.上周_认购套内均价);

                    dt.Columns.Add(Base_Config_Cjba.本周_备案套数);
                    dt.Columns.Add(Base_Config_Cjba.本周_套内均价);
                    dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                    dt.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);

                    dt.Columns.Add(Base_Config_Rgsj.成交套数环比);
                    dt.Columns.Add(Base_Config_Rgsj.套内均价环比);
                    dt.Columns.Add(Base_Config_Rgsj.变化原因);


                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_BX(dt, item.jpxmlb);
                        Office_Tables.SetJP_JiangBeiZuiZhiYe_JPBX_Table(page, dt, 3, null, null);
                        t.AddClone(page);
                    }
                    #endregion

                }
                #endregion
                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log(e.Message);
                return null;
            }

        }

        /// <summary>
        /// 竞品表现
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="jpxm"></param>
        /// <returns></returns>
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
                        dr1 = GET_ROW(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item);
                        dt.Rows.Add(dr1);

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
