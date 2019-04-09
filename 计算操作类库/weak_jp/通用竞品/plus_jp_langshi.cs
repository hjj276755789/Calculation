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
    public class plus_jp_langshi:plus_jp_base
    {

        /// <summary>
        /// 差别十分巨大，无法重用
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_langshi_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);

                foreach (var item in jbzzs(str))
                {
                    t.AddClone(item);
                }

                foreach (var item in param)
                {
                    var tp = new Presentation(str);
                    var temp = tp.Slides;

                    #region 竞品分布
                    var page1 = temp[9];
                    #endregion
                    t.AddClone(page1);

                    #region 格局统计
                    var page2 = temp[10];
                    IAutoShape text1 = (IAutoShape)page2.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc);
                    DataTable dt = new DataTable();
                    dt.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt.Columns.Add(Base_Config_Jzgj.业态);

                    dt.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                    dt.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                    dt.Columns.Add("bz"+Base_Config_Rgsj.本周_认购建面均价);

                    dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                    dt.Columns.Add(Base_Config_Rgsj.上周_认购套内均价);
                    dt.Columns.Add(Base_Config_Rgsj.上周_认购建面均价);

                    dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                    dt.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);
                    dt.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);

                    dt.Columns.Add("heji");
                  
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_BX(dt, item.jpxmlb);
                        Office_Tables.SetJP_Langshi_JPBX_Table(page2, dt, 5, null, null);  
                    }
                    #endregion
                    t.AddClone(page2);

                    #region 近期动作
                    var page3 = temp[11];
                    DataTable dt1 = new DataTable();
                    dt1.Columns.Add(Base_Config_Jzgj.竞争格局名称);
                    dt1.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt1.Columns.Add(Base_Config_Jzgj.业态);
                    dt1.Columns.Add(Base_Config_Rgsj.本周_优惠);
                    dt1.Columns.Add(Base_Config_Rgsj.本周_活动);
                    dt1.Columns.Add(Base_Config_Rgsj.本周_营销动作);
                    dt1.Columns.Add("bkfsjcxqk");
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt1 = GET_JPXM_BX(dt1, item.jpxmlb);
                        Office_Tables.SetTable(page3, dt1, 2, null, null);
                    }

                    #endregion
                    t.AddClone(page3);
                }
                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log("插件：" + cjbh + "生成报错*****" + e.Message);
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
                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, item));

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

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion
                        dt.Rows.Add(GET_ROW(item.hxcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, item));
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, item));
                }


            }


            return dt;
        }

        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,DataRow temp_ba_bz,DataRow temp_ba_sz,JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                if (dt.Columns[j].ColumnName == "heji")
                {

                    dr1[dt.Columns[j].ColumnName] = (temp_ba_sz != null ? temp_ba_sz[Base_Config_Rgsj.上周_新开套数._ConfigRgsjMc()].ints() : 0) + (temp_ba_bz != null ? temp_ba_bz[Base_Config_Rgsj.本周_新开套数._ConfigRgsjMc()].ints() : 0);
                }
                else if (dt.Columns[j].ColumnName == "bz" + Base_Config_Rgsj.本周_认购建面均价)
                {
                    dr1[dt.Columns[j].ColumnName] = temp_ba_bz!=null? temp_ba_bz[Base_Config_Rgsj.本周_认购建面均价._ConfigRgsjMc()]:0;
                }
                else if (Base_Config_Rgsj._认购数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Rgsj.上周_新开套数:
                        case Base_Config_Rgsj.上周_新开销售套数:
                        case Base_Config_Rgsj.上周_认购套数:
                        case Base_Config_Rgsj.上周_认购套内体量:
                        case Base_Config_Rgsj.上周_认购套内均价:
                        case Base_Config_Rgsj.上周_认购建面体量:
                        case Base_Config_Rgsj.上周_认购建面均价:
                        case Base_Config_Rgsj.上周_认购金额:
                            {
                                if (temp_ba_sz != null)
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_sz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                }
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                        case Base_Config_Rgsj.本周_新开套数:
                        case Base_Config_Rgsj.本周_新开销售套数:
                        case Base_Config_Rgsj.本周_认购套数:
                        case Base_Config_Rgsj.本周_认购套内体量:
                        case Base_Config_Rgsj.本周_认购套内均价:
                        case Base_Config_Rgsj.本周_认购建面体量:
                        case Base_Config_Rgsj.本周_认购建面均价:
                        case Base_Config_Rgsj.本周_认购金额:
                            {
                                if (temp_ba_bz != null)
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                }
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                        default:
                            {
                                if (temp_ba_bz != null)
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName];
                                }
                                else
                                    dr1[dt.Columns[j].ColumnName] = "";
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
                                dr1[dt.Columns[j].ColumnName] = string.Join(",", item.ztcs);
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
            }

            return dr1;
        }

        public ISlideCollection jbzzs(string str)
        {
            var p = new Presentation();
            var t = p.Slides;
            t.RemoveAt(0);
            var tp = new Presentation(str);
            var temp = tp.Slides;



            string[] zt = { "蔡家", "礼嘉", "悦来", "中央公园" };
            #region  P1
            var page1 = temp[0];
            var dt1_1 = from a in Cache_data_xzys.jbz.AsEnumerable()
                        where zt.Contains(a["zt"]) && a["tyyt"].ToString() == "高层"
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            xzgyl = s.Sum(m => m["jzmj"].doubls())
                        };
            var dt1_2 = from a in Cache_data_cjjl.jbz.AsEnumerable()
                        where zt.Contains(a["zt"]) && a["yt"].ToString() == "高层"
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            cjje = s.Sum(m => m["cjje"].longs()),
                            jzmj = s.Sum(m => m["jzmj"].doubls())
                        };
            DataTable dt1 = new DataTable();
            dt1.Columns.Add("周次");
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 7), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 6), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 5), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 4), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz ), typeof(double));
            DataRow dr1 = dt1.NewRow();
            DataRow dr2 = dt1.NewRow();
            DataRow dr3 = dt1.NewRow();
            dr1[0] = "供应体量";
            dr2[0] = "成交体量";
            dr3[0] = "建面均价";
            
            for (int i = 0; i < 8; i++)
            {
                var xzys = dt1_1.FirstOrDefault(m => m.zc.ints() == (Base_date.bz - (7 - i)));
                var cjba = dt1_2.FirstOrDefault(m => m.zc.ints() == (Base_date.bz - (7 - i)));
                dr1[i + 1] = (xzys != null) ? xzys.xzgyl.mj_wf() : 0;
                dr2[i + 1] = cjba != null ? cjba.jzmj.mj_wf() : 0;
                dr3[i + 1] = cjba != null ? (cjba.cjje/cjba.jzmj).je_y() : 0;
                if(i==7)
                {
                    IAutoShape text1 = (IAutoShape)page1.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, (xzys != null) ? xzys.xzgyl.mj_wf():0, cjba != null ? cjba.jzmj.mj_wf() : 0, cjba != null ? (cjba.cjje / cjba.jzmj).je_y() : 0);
                }
            }
            dt1.Rows.Add(dr1);
            dt1.Rows.Add(dr2);
            dt1.Rows.Add(dr3);
            Office_Charts.Chart_jp_langshi_chart1(page1, dt1, 4);
            t.AddClone(page1);
            #endregion
            #region  P2
            var page2 = temp[1];
            var dt2_1 = from a in Cache_data_xzys.jbz.AsEnumerable()
                        where zt.Contains(a["zt"]) && a["tyyt"].ToString() == "洋房"
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            xzgyl = s.Sum(m => m["jzmj"].doubls())
                        };
            var dt2_2 = from a in Cache_data_cjjl.jbz.AsEnumerable()
                        where zt.Contains(a["zt"]) && a["yt"].ToString() == "洋房"
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            cjje = s.Sum(m => m["cjje"].longs()),
                            jzmj = s.Sum(m => m["jzmj"].doubls())
                        };
            DataTable dt2 = new DataTable();
            dt2.Columns.Add("周次");
            dt2.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 7), typeof(double));
            dt2.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 6), typeof(double));
            dt2.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 5), typeof(double));
            dt2.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 4), typeof(double));
            dt2.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3), typeof(double));
            dt2.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2), typeof(double));
            dt2.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1), typeof(double));
            dt2.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz), typeof(double));
            DataRow dr2_1 = dt2.NewRow();
            DataRow dr2_2 = dt2.NewRow();
            DataRow dr2_3 = dt2.NewRow();
            dr2_1[0] = "供应体量";
            dr2_2[0] = "成交体量";
            dr2_3[0] = "建面均价";

            for (int i = 0; i < 8; i++)
            {
                var xzys = dt2_1.FirstOrDefault(m => m.zc.ints() == (Base_date.bz - (7 - i)));
                var cjba = dt2_2.FirstOrDefault(m => m.zc.ints() == (Base_date.bz - (7 - i)));
                dr2_1[i + 1] = (xzys != null) ? xzys.xzgyl.mj_wf() : 0;
                dr2_2[i + 1] = cjba != null ? cjba.jzmj.mj_wf() : 0;
                dr2_3[i + 1] = cjba != null ? (cjba.cjje / cjba.jzmj).je_y() : 0;
                if (i == 7)
                {
                    IAutoShape text1 = (IAutoShape)page2.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, (xzys != null) ? xzys.xzgyl.mj_wf() : 0, cjba != null ? cjba.jzmj.mj_wf() : 0, cjba != null ? (cjba.cjje / cjba.jzmj).je_y() : 0);
                }
            } 
            dt2.Rows.Add(dr2_1);
            dt2.Rows.Add(dr2_2);
            dt2.Rows.Add(dr2_3);
            Office_Charts.Chart_jp_langshi_chart1(page2, dt2, 4);
            t.AddClone(page2);
            #endregion
            #region  P3
            var page3 = temp[2];
            var dt3_1 = from a in Cache_data_xzys.jbz.AsEnumerable()
                        where zt.Contains(a["zt"]) && a["tyyt"].ToString() == "别墅"
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            xzgyl = s.Sum(m => m["jzmj"].doubls())
                        };
            var dt3_2 = from a in Cache_data_cjjl.jbz.AsEnumerable()
                        where zt.Contains(a["zt"]) && a["yt"].ToString() == "别墅"
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            cjje = s.Sum(m => m["cjje"].longs()),
                            jzmj = s.Sum(m => m["jzmj"].doubls())
                        };
            DataTable dt3 = new DataTable();
            dt3.Columns.Add("周次");
            dt3.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 7), typeof(double));
            dt3.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 6), typeof(double));
            dt3.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 5), typeof(double));
            dt3.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 4), typeof(double));
            dt3.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3), typeof(double));
            dt3.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2), typeof(double));
            dt3.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1), typeof(double));
            dt3.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz), typeof(double));
            DataRow dr3_1 = dt3.NewRow();
            DataRow dr3_2 = dt3.NewRow();
            DataRow dr3_3 = dt3.NewRow();
            dr3_1[0] = "供应体量";
            dr3_2[0] = "成交体量";
            dr3_3[0] = "建面均价";

            for (int i = 0; i < 8; i++)
            {
                var xzys = dt3_1.FirstOrDefault(m => m.zc.ints() == (Base_date.bz - (7 - i)));
                var cjba = dt3_2.FirstOrDefault(m => m.zc.ints() == (Base_date.bz - (7 - i)));
                dr3_1[i + 1] = xzys != null ? xzys.xzgyl.mj_wf() : 0;
                dr3_2[i + 1] = cjba != null ? cjba.jzmj.mj_wf() : 0;
                dr3_3[i + 1] = cjba != null ? (cjba.cjje / cjba.jzmj).je_y() : 0;
                if (i == 7)
                {
                    IAutoShape text1 = (IAutoShape)page3.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, (xzys != null) ? xzys.xzgyl.mj_wf() : 0, cjba != null ? cjba.jzmj.mj_wf() : 0, cjba != null ? (cjba.cjje / cjba.jzmj).je_y() : 0);
                }
            }
            dt3.Rows.Add(dr3_1);
            dt3.Rows.Add(dr3_2);
            dt3.Rows.Add(dr3_3);
            Office_Charts.Chart_jp_langshi_chart1(page3, dt3, 4);
            t.AddClone(page3);
            #endregion

            #region P4
            var page4 = temp[3];
            var dt4_1 = from a in Cache_data_cjjl.bz.AsEnumerable()
                        where zt.Contains(a["zt"]) && (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                        group a by new { lpmc = a["lpmc"], zt = a["zt"] } into s
                        select new
                        {
                            lpmc = s.Key.lpmc,
                            zt = s.Key.zt,
                            ts = s.Sum(m => m["ts"].ints()),
                            cjje = s.Sum(m => m["cjje"].longs()),
                            jzmj = s.Sum(m => m["jzmj"].doubls()),
                            tnmj = s.Sum(m => m["tnmj"].doubls())
                        };
            DataTable dt4 = new DataTable();
            dt4.Columns.Add("pm");
            dt4.Columns.Add("lpmc1");
            dt4.Columns.Add("zt1");
            dt4.Columns.Add("ts");

            dt4.Columns.Add("lpmc2");
            dt4.Columns.Add("zt2");
            dt4.Columns.Add("jzmj");

            dt4.Columns.Add("lpmc3");
            dt4.Columns.Add("zt3");
            dt4.Columns.Add("cjje");
            var ts = dt4_1.OrderByDescending(m => m.ts).Take(5).ToList();
            var cjmj = dt4_1.OrderByDescending(m => m.jzmj).Take(5).ToList();
            var cjje = dt4_1.OrderByDescending(m => m.cjje).Take(5).ToList();

            for (int i = 0; i < 5; i++)
            {
                DataRow dr = dt4.NewRow();
                dr["pm"] = i+1;
                if(ts!=null&&ts.Count>i)
                {
                    dr["lpmc1"] = ts[i].lpmc;
                    dr["zt1"] = ts[i].zt;
                    dr["ts"] = ts[i].ts;
                }
                if (cjmj != null && cjmj.Count > i)
                {
                    dr["lpmc2"] = cjmj[i].lpmc;
                    dr["zt2"] = cjmj[i].zt;
                    dr["jzmj"] = cjmj[i].jzmj.mj();
                }
                if (cjje != null && cjje.Count > i)
                {
                    dr["lpmc3"] = cjje[i].lpmc;
                    dr["zt3"] = cjje[i].zt;
                    dr["cjje"] = cjje[i].cjje.je_wy();
                }
                dt4.Rows.Add(dr);
                if(i==0)
                {
                    IAutoShape text4 = (IAutoShape)page4.Shapes[2];
                    text4.TextFrame.Text = string.Format(text4.TextFrame.Text, ts[i].lpmc, cjmj[i].lpmc, cjje[i].lpmc);
                }

            }
            Office_Tables.SetChart(page4, dt4, 5, null, null);
            t.AddClone(page4);
            #endregion
            #region P5
            var page5 = temp[4];
            var dt5_1 = from a in Cache_data_cjjl.bz.AsEnumerable()
                                    where zt.Contains(a["zt"]) && ( a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层")
                                    group a by new { lpmc = a["lpmc"], zt = a["zt"] } into s
                                    select new
                                    {
                                        lpmc = s.Key.lpmc,
                                        zt = s.Key.zt,
                                        ts = s.Sum(m => m["ts"].ints()),
                                        cjje = s.Sum(m => m["cjje"].longs()),
                                        jzmj = s.Sum(m => m["jzmj"].doubls()),
                                        tnmj = s.Sum(m => m["tnmj"].doubls())
                                    };
            DataTable dt5 = new DataTable();
            dt5.Columns.Add("pm");
            dt5.Columns.Add("lpmc");
            dt5.Columns.Add("zt");
            dt5.Columns.Add("ts");
            dt5.Columns.Add("jzmj");
            dt5.Columns.Add("cjje");
            dt5.Columns.Add("jmjj");
            dt5.Columns.Add("tnjj");
            var dt5_1_1 = dt5_1.OrderByDescending(m => m.ts).Take(5).ToList();
            for (int i = 0; i < 5; i++)
            {
                if (dt5_1_1 != null && dt5_1_1.Count > i)
                {
                    DataRow dr = dt5.NewRow();
                    dr["pm"] = i + 1;
                    dr["lpmc"] = dt5_1_1[i].lpmc;
                    dr["zt"] = dt5_1_1[i].zt;
                    dr["ts"] = dt5_1_1[i].ts;
                    dr["jzmj"] = dt5_1_1[i].jzmj.mj();
                    dr["cjje"] = dt5_1_1[i].cjje.je_wy();
                    dr["jmjj"] = dt5_1_1[i].jzmj != 0 ? (dt5_1_1[i].cjje / dt5_1_1[i].jzmj).je_y() : 0;
                    dr["tnjj"] = dt5_1_1[i].tnmj != 0 ? (dt5_1_1[i].cjje / dt5_1_1[i].tnmj).je_y() : 0;
                    dt5.Rows.Add(dr);
                    if (i == 0)
                    {
                        IAutoShape text5 = (IAutoShape)page5.Shapes[2];
                        text5.TextFrame.Text = string.Format(text5.TextFrame.Text, dt5_1_1[i].lpmc);
                    }
                }
                else
                    break;
            }
            Office_Tables.SetChart(page5, dt5, 5, null, null);
            t.AddClone(page5);
            #endregion

            #region P6
            var page6 = temp[5];
            var dt6_1 = from a in Cache_data_cjjl.bz.AsEnumerable()
                        where zt.Contains(a["zt"]) && (a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                        group a by new { lpmc = a["lpmc"], zt = a["zt"] } into s
                        select new
                        {
                            lpmc = s.Key.lpmc,
                            zt = s.Key.zt,
                            ts = s.Sum(m => m["ts"].ints()),
                            cjje = s.Sum(m => m["cjje"].longs()),
                            jzmj = s.Sum(m => m["jzmj"].doubls()),
                            tnmj = s.Sum(m => m["tnmj"].doubls())
                        };
            DataTable dt6 = new DataTable();
            dt6.Columns.Add("pm");
            dt6.Columns.Add("lpmc");
            dt6.Columns.Add("zt");
            dt6.Columns.Add("ts");
            dt6.Columns.Add("jzmj");
            dt6.Columns.Add("cjje");
            dt6.Columns.Add("jmjj");
            dt6.Columns.Add("tnjj");
            var dt6_1_1 = dt6_1.OrderByDescending(m => m.ts).Take(5).ToList();
            for (int i = 0; i < 5; i++)
            {
                if (dt6_1_1 != null && dt6_1_1.Count > i)
                {
                    DataRow dr = dt6.NewRow();
                    dr["pm"] = i + 1;
                    dr["lpmc"] = dt6_1_1[i].lpmc;
                    dr["zt"] =   dt6_1_1[i].zt;
                    dr["ts"] =   dt6_1_1[i].ts;
                    dr["jzmj"] = dt6_1_1[i].jzmj.mj();
                    dr["cjje"] = dt6_1_1[i].cjje.je_wy();
                    dr["jmjj"] = dt6_1_1[i].jzmj != 0 ? (dt6_1_1[i].cjje / dt6_1_1[i].jzmj).je_y() : 0;
                    dr["tnjj"] = dt6_1_1[i].tnmj != 0 ? (dt6_1_1[i].cjje / dt6_1_1[i].tnmj).je_y() : 0;
                    dt6.Rows.Add(dr);
                    if (i == 0)
                    {
                        IAutoShape text6 = (IAutoShape)page6.Shapes[5];
                        text6.TextFrame.Text = string.Format(text6.TextFrame.Text, dt6_1_1[i].lpmc);
                    }
                }
                else
                    break;
            }
            Office_Tables.SetChart(page6, dt6, 4, null, null);
            t.AddClone(page6);
            #endregion
            #region P7
            var page7 = temp[6];
            var dt7_1 = from a in Cache_data_cjjl.bz.AsEnumerable()
                        where zt.Contains(a["zt"]) && (a["xfyt"].ToString() == "叠加别墅")
                        group a by new { lpmc = a["lpmc"], zt = a["zt"] } into s
                        select new
                        {
                            lpmc = s.Key.lpmc,
                            zt = s.Key.zt,
                            ts = s.Sum(m => m["ts"].ints()),
                            cjje = s.Sum(m => m["cjje"].longs()),
                            jzmj = s.Sum(m => m["jzmj"].doubls()),
                            tnmj = s.Sum(m => m["tnmj"].doubls())
                        };
            DataTable dt7 = new DataTable();
            dt7.Columns.Add("pm");
            dt7.Columns.Add("lpmc");
            dt7.Columns.Add("zt");
            dt7.Columns.Add("ts");
            dt7.Columns.Add("jzmj");
            dt7.Columns.Add("cjje");
            dt7.Columns.Add("tjzj");
            dt7.Columns.Add("jmjj");
            dt7.Columns.Add("tnjj");
            var dt7_1_1 = dt7_1.OrderByDescending(m => m.ts).Take(5).ToList();
            for (int i = 0; i < 5; i++)
            {
                if (dt7_1_1 != null && dt7_1_1.Count > i)
                {
                    DataRow dr = dt7.NewRow();
                    dr["pm"] = i + 1;
                    dr["lpmc"] = dt7_1_1[i].lpmc;
                    dr["zt"] =   dt7_1_1[i].zt;
                    dr["ts"] =   dt7_1_1[i].ts;
                    dr["jzmj"] = dt7_1_1[i].jzmj.mj();
                    dr["cjje"] = dt7_1_1[i].cjje.je_wy();
                    dr["tjzj"] = (dt7_1_1[i].cjje / dt7_1_1[i].ts).je_wy();
                    dr["jmjj"] = dt7_1_1[i].jzmj != 0 ? (dt7_1_1[i].cjje / dt7_1_1[i].jzmj).je_y() : 0;
                    dr["tnjj"] = dt7_1_1[i].tnmj != 0 ? (dt7_1_1[i].cjje / dt7_1_1[i].tnmj).je_y() : 0;
                    dt7.Rows.Add(dr);
                    if (i == 0)
                    {
                        IAutoShape text7 = (IAutoShape)page7.Shapes[5];
                        text7.TextFrame.Text = string.Format(text7.TextFrame.Text, dt7_1_1[i].lpmc);
                    }
                }
                else
                    break;
            }
            Office_Tables.SetChart(page7, dt7,4, null, null);
            t.AddClone(page7);
            #endregion
            #region P8
            var page8 = temp[7];
            var dt8_1 = from a in Cache_data_cjjl.bz.AsEnumerable()
                        where zt.Contains(a["zt"]) && (a["xfyt"].ToString() == "联排别墅")
                        group a by new { lpmc = a["lpmc"], zt = a["zt"] } into s
                        select new
                        {
                            lpmc = s.Key.lpmc,
                            zt = s.Key.zt,
                            ts = s.Sum(m => m["ts"].ints()),
                            cjje = s.Sum(m => m["cjje"].longs()),
                            jzmj = s.Sum(m => m["jzmj"].doubls()),
                            tnmj = s.Sum(m => m["tnmj"].doubls())
                        };
            DataTable dt8 = new DataTable();
            dt8.Columns.Add("pm");
            dt8.Columns.Add("lpmc");
            dt8.Columns.Add("zt");
            dt8.Columns.Add("ts");
            dt8.Columns.Add("jzmj");
            dt8.Columns.Add("cjje");
            dt8.Columns.Add("tjzj");
            dt8.Columns.Add("jmjj");
            dt8.Columns.Add("tnjj");
            var dt8_1_1 = dt8_1.OrderByDescending(m => m.ts).Take(5).ToList();
            for (int i = 0; i < 5; i++)
            {
                if (dt8_1_1 != null && dt8_1_1.Count > i)
                {
                    DataRow dr = dt8.NewRow();
                    dr["pm"] = i + 1;
                    dr["lpmc"] = dt8_1_1[i].lpmc;
                    dr["zt"] =   dt8_1_1[i].zt;
                    dr["ts"] =   dt8_1_1[i].ts;
                    dr["jzmj"] = dt8_1_1[i].jzmj.mj();
                    dr["cjje"] = dt8_1_1[i].cjje.je_wy();
                    dr["tjzj"] = (dt8_1_1[i].cjje / dt8_1_1[i].ts).je_wy();
                    dr["jmjj"] = dt8_1_1[i].jzmj != 0 ? (dt8_1_1[i].cjje / dt8_1_1[i].jzmj).je_y() : 0;
                    dr["tnjj"] = dt8_1_1[i].tnmj != 0 ? (dt8_1_1[i].cjje / dt8_1_1[i].tnmj).je_y() : 0;
                    dt8.Rows.Add(dr);
                    if (i == 0)
                    {
                        IAutoShape text8 = (IAutoShape)page8.Shapes[5];
                        text8.TextFrame.Text = string.Format(text8.TextFrame.Text, dt8_1_1[i].lpmc);
                    }
                }
                else
                    break;
            }
            Office_Tables.SetChart(page8, dt8, 4, null, null);
            t.AddClone(page8);
            #endregion
            #region P9
            var page9 = temp[8];
            var dt9_1_1 = from a in Cache_data_xzys.bz.AsEnumerable()
                          where a["zt"].ToString() == "蔡家" && (a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "洋房")
                          group a by new { yt = a["tyyt"] } into s
                          select new
                          {
                              yt = s.Key.yt,
                              jzmj = s.Sum(m => m["jzmj"].doubls())
                          };
            var dt9_1_2 = from a in Cache_data_xzys.bz.AsEnumerable()
                        where a["zt"].ToString()=="蔡家" && (a["wylx"].ToString() == "联排别墅" || a["wylx"].ToString() == "叠加别墅")
                        group a by new { yt=a["wylx"]} into s
                        select new
                        {
                            yt = s.Key.yt,
                            jzmj = s.Sum(m => m["jzmj"].doubls())
                        };
                            

          
            var dt9_2_1 = from a in Cache_data_cjjl.bz.AsEnumerable()
                        where a["zt"].ToString() == "蔡家" && (a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                        group a by new { yt = a["yt"] } into s
                        select new
                        {
                            yt=s.Key.yt,
                            ts = s.Sum(m => m["ts"].ints()),
                            cjje = s.Sum(m => m["cjje"].longs()),
                            jzmj = s.Sum(m => m["jzmj"].doubls()),
                            tnmj = s.Sum(m => m["tnmj"].doubls())
                        };
            var dt9_2_2 = from a in Cache_data_cjjl.bz.AsEnumerable()
                          where a["zt"].ToString() == "蔡家" && (a["xfyt"].ToString() == "联排别墅" || a["xfyt"].ToString() == "叠加别墅")
                          group a by new { yt = a["xfyt"] } into s
                          select new
                          {
                              yt = s.Key.yt,
                              ts = s.Sum(m => m["ts"].ints()),
                              cjje = s.Sum(m => m["cjje"].longs()),
                              jzmj = s.Sum(m => m["jzmj"].doubls()),
                              tnmj = s.Sum(m => m["tnmj"].doubls())
                          };
            var dt9_3_1 = from a in Cache_data_cjjl.sz.AsEnumerable()
                          where a["zt"].ToString() == "蔡家" && (a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                          group a by new { yt = a["yt"] } into s
                          select new
                          {
                              yt = s.Key.yt,
                              ts = s.Sum(m => m["ts"].ints()),
                              cjje = s.Sum(m => m["cjje"].longs()),
                              jzmj = s.Sum(m => m["jzmj"].doubls()),
                              tnmj = s.Sum(m => m["tnmj"].doubls())
                          };
            var dt9_3_2 = from a in Cache_data_cjjl.sz.AsEnumerable()
                          where a["zt"].ToString() == "蔡家" && (a["xfyt"].ToString() == "联排别墅" || a["xfyt"].ToString() == "叠加别墅")
                          group a by new { yt = a["xfyt"] } into s
                          select new
                          {
                              yt=s.Key.yt,
                              ts = s.Sum(m => m["ts"].ints()),
                              cjje = s.Sum(m => m["cjje"].longs()),
                              jzmj = s.Sum(m => m["jzmj"].doubls()),
                              tnmj = s.Sum(m => m["tnmj"].doubls())
                          };

            DataTable dt9 = new DataTable();
            dt9.Columns.Add("yt");
            dt9.Columns.Add("bzgyl");
            dt9.Columns.Add("bzqhl");
            dt9.Columns.Add("jmjj");
            dt9.Columns.Add("hbjgzf");
            DataRow dr9_1 = dt9.NewRow();
            var gc_1 = dt9_1_1.FirstOrDefault(m => m.yt.ToString() == "高层");
            var gc_2 = dt9_2_1.FirstOrDefault(m => m.yt.ToString() == "高层");
            var gc_3 = dt9_3_1.FirstOrDefault(m => m.yt.ToString() == "高层");
            dr9_1["yt"] = "高层";
            dr9_1["bzgyl"] =   gc_1 != null ?  gc_1.jzmj.mj_wf() : 0;
            dr9_1["bzqhl"] =   gc_2 != null ?  gc_2.jzmj.mj_wf() : 0;
            dr9_1["jmjj"] =    gc_2 != null ? (gc_2.cjje / gc_2.jzmj).je_y() : 0;
            dr9_1["hbjgzf"] = (gc_2 != null && gc_3 != null) ? ((gc_2.cjje / gc_2.jzmj - gc_3.cjje / gc_3.jzmj) / (gc_3.cjje / gc_3.jzmj)).ss_bfb() : "";
            dt9.Rows.Add(dr9_1);
            DataRow dr9_2 = dt9.NewRow();
            var yf_1 = dt9_1_1.FirstOrDefault(m => m.yt.ToString() == "洋房");
            var yf_2 = dt9_2_1.FirstOrDefault(m => m.yt.ToString() == "洋房");
            var yf_3 = dt9_3_1.FirstOrDefault(m => m.yt.ToString() == "洋房");
            dr9_2["yt"] = "洋房";
            dr9_2["bzgyl"] = yf_1 != null ? yf_1.jzmj.mj_wf() : 0;
            dr9_2["bzqhl"] = yf_2 != null ? yf_2.jzmj.mj_wf() : 0;
            dr9_2["jmjj"] = yf_2 != null ? (yf_2.cjje / yf_2.jzmj).je_y() : 0 ;
            dr9_2["hbjgzf"] = (yf_2!=null&&yf_3!=null)?((yf_2.cjje / yf_2.jzmj - yf_3.cjje / yf_3.jzmj) / (yf_3.cjje / yf_3.jzmj)).ss_bfb():"";
            dt9.Rows.Add(dr9_2);
            DataRow dr9_3 = dt9.NewRow();
            var lpbs_1 = dt9_1_2.FirstOrDefault(m => m.yt.ToString() == "联排别墅");
            var lpbs_2 = dt9_2_2.FirstOrDefault(m => m.yt.ToString() == "联排别墅");
            var lpbs_3 = dt9_3_2.FirstOrDefault(m => m.yt.ToString() == "联排别墅");
            dr9_3["yt"] = "联排别墅";
            dr9_3["bzgyl"] = lpbs_1 != null ? lpbs_1.jzmj.mj_wf() : 0;
            dr9_3["bzqhl"] = lpbs_2 != null ? lpbs_2.jzmj.mj_wf() : 0;
            dr9_3["jmjj"] = lpbs_2 != null ? (lpbs_2.cjje / lpbs_2.tnmj).je_y() : 0;
            dr9_3["hbjgzf"] = (lpbs_2 != null && lpbs_3 != null) ? ((lpbs_2.cjje / lpbs_2.tnmj - lpbs_3.cjje / lpbs_3.tnmj) / (lpbs_3.cjje / lpbs_3.tnmj)).ss_bfb() : "";
            dt9.Rows.Add(dr9_3);

            DataRow dr9_4 = dt9.NewRow();
            var djbs_1 = dt9_1_2.FirstOrDefault(m => m.yt.ToString() == "叠加别墅");
            var djbs_2 = dt9_2_2.FirstOrDefault(m => m.yt.ToString() == "叠加别墅");
            var djbs_3 = dt9_3_2.FirstOrDefault(m => m.yt.ToString() == "叠加别墅");
            dr9_4["yt"] = "叠加别墅";
            dr9_4["bzgyl"] = djbs_1 != null ? djbs_1.jzmj.mj_wf() : 0;
            dr9_4["bzqhl"] = djbs_2 != null ? djbs_2.jzmj.mj_wf() : 0;
            dr9_4["jmjj"] = djbs_2 != null ? (djbs_2.cjje / djbs_2.tnmj).je_y() : 0;
            dr9_4["hbjgzf"] = (djbs_2 != null && djbs_3 != null) ? ((djbs_2.cjje / djbs_2.tnmj - djbs_3.cjje / djbs_3.tnmj) / (djbs_3.cjje / djbs_3.tnmj)).ss_bfb() : "";

            dt9.Rows.Add(dr9_4);

            DataRow dr9_5 = dt9.NewRow();
            string[] yt = { "洋房","洋楼" ,"高层","小高层", "别墅"};
            var hj_1 = Cache_data_xzys.bz.AsEnumerable().Where(m => yt.Contains(m["tyyt"]) && m["zt"].ToString() == "蔡家");
            var hj_2 = Cache_data_cjjl.bz.AsEnumerable().Where(m => yt.Contains(m["yt"]) && m["zt"].ToString() == "蔡家");
            var hj_3 = Cache_data_cjjl.sz.AsEnumerable().Where(m => yt.Contains(m["yt"]) && m["zt"].ToString() == "蔡家");
            dr9_5["yt"] = "合计";
            dr9_5["bzgyl"] = hj_1!= null? hj_1.Sum(m => m["jzmj"].doubls()).mj_wf():0;
            dr9_5["bzqhl"] = hj_2 != null? hj_2.Sum(m => m["jzmj"].doubls()).mj_wf():0;
            dr9_5["jmjj"] = hj_2 != null ? (hj_2.Sum(m => m["cjje"].doubls()) / hj_2.Sum(m => m["jzmj"].doubls())).je_y() : 0;
            dr9_5["hbjgzf"] = (hj_2 != null && hj_3 != null) ? ((hj_2.Sum(m => m["cjje"].doubls()) / hj_2.Sum(m => m["jzmj"].doubls()) - hj_3.Sum(m => m["cjje"].doubls()) / hj_3.Sum(m => m["jzmj"].doubls())) / (hj_3.Sum(m => m["cjje"].doubls()) / hj_3.Sum(m => m["jzmj"].doubls()))).ss_bfb() : "";

            dt9.Rows.Add(dr9_5);

            Office_Tables.SetChart(page9, dt9, 5, null, null);
            t.AddClone(page9);
            #endregion
            return t;
        }
    }
}
