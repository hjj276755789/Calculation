using Aspose.Slides;
using Calculation.Base;
using Calculation.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    public class plus_jp_yangguang100ximalaya : plus_jp_base
    {
        public ISlideCollection _plus_jp_yangguang100ximalaya_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);
                foreach (var item in _plus_jp_dyt_jzgj(cjbh))
                {
                    if (item != null)
                        t.AddClone(item);
                }

                #region 竞争格局    
                foreach (var item in param)
                {
                    var tp = new Presentation(str);
                    var temp = tp.Slides;
                    #region 格局图片
                   
                    #endregion


                    #region 格局统计


                    var page = temp[1];
                    IAutoShape text = (IAutoShape)page.Shapes[2];
                    text.TextFrame.Text = string.Format(text.TextFrame.Text, item.lpcs[0], item.ytcs[0]);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add(Base_Config_Jzgj.竞争格局名称);
                    dt.Columns.Add(Base_Config_Jzgj.项目名称);

                    dt.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                    dt.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                    dt.Columns.Add(Base_Config_Rgsj.本周_新开套内均价);

                    dt.Columns.Add(Base_Config_Cjba.上周_备案套数);
                    dt.Columns.Add(Base_Config_Cjba.上周_套内均价);
                    dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                    dt.Columns.Add(Base_Config_Rgsj.上周_认购套内均价);

                    dt.Columns.Add(Base_Config_Cjba.本周_备案套数);
                    dt.Columns.Add(Base_Config_Cjba.本周_套内均价);
                    dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                    dt.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);

                    dt.Columns.Add(Base_Config_Rgsj.本周_成交套数环比);
                    dt.Columns.Add(Base_Config_Rgsj.本周_套内均价环比);
                    dt.Columns.Add(Base_Config_Rgsj.本周_变化原因);
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_BX(dt, item.jpxmlb);
                        Office_Tables.SetJP_RUIAN_JPBX_Table(page, dt.AsEnumerable().OrderBy(m => m["jzgjmc"]).CopyToDataTable(), 4, null, null);
                        t.AddClone(page);
                    }
                    #endregion

                    #region 竞争格局
                    var page1 = temp[2];
                    IAutoShape text1 = (IAutoShape)page1.Shapes[1];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.lpcs[0], item.ytcs[0]);
                    System.Data.DataTable dt1 = new System.Data.DataTable();
                    dt1.Columns.Add("xm");
                    dt1.Columns.Add("yt");
                    dt1.Columns.Add("yh");
                    dt1.Columns.Add("yxdz");
                    dt1.Columns.Add("xzjtyj");
                    dt1.Columns.Add("bkfs");
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_JQDZ(dt1, item.jpxmlb);
                        Office_Tables.SetJP_RUIAN_JQHD_Table(page1, dt, 0, null, null);
                        t.AddClone(page1);
                    }
                    #endregion

                    #region 周度排名
                    ISlide sld1 = new Presentation(str).Slides[3];
                    t.AddClone(this._plus_jp_zdpm(sld1, item.bamc, new string[] { "高层" }));
                    ISlide sld2 = new Presentation(str).Slides[3];
                    t.AddClone(this._plus_jp_zdpm(sld2, item.bamc, new string[] { "洋房", "别墅" }));
                    ISlide sld3 = new Presentation(str).Slides[3];
                    t.AddClone(this._plus_jp_zdpm(sld3, item.bamc, new string[] { "商铺" }));
                    #endregion


                }
                #endregion


                #region 推广图片    
                foreach (var item in _plus_jp_dyt_tgtp(cjbh))
                {
                    if (item != null)
                        t.AddClone(item);
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
                    if (item.xfytcs != null)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {

                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);

                            var temp_basj_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_basj_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                            //本周本案认购数据
                            var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion
                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_basj_bz, temp_basj_sz, item));

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

                        dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));
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
        public DataTable GET_JPXM_JQDZ(DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            var temp = jpxm.OrderBy(m => m.id);
            foreach (var item in temp)
            {
                if (item.ytcs[0] == "别墅")
                {
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {

                        DataRow dr1 = dt.NewRow();
                        dr1[0] = item.lpcs[0] + "(" + item.xfytcs[i] + ")";//竞争楼盘名称

                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion

                        #region 上周认购数据
                        if (temp_ba_bz != null)
                        {
                            dr1["yh"] = temp_ba_bz["yh"];
                            dr1["yxdz"] = temp_ba_bz["yxdz"];
                            dr1["xzjtyj"] = temp_ba_bz["xzjtyj"];
                            dr1["bkfs"] = "-";

                        }
                        else
                        {
                            dr1["yh"] = "";
                            dr1["yxdz"] = "无";
                            dr1["xzjtyj"] = "--";
                            dr1["bkfs"] = "--";
                        }
                        #endregion
                        dt.Rows.Add(dr1);

                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    for (int i = 0; i < item.hxcs.Length; i++)
                    {
                        DataRow dr1 = dt.NewRow();
                        dr1[0] = item.lpcs[0] + "(" + item.hxcs[i] + ")";//竞争楼盘名称
                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion

                        #region 上周认购数据
                        if (temp_ba_bz != null)
                        {
                            dr1["yh"] = temp_ba_bz["yh"];
                            dr1["yxdz"] = temp_ba_bz["yxdz"];
                            dr1["xzjtyj"] = temp_ba_bz["xzjtyj"];
                            dr1["bkfs"] = "-";

                        }
                        else
                        {
                            dr1["yh"] = "";
                            dr1["yxdz"] = "无";
                            dr1["xzjtyj"] = "--";
                            dr1["bkfs"] = "--";
                        }
                        #endregion
                        dt.Rows.Add(dr1);

                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();
                    dr1[0] = item.lpcs[0] + "(" + item.ytcs[0] + ")";//竞争楼盘名称

                    #region 数据准备
                    //竞品业态

                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    #endregion

                    #region 上周认购数据
                    if (temp_ba_bz != null)
                    {
                        dr1["yh"] = temp_ba_bz["yh"];
                        dr1["yxdz"] = temp_ba_bz["yxdz"];
                        dr1["xzjtyj"] = temp_ba_bz["xzjtyj"];
                        dr1["bkfs"] = "-";

                    }
                    else
                    {
                        dr1["yh"] = "";
                        dr1["yxdz"] = "无";
                        dr1["xzjtyj"] = "--";
                        dr1["bkfs"] = "--";
                    }
                    #endregion
                    dt.Rows.Add(dr1);
                }
            }


            return dt;
        }

        public ISlide _plus_jp_zdpm(ISlide sld,string bamc, string[] yt)
        {
            #region 准备数据

            var data_zd = (from a in Cache_data_cjjl.bz.AsEnumerable()
                           where yt.Contains(a["yt"])
                           group a by new
                           {
                               lpmc = a["lpmc"],
                               zt = a["zt"]
                           } into g
                           select new
                           {
                               lpmc = g.Key.lpmc,
                               zt = g.Key.zt,
                               cjts = g.Sum(m => m["ts"].ints()),
                               cjje = g.Sum(m => m["cjje"].longs()).je_y(),
                               jzmj = g.Sum(m => m["jzmj"].doubls()).mj(),
                               tnmj = g.Sum(m => m["tnmj"].doubls()).mj(),
                           }
                           into b
                           orderby b.cjje descending
                           select b).Take(5).ToList();


            #endregion

            #region 生成页面

            if (data_zd != null & data_zd.Count > 0)
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Columns.Add("pm");
                dt.Columns.Add("lpmc");
                dt.Columns.Add("cjts");
                dt.Columns.Add("cjje");
                dt.Columns.Add("jzmj");
                dt.Columns.Add("tnmj");
                dt.Columns.Add("jmjj");
                dt.Columns.Add("tnjj");
                dt.Columns.Add("tjzj");
                dt.Columns.Add("rxyy");
                
               
                for (int i = 0; i < data_zd.Count; i++)
                {
                    DataRow dr = dt.NewRow();
                    dr["pm"] = i+1;
                    dr["lpmc"] = data_zd[i].lpmc;
                    dr["cjts"] = data_zd[i].cjts;
                    dr["cjje"] = data_zd[i].cjje.je_wy();
                    dr["jzmj"] = data_zd[i].jzmj.mj();
                    dr["tnmj"] = data_zd[i].tnmj.mj();
                    dr["jmjj"] = (data_zd[i].cjje / data_zd[i].jzmj).je_y();
                    dr["tnjj"] = (data_zd[i].cjje / data_zd[i].tnmj).je_y();
                    dr["tjzj"] = (data_zd[i].cjje / data_zd[i].cjts).je_wy();
                    dr["rxyy"] = "自填";
                    dt.Rows.Add(dr);
                }

                IAutoShape text1 = (IAutoShape)sld.Shapes[1];
                text1.TextFrame.Text = string.Format(text1.TextFrame.Text, bamc, string.Join(",", yt));
                Office_Tables.SetJP_YG100XMLY_ZDYTPM_Table(sld, dt, 2, null, null);

                IAutoShape text2 = (IAutoShape)sld.Shapes[3];
                text2.TextFrame.Text = string.Format(text2.TextFrame.Text, string.Join(",", yt), Base_date.GET_ZCMC(Base_date.bn, Base_date.bz));

                return sld;
            }
            #endregion
            return null;
        }
    }
}
