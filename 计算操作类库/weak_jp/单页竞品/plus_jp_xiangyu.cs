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
    public class plus_jp_xiangyu : plus_jp_base
    {
        private static DataTable by;

        public class Base_Config_Cjba_BY
        {

            public const string 本月_备案套数 = "by_ts";
            public const string 本月_成交金额 = "by_cjje";
            public const string 本月_建筑面积 = "by_jzmj";
            public const string 本月_套内面积 = "by_tnmj";
            public const string 本月_建面均价 = "by_jmjj";
            public const string 本月_套内均价 = "by_tnjj";
            public const string 本月_套均总价 = "by_tjzj";
            public static string[] _备案数据 = { "by_ts", "by_cjje", "by_jzmj", "by_tnmj", "by_jmjj", "by_tnjj", "by_tjzj", };
        }
        public class Base_Config_Cjba_bh
        {

            public const string 本周_套数变化 = "bz_tsbh";
            public const string 本周_价格变化 = "bh_jgbh";
           
            public static string[] _备案数据 = { "bz_tsbh", "bh_jgbh" };
        }

        public plus_jp_xiangyu()
        {
            Base_date.init_yb(Base_date.bn, Base_date.GET_Y_FROM_Z(Base_date.bn, Base_date.bz));
            by = ZB_Data_CJBA_DataProvider.GET_ZB(Base_date.by_First, Base_date.bz_Last);
        }
        /// <summary>
        /// 象屿市场简报竞品
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_xiangyu_1(string str, int cjbh)
        {
            try
            {

                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);
                foreach (var item in param)
                {


                    #region 竞品备案数据


                    if (item.qtcs == "竞品备案数据")
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page1 = temp[0];

                        IAutoShape text1 = (IAutoShape)page1.Shapes[1];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc ,item.ytcs[0]);

                        DataTable dt_jpbasj = new DataTable();
                        dt_jpbasj.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt_jpbasj.Columns.Add(Base_Config_Jzgj.业态);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_建筑面积);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_建面均价);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_套均总价);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba_bh.本周_套数变化);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_备案套数环比);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba_bh.本周_价格变化);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_建面均价环比);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.上周_建筑面积);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.上周_建面均价);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.上周_套均总价);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba_BY.本月_备案套数);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba_BY.本月_建筑面积);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba_BY.本月_建面均价);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba_BY.本月_套均总价);

                        dt_jpbasj.Columns.Add("本年_累计套数");
                        dt_jpbasj.Columns.Add("本年_建筑面积");
                        dt_jpbasj.Columns.Add("本年_建面均价");
                        dt_jpbasj.Columns.Add("本年_套均总价");
                        //获取本案数据
                        dt_jpbasj = GET_JPBA_BX(dt_jpbasj, item);
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            //获取竞品项目数据
                            dt_jpbasj = GET_JPXM_BX(dt_jpbasj, item.jpxmlb);
                            Office_Tables.SetJP_XIANGYU_JINGPINGBEIAN_Table(page1, dt_jpbasj, 2, null, null);
                            t.AddClone(page1);
                        }
                       

                    }
                    #endregion




                    else
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page2 = temp[1];
                        foreach (var jpfb in _plus_jp_dyt_jzgj_taonei(page2, item))
                        {
                            t.AddClone(jpfb);
                        }

                        #region 竞品市场表现

                        var page3 = temp[2];

                        IAutoShape text1 = (IAutoShape)page3.Shapes[3];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);

                        DataTable dt_jpscbx = new DataTable();
                        dt_jpscbx.Columns.Add(Base_Config_Jzgj.竞争格局名称);
                        dt_jpscbx.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt_jpscbx.Columns.Add(Base_Config_Jzgj.业态);

                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.本周_新开套内均价);

                        dt_jpscbx.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt_jpscbx.Columns.Add(Base_Config_Cjba.上周_套内均价);
                        dt_jpscbx.Columns.Add(Base_Config_Cjba.上周_建面均价);

                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.上周_认购套内均价);
                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.上周_认购建面均价);

                        dt_jpscbx.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt_jpscbx.Columns.Add(Base_Config_Cjba.本周_套内均价);
                        dt_jpscbx.Columns.Add(Base_Config_Cjba.本周_建面均价);
                                                              
                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);
                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);

                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.本周_认购套内均价环比);
                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.本周_认购建面均价环比);
                        dt_jpscbx.Columns.Add(Base_Config_Rgsj.本周_本周库存);

                        dt_jpscbx = GET_JPBA_BX(dt_jpscbx, item);
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            //获取竞品项目数据
                            dt_jpscbx = GET_JPXM_BX(dt_jpscbx, item.jpxmlb);
                            Office_Tables.SetJP_XIANGYU_biaoxian_Table(page3, dt_jpscbx, 0, null, null);
                            t.AddClone(page3);
                        }
                        #endregion

                        #region 竞品近期动作
                        var page4 = temp[3];
                        IAutoShape text3 = (IAutoShape)page4.Shapes[2];
                        text3.TextFrame.Text = string.Format(text3.TextFrame.Text, item.bamc,item.ytcs[0]);

                        DataTable dt1 = new DataTable();
                        dt1.Columns.Add(Base_Config_Jzgj.竞争格局名称);
                        dt1.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt1.Columns.Add(Base_Config_Jzgj.业态);
                        dt1.Columns.Add(Base_Config_Rgsj.本周_优惠);
                        dt1.Columns.Add(Base_Config_Rgsj.本周_营销动作);
                        dt1.Columns.Add("近期加推");
                        dt1.Columns.Add("办卡方式及蓄客情况");
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt1 = GET_JPXM_BX(dt1, item.jpxmlb);
                            Office_Tables.SetTable(page4, dt1, 1, null, null);
                        }
                        t.AddClone(page4);
                        Base_Log.Log("近期动作开始");



                        #endregion
                        #region 推广图片
                        foreach (var page5 in _plus_jp_dyt_tgtp(item))
                        {
                            t.AddClone(page5);
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
        /// <summary>
        /// 象屿周报竞品
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_xiangyu_2(string str, int cjbh)
        {
            try
            {

                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);
                var tp_nomal = new Presentation(str);
                var temp1 = tp_nomal.Slides;


                #region p1

                #endregion
                #region p2 全市商品住宅连续二十六周（半年）供求价走势
                var page2 = temp1[1];
                #region  数据

                
                DataTable data2_0 = ZB_Data_XZYS_DataProvider.GET_ZB(Base_date.CalcWeekDay_first(Base_date.bn,Base_date.bz-26), Base_date.bz_Last);
                DataTable data2_1 = ZB_Data_CJBA_DataProvider.GET_ZB(Base_date.CalcWeekDay_first(Base_date.bn, Base_date.bz - 26), Base_date.bz_Last);
                var jbz_cjba_zz = (from a in data2_1.AsEnumerable()
                                   where (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                                   group a by new { nf=a["nf"],zc = a["zc"], zcmc = a["zcmc"] } into s
                                   select new
                                   {
                                       nf= s.Key.nf,
                                       zc = s.Key.zc,
                                       zcmc = s.Key.zcmc,
                                       cjje = s.Sum(a => a["cjje"].longs()),
                                       jzmj = s.Sum(a => a["jzmj"].doubls()),
                                   }).OrderBy(m => m.zc).ToList();
                var jbz_xzys_zz = (from a in data2_0.AsEnumerable()
                                   where (a["tyyt"].ToString() == "别墅" || a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "小高层" || a["tyyt"].ToString() == "洋房" || a["tyyt"].ToString() == "洋楼")
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
                                   zcmc = a.nf+"年"+a.zc+"周",
                                   xzgyl = tt == null ? 0 : tt.xzgy,//这里主要第二个集合有可能为空。需要判断
                                   cjmj = a.jzmj,
                                   jmjj = a.cjje / a.jzmj
                               }).ToList();

                var bz_xzys_jzmj = Cache_data_xzys.bz.AsEnumerable().Where(m => new[] { "别墅", "高层", "小高层", "洋房", "洋楼" }.Contains(m["tyyt"])).Sum(m => m["jzmj"].doubls());
                var sz_xzys_jzmj = Cache_data_xzys.sz.AsEnumerable().Where(m=>new []{ "别墅", "高层", "小高层", "洋房", "洋楼" }.Contains( m["tyyt"])).Sum(m => m["jzmj"].doubls());
                var bz_cjba_jzmj = Cache_data_cjjl.bz.AsEnumerable().Where(m => new[] { "别墅", "高层", "小高层", "洋房", "洋楼" }.Contains(m["yt"])).Sum(m => m["jzmj"].doubls());
                var sz_cjba_jzmj = Cache_data_cjjl.sz.AsEnumerable().Where(m => new[] { "别墅", "高层", "小高层", "洋房", "洋楼" }.Contains(m["yt"])).Sum(m => m["jzmj"].doubls());
                var bz_cjba_cjje = Cache_data_cjjl.bz.AsEnumerable().Where(m => new[] { "别墅", "高层", "小高层", "洋房", "洋楼" }.Contains(m["yt"])).Sum(m => m["cjje"].longs());
                var sz_cjba_cjje = Cache_data_cjjl.sz.AsEnumerable().Where(m => new[] { "别墅", "高层", "小高层", "洋房", "洋楼" }.Contains(m["yt"])).Sum(m => m["cjje"].longs());
                #endregion
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
                Office_Charts.Chart_gxfx(page2, dt_zz, 1);

                IAutoShape text0_1 = (IAutoShape)page2.Shapes[2];
                text0_1.TextFrame.Text = string.Format(text0_1.TextFrame.Text,
                      bz_xzys_jzmj.mj_wf(),
                     ((bz_xzys_jzmj - sz_xzys_jzmj) / sz_xzys_jzmj).ss_bfb(),
                     bz_cjba_jzmj.mj_wf(),
                    ((bz_cjba_jzmj - sz_cjba_jzmj) / bz_cjba_jzmj).ss_bfb(),
                     (bz_cjba_cjje / bz_cjba_jzmj).je_y(),
                    (((bz_cjba_cjje / bz_cjba_jzmj) - (sz_cjba_cjje / sz_cjba_jzmj)).doubls() / (sz_cjba_cjje / sz_cjba_jzmj).doubls()).ss_bfb()
                    );

                t.AddClone(page2);

                #endregion
                #region p3 全市区域单周成交数据
                var page3 = temp1[2];
                DataTable dt3 = new DataTable();
                dt3.Columns.Add(Base_Config_TJXM.区域);
                dt3.Columns.Add(Base_Config_TJXM.建筑面积);
                dt3.Columns.Add(Base_Config_TJXM.备案套数);
                dt3.Columns.Add(Base_Config_TJXM.成交金额);
                dt3.Columns.Add(Base_Config_TJXM.建面均价);
                dt3 = _plus_qy_ba_zdpm(new []{ "别墅", "高层", "小高层", "洋房", "洋楼" }, 10, dt3);
                Office_Tables.SetTable(page3, dt3,1, null, null);
                t.AddClone(page3);
                #endregion
                #region p4 全市住宅周签约排行榜
                var page4 = temp1[3];
                #region 数据
                var data4_0 = (from a in Cache_data_cjjl.bz.AsEnumerable()
                                   where (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                                   group a by new { lpmc = a["lpmc"],qy=a["qy"],zt=a["zt"] } into s
                                   select new
                                   {
                                       lpmc=s.Key.lpmc,
                                       qy=s.Key.qy,
                                       zt =s.Key.zt,
                                       cjts = s.Sum(a => a["ts"].doubls()),
                                       jzmj = s.Sum(a => a["jzmj"].doubls()),
                                       cjje = s.Sum(a => a["cjje"].longs()),
                                       jmjj = s.Sum(a => a["cjje"].longs()) / s.Sum(a => a["jzmj"].doubls())
                                   }).OrderByDescending(m => m.cjje).Take(10).ToList();
                DataTable dt4_0 = new DataTable();
                dt4_0.Columns.Add("pm");
                dt4_0.Columns.Add("lpmc");
                dt4_0.Columns.Add("qy");
                dt4_0.Columns.Add("zt");
                dt4_0.Columns.Add("cjts");
                dt4_0.Columns.Add("jzmj");
                dt4_0.Columns.Add("cjje");
                dt4_0.Columns.Add("jmjj");
                for (int i = 0; i < data4_0.Count; i++)
                {
                    DataRow dr = dt4_0.NewRow();
                    dr["pm"] = i+1;
                    dr["lpmc"] = data4_0[i].lpmc;
                    dr["qy"] = data4_0[i].qy;
                    dr["zt"] = data4_0[i].zt;
                    dr["cjts"] = data4_0[i].cjts;
                    dr["jzmj"] = data4_0[i].jzmj.mj();
                    dr["cjje"] = data4_0[i].cjje.je_wy();
                    dr["jmjj"] = data4_0[i].jmjj.je_y();
                    dt4_0.Rows.Add(dr);
                }
                #endregion
                Office_Tables.SetTable(page4, dt4_0, 0, null, null);
                IAutoShape text4_1 = (IAutoShape)page4.Shapes[1];
                text4_1.TextFrame.Text = string.Format(text4_1.TextFrame.Text,
                      bz_xzys_jzmj.mj_wf(),
                     ((bz_xzys_jzmj - sz_xzys_jzmj) / sz_xzys_jzmj).ss_bfb(),
                     bz_cjba_jzmj.mj_wf(),
                    ((bz_cjba_jzmj - sz_cjba_jzmj) / bz_cjba_jzmj).ss_bfb(),
                     (bz_cjba_cjje / bz_cjba_jzmj).je_y(),
                    (((bz_cjba_cjje / bz_cjba_jzmj) - (sz_cjba_cjje / sz_cjba_jzmj)).doubls() / (sz_cjba_cjje / sz_cjba_jzmj).doubls()).ss_bfb()
                    );
                t.AddClone(page4);
                #endregion

                #region p5 象屿观音桥项目竞品周分析
                foreach (var item in param )
                {
                    var tp = new Presentation(str);
                    var temp = tp.Slides;

                    #region 竞品成交情况
                    var page5 = temp[4];
                    DataTable data5_0 = new DataTable();
                    data5_0.Columns.Add(Base_Config_Jzgj.项目名称);
                    data5_0.Columns.Add(Base_Config_Cjba.上上上周_备案套数);
                    data5_0.Columns.Add(Base_Config_Cjba.上上周_备案套数);
                    data5_0.Columns.Add(Base_Config_Cjba.上周_备案套数);
                    data5_0.Columns.Add(Base_Config_Cjba.本周_备案套数);
                    data5_0.Columns.Add(Base_Config_Cjba.上上上周_建面均价);
                    data5_0.Columns.Add(Base_Config_Cjba.上上周_建面均价);
                    data5_0.Columns.Add(Base_Config_Cjba.上周_建面均价);
                    data5_0.Columns.Add(Base_Config_Cjba.本周_建面均价);
                    
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        data5_0 = GET_JPXM_ZFX(data5_0, item.jpxmlb);
                        Office_Tables.SetJP_XIANGYU_ZHOUBAO_Table(page5, data5_0, 0, null, null);
                    }
                    IAutoShape text5_1 = (IAutoShape)page5.Shapes[1];
                    text5_1.TextFrame.Text = string.Format(text5_1.TextFrame.Text,item.bamc);
                    t.AddClone(page5);
                    #endregion
                    #region 竞品存量情况
                    var page6 = temp[5];
                    DataTable data5_1 = new DataTable();
                    data5_1.Columns.Add(Base_Config_Jzgj.组团);
                    data5_1.Columns.Add(Base_Config_Jzgj.项目名称);
                    data5_1.Columns.Add("kcsl");//无数据
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        data5_1 = GET_JPXM_BX(data5_1, item.jpxmlb);
                        Office_Tables.SetTable(page6, data5_1, 0, null, null);
                    }
                    t.AddClone(page6);
                    #endregion
                    #region 竞品动态
                    var page7 = temp[6];

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

        public ISlideCollection _plus_jp_xiangyu_3(string str, int cjbh)
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
                        var page1 = temp[0];

                        IAutoShape text1 = (IAutoShape)page1.Shapes[0];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);

                        DataTable dt_jpbasj = new DataTable();
                        dt_jpbasj.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt_jpbasj.Columns.Add(Base_Config_Jzgj.业态);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_建筑面积);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_建面均价);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_套均总价);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba_bh.本周_套数变化);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_备案套数环比);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba_bh.本周_价格变化);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_建面均价环比);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.上周_建筑面积);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.上周_建面均价);
                        dt_jpbasj.Columns.Add(Base_Config_Cjba.上周_套均总价);


                        //获取本案数据
                        dt_jpbasj = GET_JPBA_BX(dt_jpbasj, item);
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            //获取竞品项目数据
                            dt_jpbasj = GET_JPXM_ZFX(dt_jpbasj, item.jpxmlb);
                            Office_Tables.SetJP_XIANGYU_JINGPINGBEIAN_1_Table(page1, dt_jpbasj, 1, null, null);
                            t.AddClone(page1);
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
                if(item.ytcs==null||item.ytcs.Length<=0)
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0]);
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0]);

                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0]);
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW("", dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));
                }
                else if (item.ytcs[0] == "别墅")
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


        public System.Data.DataTable GET_JPBA_BX(System.Data.DataTable dt, JP_BA_INFO item)
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
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        var temp_rgsj_sy = by.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW_BA(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));

                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);

                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));
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

                    dt.Rows.Add(GET_ROW_BA(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));
                }
            }
            else if (item.ytcs[0] == "商业")
            {

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

                dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));
            }
            return dt;
        }



        public System.Data.DataTable GET_JPXM_ZFX(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {
                if (item.ytcs == null || item.ytcs.Length <= 0)
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                 
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0]);
                    var temp_cjba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0]);
                    var temp_cjba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0]);

                   
                    #endregion

                    dt.Rows.Add(GET_ROW_JP_ZFX("", dr1, dt, temp_cjba_bz, temp_cjba_sz, temp_cjba_ssz, temp_cjba_sssz, item));
                }
                else if (item.ytcs[0] == "别墅")
                {
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {

                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        #endregion

                        dt.Rows.Add(GET_ROW_JP_ZFX(item.xfytcs[i], dr1, dt, temp_cjba_bz, temp_cjba_sz, temp_cjba_ssz, temp_cjba_sssz, item));

                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    if (item.hxcs != null && item.hxcs.Length > 0)
                    {
                        for (int i = 0; i < item.hxcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                            var temp_cjba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                            var temp_cjba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                            #endregion

                            dt.Rows.Add(GET_ROW_JP_ZFX(item.hxcs[i], dr1, dt, temp_cjba_bz, temp_cjba_sz, temp_cjba_ssz, temp_cjba_sssz, item));
                        }
                    }
                    else if(item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.hxcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_cjba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_cjba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            #endregion

                            dt.Rows.Add(GET_ROW_JP_ZFX(item.xfytcs[i], dr1, dt, temp_cjba_bz, temp_cjba_sz, temp_cjba_ssz, temp_cjba_sssz, item));
                        }
                    }
                    else
                    {

                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains( m["yt"].ToString()));
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                            var temp_cjba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                            var temp_cjba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                            #endregion

                            dt.Rows.Add(GET_ROW_JP_ZFX(string.Join(",",item.ytcs), dr1, dt, temp_cjba_bz, temp_cjba_sz, temp_cjba_ssz, temp_cjba_sssz, item));
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    var temp_cjba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    var temp_cjba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    #endregion

                    dt.Rows.Add(GET_ROW_JP_ZFX(string.Join(",", item.ytcs), dr1, dt, temp_cjba_bz, temp_cjba_sz, temp_cjba_ssz, temp_cjba_sssz, item));
                }



            }


            return dt;
        }

        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                                DataRow temp_rg_bz,
                                DataRow temp_rg_sz,
                                DataRow temp_rg_ssz,
                                DataRow temp_rg_sssz,
                                EnumerableRowCollection<DataRow> temp_cj_sy,
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
                                dr1[dt.Columns[j].ColumnName] = temp_rg_bz != null ? temp_rg_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.本周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_rg_bz != null && temp_rg_bz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_rg_bz[Base_Config_Rgsj.本周_认购套内均价._ConfigRgsjMc()].doubls() * temp_rg_bz[Base_Config_Rgsj.本周_认购套内体量._ConfigRgsjMc()].doubls() / temp_rg_bz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
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
                                dr1[dt.Columns[j].ColumnName] = temp_rg_sz != null ? temp_rg_sz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.上周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_rg_sz != null && temp_rg_sz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_rg_sz[Base_Config_Rgsj.本周_认购套内均价._ConfigRgsjMc()].doubls() * temp_rg_sz[Base_Config_Rgsj.本周_认购套内体量._ConfigRgsjMc()].doubls() / temp_rg_sz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
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
                                dr1[dt.Columns[j].ColumnName] = temp_rg_ssz != null ? temp_rg_ssz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.上上周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_rg_ssz != null && temp_rg_ssz[Base_Config_Rgsj.上上周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_rg_ssz[Base_Config_Rgsj.上上周_认购套内均价._ConfigRgsjMc()].doubls() * temp_rg_ssz[Base_Config_Rgsj.上上周_认购套内体量._ConfigRgsjMc()].doubls() / temp_rg_ssz[Base_Config_Rgsj.上上周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
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
                                dr1[dt.Columns[j].ColumnName] = temp_rg_sssz != null ? temp_rg_sssz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.上上上周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_rg_sssz != null && temp_rg_sssz[Base_Config_Rgsj.上上上周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_rg_sssz[Base_Config_Rgsj.上上上周_认购套内均价._ConfigRgsjMc()].doubls() * temp_rg_sssz[Base_Config_Rgsj.上上上周_认购套内体量._ConfigRgsjMc()].doubls() / temp_rg_sssz[Base_Config_Rgsj.上上上周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
                            }; break;
                        default:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_rg_bz != null ? temp_rg_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : "-";
                            }; break;
                    }
                }
                else if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                {

                }
                else if (Base_Config_Cjba_BY._备案数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Cjba_BY.本月_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cj_sy != null ? temp_cj_sy.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba_BY.本月_建面均价:
                            {

                                if ((temp_cj_sy != null && temp_cj_sy.Sum(m => m[Base_Config_Cjba_BY.本月_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cj_sy.Sum(m => m[Base_Config_Cjba_BY.本月_成交金额._ConfigCjbaMc()].longs()) / temp_cj_sy.Sum(m => m[Base_Config_Cjba_BY.本月_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "0";
                                }
                            }; break;
                    }

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
        //竞品备案本案数据
        public DataRow GET_ROW_JPBA_BA(string yt, DataRow dr1, System.Data.DataTable dt,
                              EnumerableRowCollection<DataRow> temp_cj_bz,
                              EnumerableRowCollection<DataRow> temp_cj_sz,
                              EnumerableRowCollection<DataRow> temp_cj_by,
                              JP_BA_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Cjba.本周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cj_bz != null ? temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba.本周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cj_bz != null ? temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba.本周_建面均价: {
                                if ((temp_cj_bz != null && temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "0";
                                }
                            }; break;
                        case Base_Config_Cjba.本周_套均总价: {
                                long bz_cjje = temp_cj_bz.Sum(m => m["cjje"].ints());
                                long bz_ts = temp_cj_bz.Sum(m => m["ts"].ints());
                                dr1[dt.Columns[j].ColumnName] = bz_cjje / bz_ts;
                            }; break;
                        case Base_Config_Cjba_bh.本周_套数变化:
                            {
                                long bz_ts = temp_cj_bz.Sum(m => m["ts"].ints());
                                long sz_ts = temp_cj_sz.Sum(m => m["ts"].ints());
                                dr1[dt.Columns[j].ColumnName] = bz_ts - sz_ts;
                            }; break;
                        case Base_Config_Cjba.本周_备案套数环比:
                            {
                                dr1[dt.Columns[j].ColumnName] = ((temp_cj_bz.Sum(m => m["ts"].ints()) - temp_cj_sz.Sum(m => m["ts"].ints())) / temp_cj_sz.Sum(m => m["ts"].ints())).doubls().ss_bfb();
                            }; break;
                        case Base_Config_Cjba_bh.本周_价格变化:
                            {
                                long bz_je = temp_cj_bz.Sum(m => m["cjje"].ints());
                                long bz_mj = temp_cj_bz.Sum(m => m["jzmj"].ints());
                                long sz_je = temp_cj_sz.Sum(m => m["cjje"].ints());
                                long sz_mj = temp_cj_sz.Sum(m => m["jzmj"].ints());
                                dr1[dt.Columns[j].ColumnName] = bz_je / bz_mj - sz_je / sz_mj;
                            };break;
                        case Base_Config_Cjba.本周_建面均价环比:
                            {
                                long bz_je = temp_cj_bz.Sum(m => m["cjje"].ints());
                                long bz_mj = temp_cj_bz.Sum(m => m["jzmj"].ints());
                                long sz_je = temp_cj_sz.Sum(m => m["cjje"].ints());
                                long sz_mj = temp_cj_sz.Sum(m => m["jzmj"].ints());
                                dr1[dt.Columns[j].ColumnName] = (bz_je / bz_mj - sz_je / sz_mj)/ (sz_je / sz_mj);
                            }; break;
                        case Base_Config_Cjba.上周_备案套数: 
                        case Base_Config_Cjba.上周_建筑面积:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_cj_sz != null ? temp_cj_sz.Sum(m => m[dt.Columns[j].ColumnName._ConfigCjbaMc()].ints()) : 0;
                            }; break;
                        case Base_Config_Cjba.上周_建面均价:
                            {

                                long sz_je = temp_cj_sz.Sum(m => m["cjje"].ints());
                                long sz_mj = temp_cj_sz.Sum(m => m["jzmj"].ints());
                                dr1[dt.Columns[j].ColumnName] = sz_mj != 0 ? (sz_je / sz_mj).je_y() : 0;
                            }; break;
                        case Base_Config_Cjba.上周_套均总价:
                            {

                                long sz_je = temp_cj_sz.Sum(m => m["cjje"].ints());
                                long sz_ts = temp_cj_sz.Sum(m => m["ts"].ints());
                                dr1[dt.Columns[j].ColumnName] = sz_ts != 0 ? (sz_je / sz_ts).je_y() : 0;
                            }; break;
                    }
                }
                else if (Base_Config_Cjba_BY._备案数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Cjba_BY.本月_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cj_by != null ? temp_cj_by.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba_BY.本月_建面均价:
                            {

                                if ((temp_cj_by != null && temp_cj_by.Sum(m => m[Base_Config_Cjba_BY.本月_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cj_by.Sum(m => m[Base_Config_Cjba_BY.本月_成交金额._ConfigCjbaMc()].longs()) / temp_cj_by.Sum(m => m[Base_Config_Cjba_BY.本月_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "0";
                                }
                            }; break;
                    }

                }
                else if (Base_Config_Jzgj._竞争格局参数名称.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Jzgj.组团: { dr1[dt.Columns[j].ColumnName] = item.ztcs[0]; }; break;
                        case Base_Config_Jzgj.项目名称: { dr1[dt.Columns[j].ColumnName] = item.lpcs[0]; }; break;
                        case Base_Config_Jzgj.业态: { dr1[dt.Columns[j].ColumnName] = yt; }; break;
                        case Base_Config_Jzgj.竞争格局名称: { dr1[dt.Columns[j].ColumnName] = "本案"; }; break;
                        case Base_Config_Jzgj.竞争格局_主力面积区间: { dr1[dt.Columns[j].ColumnName] = item.zlmjqj; }; break;
                        default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                    }

                }

            }

            return dr1;
        }

        public DataRow GET_ROW_JP_ZFX(string yt, DataRow dr1, System.Data.DataTable dt,
                              EnumerableRowCollection<DataRow> temp_cj_bz,
                              EnumerableRowCollection<DataRow> temp_cj_sz,
                              EnumerableRowCollection<DataRow> temp_cj_ssz,
                              EnumerableRowCollection<DataRow> temp_cj_sssz,
                              JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName)|| Base_Config_Cjba_bh._备案数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Cjba.本周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cj_bz != null ? temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba.本周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cj_bz != null ? temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba.本周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cj_bz != null ? temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].ints()) : 0; }; break;

                        case Base_Config_Cjba.本周_建面均价:
                            {
                                if ((temp_cj_bz != null && temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cj_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "0";
                                }
                            }; break;
                        case Base_Config_Cjba.本周_套均总价:
                            {
                                long bz_cjje = temp_cj_bz.Sum(m => m["cjje"].ints());
                                long bz_ts = temp_cj_bz.Sum(m => m["ts"].ints());
                                dr1[dt.Columns[j].ColumnName] = bz_ts != 0 ? (bz_cjje / bz_ts).je_wy() : 0 ;
                            }; break;
                        case Base_Config_Cjba_bh.本周_套数变化:
                            {
                                long bz_ts = temp_cj_bz.Sum(m => m["ts"].ints());
                                long sz_ts = temp_cj_sz.Sum(m => m["ts"].ints());
                                dr1[dt.Columns[j].ColumnName] = bz_ts - sz_ts;
                            }; break;
                        case Base_Config_Cjba.本周_备案套数环比:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_cj_sz.Sum(m => m["ts"].ints())!=0&& temp_cj_bz.Sum(m => m["ts"].ints())!=0 ? ((temp_cj_bz.Sum(m => m["ts"].ints()) - temp_cj_sz.Sum(m => m["ts"].ints())) / temp_cj_sz.Sum(m => m["ts"].doubls())).ss_bfb_sz()+"%":"—";
                            }; break;
                        case Base_Config_Cjba_bh.本周_价格变化:
                            {
                                long bz_je = temp_cj_bz.Sum(m => m["cjje"].ints());
                                double bz_mj = temp_cj_bz.Sum(m => m["jzmj"].doubls());
                                long sz_je = temp_cj_sz.Sum(m => m["cjje"].ints());
                                double sz_mj = temp_cj_sz.Sum(m => m["jzmj"].doubls());
                                if (sz_mj != 0 && bz_mj != 0)
                                {

                                    dr1[dt.Columns[j].ColumnName] = ((bz_je / bz_mj) - (sz_je / sz_mj)).je_y();
                                }
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "—";
                                }
                               
                            }; break;
                        case Base_Config_Cjba.本周_建面均价环比:
                            {
                                long bz_je = temp_cj_bz.Sum(m => m["cjje"].longs());
                                double bz_mj = temp_cj_bz.Sum(m => m["jzmj"].doubls());
                                long sz_je = temp_cj_sz.Sum(m => m["cjje"].longs());
                                double sz_mj = temp_cj_sz.Sum(m => m["jzmj"].doubls());
                                if (sz_mj != 0&&bz_mj != 0 && sz_je != 0)
                                {
                                        dr1[dt.Columns[j].ColumnName] = (((bz_je / bz_mj) - (sz_je / sz_mj)) / (sz_je / sz_mj)).ss_bfb_sz();
                                }
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "—";
                                }

                                
                            }; break;
                        case Base_Config_Cjba.上周_备案套数:
                        case Base_Config_Cjba.上周_建筑面积:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_cj_sz != null ? temp_cj_sz.Sum(m => m[dt.Columns[j].ColumnName._ConfigCjbaMc()].ints()) : 0;
                            }; break;
                        case Base_Config_Cjba.上周_建面均价:
                            {

                                long sz_je = temp_cj_sz.Sum(m => m["cjje"].ints());
                                long sz_mj = temp_cj_sz.Sum(m => m["jzmj"].ints());
                                dr1[dt.Columns[j].ColumnName] = sz_mj != 0 ? (sz_je / sz_mj).je_y() : 0;
                            }; break;
                        case Base_Config_Cjba.上周_套均总价:
                            {

                                long sz_je = temp_cj_sz.Sum(m => m["cjje"].ints());
                                long sz_ts = temp_cj_sz.Sum(m => m["ts"].ints());
                                dr1[dt.Columns[j].ColumnName] = sz_ts != 0 ? (sz_je / sz_ts).je_wy() : 0;
                            }; break;
                        case Base_Config_Cjba.上上周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cj_ssz != null ? temp_cj_ssz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba.上上周_建面均价:
                            {
                                if ((temp_cj_ssz != null && temp_cj_ssz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cj_ssz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cj_ssz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "0";
                                }
                            }; break;
                        case Base_Config_Cjba.上上上周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cj_sssz != null ? temp_cj_sssz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba.上上上周_建面均价:
                            {
                                if ((temp_cj_sssz != null && temp_cj_sssz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cj_sssz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cj_sssz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "0";
                                }
                            }; break;
                    }
                }
                
                else if (Base_Config_Jzgj._竞争格局参数名称.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Jzgj.组团: { dr1[dt.Columns[j].ColumnName] = item.ztcs[0]; }; break;
                        case Base_Config_Jzgj.项目名称: { dr1[dt.Columns[j].ColumnName] = item.lpcs[0]; }; break;
                        case Base_Config_Jzgj.业态: { dr1[dt.Columns[j].ColumnName] = yt; }; break;
                        case Base_Config_Jzgj.竞争格局名称: { dr1[dt.Columns[j].ColumnName] = "本案"; }; break;
                        case Base_Config_Jzgj.竞争格局_主力面积区间: { dr1[dt.Columns[j].ColumnName] = item.zlmjqj; }; break;
                        default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                    }

                }

            }

            return dr1;
        }

        public ISlideCollection _plus_jp_dyt_jzgj(ISlide sld, JP_BA_INFO item)
        {
            try
            {
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);
                var page = sld;
                #region 商务
                if (item.ytcs[0] == "商务")
                {

                    IAutoShape text1 = (IAutoShape)page.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                    //数据
                    System.Data.DataTable jzgjt = new System.Data.DataTable();
                    jzgjt.Columns.Add("");
                    jzgjt.Columns.Add("成交套数", typeof(int));
                    jzgjt.Columns.Add("建面均价", typeof(double));
                    //图表
                    Aspose.Slides.Charts.IChart chart = (Aspose.Slides.Charts.IChart)page.Shapes[3];
                    foreach (var item_jp in item.jpxmlb)
                    {
                        if (item_jp.hxcs != null)
                        {
                            for (int i = 0; i < item_jp.hxcs.Length; i++)
                            {
                                var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item_jp.hxcs[i]).FirstOrDefault();

                                DataRow dr1 = jzgjt.NewRow();
                                dr1[0] = item_jp.lpcs[0] + "(" + item.hxcs[i] + ")";
                                if (jpcjxx != null)
                                {

                                    dr1[1] = jpcjxx[Base_Config_Rgsj.本周_认购套数._ConfigCjbaMc()].ints();
                                    dr1[2] = jpcjxx[Base_Config_Rgsj.本周_认购建面均价._ConfigCjbaMc()].ints();
                                }
                                else
                                {
                                    dr1[1] = 0;
                                    dr1[2] = 0;
                                }
                                jzgjt.Rows.Add(dr1);
                            }

                        }
                    }
                    Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                    t.AddClone(page);

                }
                #endregion

                #region 别墅


                else if (item.ytcs[0] == "别墅")
                {
                    IAutoShape text1 = (IAutoShape)page.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text,  item.bamc, item.ytcs[0]);
                    System.Data.DataTable jzgjt = new System.Data.DataTable();
                    jzgjt.Columns.Add("");
                    jzgjt.Columns.Add("成交套数", typeof(int));
                    jzgjt.Columns.Add("建面均价", typeof(double));
                    foreach (var item_jp in item.jpxmlb)
                    {
                        if (item_jp.xfytcs != null)
                        {
                            for (int i = 0; i < item_jp.xfytcs.Length; i++)
                            {

                                var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item_jp.xfytcs[i]).FirstOrDefault();

                                DataRow dr1 = jzgjt.NewRow();
                                dr1[0] = item_jp.lpcs[0] + "(" + item_jp.xfytcs[i] + ")";
                                if (jpcjxx != null)
                                {

                                    dr1[1] = jpcjxx[Base_Config_Rgsj.本周_认购套数._ConfigCjbaMc()].ints();
                                    dr1[2] = jpcjxx[Base_Config_Rgsj.本周_认购建面均价._ConfigCjbaMc()].ints();
                                    jzgjt.Rows.Add(dr1);
                                }
                                else
                                {
                                    if (item.xfytcs != null && item_jp.xfytcs.Contains(item.xfytcs[i]))
                                    {
                                        dr1[1] = 0;
                                        dr1[2] = 0;
                                        jzgjt.Rows.Add(dr1);
                                    }
                                    else
                                        continue;
                                }
                            }

                        }
                        else
                        {
                            var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item_jp.ytcs[0]).FirstOrDefault();

                            DataRow dr1 = jzgjt.NewRow();
                            dr1[0] = item_jp.lpcs[0] + "(" + item_jp.ytcs[0] + ")";
                            if (jpcjxx != null)
                            {
                                dr1[1] = jpcjxx[Base_Config_Rgsj.本周_认购套数._ConfigCjbaMc()].ints();
                                dr1[2] = jpcjxx[Base_Config_Rgsj.本周_认购建面均价._ConfigCjbaMc()].ints();
                                jzgjt.Rows.Add(dr1);
                            }
                            else
                            {
                                if (item_jp.ytcs.Contains(item.ytcs[0]))
                                {
                                    dr1[1] = 0;
                                    dr1[2] = 0;
                                    jzgjt.Rows.Add(dr1);
                                }
                                else
                                    continue;
                            }
                        }

                    }
                    Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                    t.AddClone(page);


                }


                #endregion

                #region 大业态


                else
                {
                    IAutoShape text1 = (IAutoShape)page.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                    //数据
                    System.Data.DataTable jzgjt = new System.Data.DataTable();
                    jzgjt.Columns.Add("");
                    jzgjt.Columns.Add("成交套数", typeof(int));
                    jzgjt.Columns.Add("建面均价", typeof(double));
                    foreach (var item_jp in item.jpxmlb)
                    {
                        string jpyt = item_jp.ytcs == null ? item.ytcs[0] : item_jp.ytcs[0];
                        var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == jpyt).FirstOrDefault();

                        DataRow dr1 = jzgjt.NewRow();
                        dr1[0] = item_jp.lpcs[0] + "(" + item.ytcs[0] + ")";
                        if (jpcjxx != null)
                        {

                            dr1[1] = jpcjxx[Base_Config_Rgsj.本周_认购套数._ConfigCjbaMc()].ints();
                            dr1[2] = jpcjxx[Base_Config_Rgsj.本周_认购建面均价._ConfigCjbaMc()].ints();
                        }
                        else
                        {
                            dr1[1] = 0;
                            dr1[2] = 0;
                        }
                        jzgjt.Rows.Add(dr1);


                    }
                    Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                    t.AddClone(page);
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

        public ISlideCollection _plus_jp_dyt_jzgj_taonei(ISlide sld, JP_BA_INFO item)
        {
            try
            {
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);
                var page = sld;
                IAutoShape text1 = (IAutoShape)page.Shapes[19];
                text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                #region 商务
                if (item.ytcs[0] == "商务")
                {
                    
                    //数据
                    System.Data.DataTable jzgjt = new System.Data.DataTable();
                    jzgjt.Columns.Add("");
                    jzgjt.Columns.Add("成交套数", typeof(int));
                    jzgjt.Columns.Add("套内均价", typeof(double));
                    //图表
                    //Aspose.Slides.Charts.IChart chart = (Aspose.Slides.Charts.IChart)page.Shapes[20];
                    foreach (var item_jp in item.jpxmlb)
                    {
                        if (item_jp.hxcs != null)
                        {
                            for (int i = 0; i < item_jp.hxcs.Length; i++)
                            {
                                var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item_jp.hxcs[i]).FirstOrDefault();

                                DataRow dr1 = jzgjt.NewRow();
                                dr1[0] = item_jp.lpcs[0] + "(" + item.hxcs[i] + ")";
                                if (jpcjxx != null)
                                {

                                    dr1[1] = jpcjxx[Base_Config_Rgsj.本周_认购套数._ConfigCjbaMc()].ints();
                                    dr1[2] = jpcjxx[Base_Config_Rgsj.本周_认购套内均价._ConfigCjbaMc()].ints();
                                }
                                else
                                {
                                    dr1[1] = 0;
                                    dr1[2] = 0;
                                }
                                jzgjt.Rows.Add(dr1);
                            }

                        }
                    }
                    Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 20);
                    t.AddClone(page);

                }
                #endregion

                #region 别墅


                else if (item.ytcs[0] == "别墅")
                { 
                    System.Data.DataTable jzgjt = new System.Data.DataTable();
                    jzgjt.Columns.Add("");
                    jzgjt.Columns.Add("成交套数", typeof(int));
                    jzgjt.Columns.Add("套内均价", typeof(double));
                    foreach (var item_jp in item.jpxmlb)
                    {
                        if (item_jp.xfytcs != null)
                        {
                            for (int i = 0; i < item_jp.xfytcs.Length; i++)
                            {

                                var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item_jp.xfytcs[i]).FirstOrDefault();

                                DataRow dr1 = jzgjt.NewRow();
                                dr1[0] = item_jp.lpcs[0] + "(" + item_jp.xfytcs[i] + ")";
                                if (jpcjxx != null)
                                {

                                    dr1[1] = jpcjxx[Base_Config_Rgsj.本周_认购套数._ConfigCjbaMc()].ints();
                                    dr1[2] = jpcjxx[Base_Config_Rgsj.本周_认购套内均价._ConfigCjbaMc()].ints();
                                    jzgjt.Rows.Add(dr1);
                                }
                                else
                                {
                                    if (item.xfytcs != null && item_jp.xfytcs.Contains(item.xfytcs[i]))
                                    {
                                        dr1[1] = 0;
                                        dr1[2] = 0;
                                        jzgjt.Rows.Add(dr1);
                                    }
                                    else
                                        continue;
                                }
                            }

                        }
                        else
                        {
                            var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item_jp.ytcs[0]).FirstOrDefault();

                            DataRow dr1 = jzgjt.NewRow();
                            dr1[0] = item_jp.lpcs[0] + "(" + item_jp.ytcs[0] + ")";
                            if (jpcjxx != null)
                            {
                                dr1[1] = jpcjxx[Base_Config_Rgsj.本周_认购套数._ConfigCjbaMc()].ints();
                                dr1[2] = jpcjxx[Base_Config_Rgsj.本周_认购套内均价._ConfigCjbaMc()].ints();
                                jzgjt.Rows.Add(dr1);
                            }
                            else
                            {
                                if (item_jp.ytcs.Contains(item.ytcs[0]))
                                {
                                    dr1[1] = 0;
                                    dr1[2] = 0;
                                    jzgjt.Rows.Add(dr1);
                                }
                                else
                                    continue;
                            }
                        }

                    }
                    Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 20);
                    t.AddClone(page);


                }


                #endregion

                #region 大业态


                else
                {
                   
                    //数据
                    System.Data.DataTable jzgjt = new System.Data.DataTable();
                    jzgjt.Columns.Add("");
                    jzgjt.Columns.Add("成交套数", typeof(int));
                    jzgjt.Columns.Add("套内均价", typeof(double));
                    foreach (var item_jp in item.jpxmlb)
                    {
                        string jpyt = item_jp.ytcs == null ? item.ytcs[0] : item_jp.ytcs[0];
                        var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == jpyt).FirstOrDefault();

                        DataRow dr1 = jzgjt.NewRow();
                        dr1[0] = item_jp.lpcs[0] + "(" + item.ytcs[0] + ")";
                        if (jpcjxx != null)
                        {

                            dr1[1] = jpcjxx[Base_Config_Rgsj.本周_认购套数._ConfigCjbaMc()].ints();
                            dr1[2] = jpcjxx[Base_Config_Rgsj.本周_认购套内均价._ConfigCjbaMc()].ints();
                        }
                        else
                        {
                            dr1[1] = 0;
                            dr1[2] = 0;
                        }
                        jzgjt.Rows.Add(dr1);


                    }
                    Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 20);
                    t.AddClone(page);
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
    }
}
