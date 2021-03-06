﻿using Aspose.Slides;
using Calculation.Base;
using Calculation.Dal;
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
    /// <summary>
    /// 金地地产
    /// </summary>
    public  class plus_jp_jingdidichan : plus_jp_base
    {

        private static DataTable sy;

        public class Base_Config_Cjba_SY
        {

            public const string 上月_备案套数 = "sy_ts";
            public const string 上月_成交金额 = "sy_cjje";
            public const string 上月_建筑面积 = "sy_jzmj";
            public const string 上月_套内面积 = "sy_tnmj";
            public const string 上月_建面均价 = "sy_jmjj";
            public const string 上月_套内均价 = "sy_tnjj";
            public const string 上月_套均总价 = "sy_tjzj";



            public static string[] _备案数据 = { "sy_ts","sy_cjje","sy_jzmj","sy_tnmj","sy_jmjj","sy_tnjj", "sy_tjzj",};
        }

        public plus_jp_jingdidichan()
        {
            Base_date.init_yb(Base_date.bn, Base_date.GET_Y_FROM_Z(Base_date.bn, Base_date.bz));
            sy = ZB_Data_CJBA_DataProvider.GET_ZB(Base_date.sy_First, Base_date.sy_Last);
        }
        public ISlideCollection _plus_jp_jingdidichan_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                int pcount = 1;
                int pagesize = param.Where(m => string.IsNullOrEmpty(m.qtcs)).Count();
                t.RemoveAt(0);
                foreach (var item in param)
                {
                    if (string.IsNullOrEmpty(item.qtcs))
                    {
                        #region MyRegion


                        var tp = new Presentation(str);
                        var temp = tp.Slides;

                        #region 竞品分布
                        var page0 = temp[0];
                        foreach (var jpfb in _plus_jp_dyt_jzgj(page0, item, "(" + pcount + "/" + pagesize + ")"))
                        {
                            t.AddClone(jpfb);
                        }
                        #endregion
                        // t.AddClone(page1);



                        #region 格局统计
                        var page1 = temp[1];
                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.业态);
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);

                        dt.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);


                        dt.Columns.Add(Base_Config_Rgsj.上上上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上上上周_认购建面均价);

                        dt.Columns.Add(Base_Config_Rgsj.上上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上上周_认购建面均价);

                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购建面均价);

                        dt.Columns.Add(Base_Config_Cjba_SY.上月_备案套数);
                        dt.Columns.Add(Base_Config_Cjba_SY.上月_建面均价);

                        dt.Columns.Add(Base_Config_Rgsj.本周_变化原因);


                        IAutoShape text1 = (IAutoShape)page1.Shapes[0];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, "(" + pcount + "/" + pagesize + ")", item.bamc);
                        //金地周报不需要本案
                        //dt = GET_JPBA_BX(dt, item);
                        #endregion
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_JINGDIDICHAN_PT_Table(page1, dt, 1,null,null);
                            t.AddClone(page1);
                        }

                        //竞争项目近期动作
                        var page3 = temp[2];
                        IAutoShape text3 = (IAutoShape)page3.Shapes[1];
                        text3.TextFrame.Text = string.Format(text3.TextFrame.Text,"("+pcount+"/"+ pagesize + ")", item.bamc);

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
                            Office_Tables.SetTable(page3, dt1, 0, null, null);
                        }
                        t.AddClone(page3);
                        Base_Log.Log("近期动作开始");

                        foreach (var page4 in _plus_jp_dyt_tgtp(item))
                        {
                            t.AddClone(page4);
                        }
                        pcount++;
                        #endregion
                    }
                    else
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page4 = temp[3];
                        DataTable dt4 = new DataTable();
                        dt4.Columns.Add("kfs");
                        dt4.Columns.Add("hj");
                        dt4.Columns.Add("sssz_cjje");
                        dt4.Columns.Add("ssz_cjje");
                        dt4.Columns.Add("sz_cjje");
                        dt4.Columns.Add("bz_cjje");
                        dt4 = GET_JPXM_ZT_CJJE(dt4, item.jpxmlb);
                        //并不需要本案
                        //dt4 = GET_JPBA_CJJE(dt4, item);
                        Office_Tables.SetJP_XUHUICHENG_XIAOSHOUE_Table(page4, dt4, 1, null, null);
                        t.AddClone(page4);

                        foreach (var item_jp in item.jpxmlb)
                        {
                            DataTable dt5 = new DataTable();
                            dt5.Columns.Add("kfs");
                            dt5.Columns.Add("hj");
                            dt5.Columns.Add("sssz");
                            dt5.Columns.Add("ssz");
                            dt5.Columns.Add("sz");
                            dt5.Columns.Add("bz");
                            dt5 = GET_JPXM_XF_CJJE(dt5, item_jp);
                            var page5 = new Presentation(str).Slides[4];
                            Office_Tables.SetJP_XUHUICHENG_XIAOSHOUE_Table(page5, dt5, 1, null, null);
                            IAutoShape text5 = (IAutoShape)page5.Shapes[0];
                            text5.TextFrame.Text = string.Format(text5.TextFrame.Text, item_jp.kfs);
                            t.AddClone(page5);
                        }
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
                    if (item.xfytcs != null)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {

                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_ssz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sssz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sy = sy.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                            var temp_rg_ssz = temp_rgsj_ssz.FirstOrDefault();
                            var temp_rg_sssz = temp_rgsj_sssz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_rg_ssz, temp_rg_sssz, temp_rgsj_sy, item));

                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_ssz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sssz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                        var temp_rgsj_sy = sy.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                        var temp_rg_ssz = temp_rgsj_ssz.FirstOrDefault();
                        var temp_rg_sssz = temp_rgsj_sssz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_rg_ssz, temp_rg_sssz, temp_rgsj_sy, item));
                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    for (int i = 0; i < item.hxcs.Length; i++)
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                        var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                        var temp_rgsj_ssz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                        var temp_rgsj_sssz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);

                        var temp_rgsj_sy = sy.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[0]);

                        //本周本案认购数据
                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                        var temp_rg_ssz = temp_rgsj_ssz.FirstOrDefault();
                        var temp_rg_sssz = temp_rgsj_sssz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.hxcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_rg_ssz, temp_rg_sssz, temp_rgsj_sy, item));
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_ssz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sssz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sy = sy.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周本案认购数据
                    var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                    var temp_rg_ssz = temp_rgsj_ssz.FirstOrDefault();
                    var temp_rg_sssz = temp_rgsj_sssz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_rg_ssz, temp_rg_sssz, temp_rgsj_sy, item));
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
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_rgsj_sy = sy.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                        var temp_rg_ssz = temp_rgsj_ssz.FirstOrDefault();
                        var temp_rg_sssz = temp_rgsj_sssz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW_BA(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_rg_ssz, temp_rg_sssz, temp_rgsj_sy, item));

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
                    var temp_rgsj_sy = sy.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周本案认购数据
                    var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                    var temp_rg_ssz = temp_rgsj_ssz.FirstOrDefault();
                    var temp_rg_sssz = temp_rgsj_sssz.FirstOrDefault();
                    
                    #endregion

                    dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_rg_ssz, temp_rg_sssz, temp_rgsj_sy, item));
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
                    var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                    var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                    var temp_rgsj_sy = sy.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                    //本周本案认购数据
                    var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                    var temp_rg_ssz = temp_rgsj_ssz.FirstOrDefault();
                    var temp_rg_sssz = temp_rgsj_sssz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW_BA(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_rg_ssz, temp_rg_sssz, temp_rgsj_sy, item));
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
                var temp_rgsj_bz = Cache_data_rgsj.bz.Select("xm = "+ item.lpcs[0] +" and yt = " + item.ytcs[0]);
                var temp_rgsj_sz = Cache_data_rgsj.sz.Select("xm = " + item.lpcs[0] + " and yt = " + item.ytcs[0]);
                var temp_rgsj_ssz = Cache_data_rgsj.ssz.Select("xm = " + item.lpcs[0] + " and yt = " + item.ytcs[0]);
                var temp_rgsj_sssz = Cache_data_rgsj.sssz.Select("xm = " + item.lpcs[0] + " and yt = " + item.ytcs[0]);
                var temp_rgsj_sy = sy.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                //本周本案认购数据
                var temp_rg_bz = temp_rgsj_bz!=null&&temp_rgsj_bz.Length>0? temp_rgsj_bz.FirstOrDefault():null;
                var temp_rg_sz = temp_rgsj_sz != null && temp_rgsj_sz.Length > 0 ? temp_rgsj_sz.FirstOrDefault() : null;
                var temp_rg_ssz = temp_rgsj_ssz != null && temp_rgsj_ssz.Length > 0 ? temp_rgsj_ssz.FirstOrDefault() : null;
                var temp_rg_sssz = temp_rgsj_sssz != null && temp_rgsj_sssz.Length > 0 ? temp_rgsj_sssz.FirstOrDefault() : null;


                #endregion

                dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_rg_ssz, temp_rg_sssz, temp_rgsj_sy, item));
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
                else if (Base_Config_Cjba_SY._备案数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Cjba_SY.上月_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cj_sy != null ? temp_cj_sy.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba_SY.上月_建面均价:
                            {

                                if ((temp_cj_sy != null && temp_cj_sy.Sum(m => m[Base_Config_Cjba_SY.上月_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cj_sy.Sum(m => m[Base_Config_Cjba_SY.上月_成交金额._ConfigCjbaMc()].longs()) / temp_cj_sy.Sum(m => m[Base_Config_Cjba_SY.上月_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "0";
                                }
                            };break;
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
        public DataRow GET_ROW_BA(string yt, DataRow dr1, System.Data.DataTable dt,
                              DataRow temp_rg_bz,
                              DataRow temp_rg_sz,
                              DataRow temp_rg_ssz,
                              DataRow temp_rg_sssz,
                              EnumerableRowCollection<DataRow> temp_cj_sy,
                              JP_BA_INFO item)
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
                                dr1[dt.Columns[j].ColumnName] = temp_rg_bz != null ? temp_rg_bz[dt.Columns[j].ColumnName] : "-";
                            }; break;
                    }
                }
                else if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                {

                }
                else if (Base_Config_Cjba_SY._备案数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Cjba_SY.上月_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cj_sy != null ? temp_cj_sy.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba_SY.上月_建面均价:
                            {

                                if ((temp_cj_sy != null && temp_cj_sy.Sum(m => m[Base_Config_Cjba_SY.上月_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cj_sy.Sum(m => m[Base_Config_Cjba_SY.上月_成交金额._ConfigCjbaMc()].longs()) / temp_cj_sy.Sum(m => m[Base_Config_Cjba_SY.上月_建筑面积._ConfigCjbaMc()].doubls())).je_y();
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

        public DataTable GET_JPBA_CJJE(DataTable dt, JP_BA_INFO ba)
        {
            var temp_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => ba.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
            var temp_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => ba.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
            var temp_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => ba.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
            var temp_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => ba.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
            DataRow dr = dt.NewRow();
            dr["kfs"] = string.Join(",", ba.kfs);
            dr["hj"] = temp_sssz + temp_ssz + temp_sz + temp_bz;
            dr["sssz_cjje"] = temp_sssz;
            dr["ssz_cjje"] = temp_ssz;
            dr["sz_cjje"] = temp_sz;
            dr["bz_cjje"] = temp_bz;
            dt.Rows.Add(dr);
            return dt;
        }


        public DataTable GET_JPXM_ZT_CJJE(DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {
                var temp_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => item.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
                var temp_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => item.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
                var temp_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => item.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
                var temp_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => item.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
                DataRow dr = dt.NewRow();
                dr["kfs"] = string.Join(",", item.kfs);
                dr["hj"] = temp_sssz + temp_ssz + temp_sz + temp_bz;
                dr["sssz_cjje"] = temp_sssz;
                dr["ssz_cjje"] = temp_ssz;
                dr["sz_cjje"] = temp_sz;
                dr["bz_cjje"] = temp_bz;
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public DataTable GET_JPXM_XF_CJJE(DataTable dt, JP_JPXM_INFO jpxm)
        {
            string sql = "zc >=" + (Base_date.bz - 3) + " and zc<=" + Base_date.bz;
            var query = (from t in Cache_data_rgsj.jbz.Select(sql).AsEnumerable()
                         where jpxm.kfs.Contains(t["qymc"])
                         group t by new { xm = t["xm"], yt = t["yt"] } into m
                         select new
                         {
                             xm = m.Key.xm + "(" + m.Key.yt + ")",
                             hj = m.Sum(n => n["rgje"].longs()),
                             sssz = m.Where(a => a["zc"].ints() == (Base_date.bz - 3)).Sum(n => n["rgje"].longs()),
                             ssz = m.Where(a => a["zc"].ints() == (Base_date.bz - 2)).Sum(n => n["rgje"].longs()),
                             sz = m.Where(a => a["zc"].ints() == (Base_date.bz - 1)).Sum(n => n["rgje"].longs()),
                             bz = m.Where(a => a["zc"].ints() == Base_date.bz).Sum(n => n["rgje"].longs()),
                         }).ToList();
            foreach (var item in query)
            {
                DataRow dr = dt.NewRow();
                dr["kfs"] = item.xm;
                dr["hj"] = item.hj;
                dr["sssz"] = item.sssz;
                dr["ssz"] = item.ssz;
                dr["sz"] = item.sz;
                dr["bz"] = item.bz;
                dt.Rows.Add(dr);
            }

            return dt;
        }


        public ISlideCollection _plus_jp_dyt_jzgj(ISlide sld, JP_BA_INFO item,string pagenow)
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
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, pagenow, item.ytcs[0]);
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
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, pagenow, item.ytcs[0]);
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
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text,pagenow, item.ytcs[0]);
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

        public ISlideCollection _plus_jp_dyt_jzgj_taonei(ISlide sld, JP_BA_INFO item, string pagenow)
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
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, pagenow, item.ytcs[0]);
                    //数据
                    System.Data.DataTable jzgjt = new System.Data.DataTable();
                    jzgjt.Columns.Add("");
                    jzgjt.Columns.Add("成交套数", typeof(int));
                    jzgjt.Columns.Add("套内均价", typeof(double));
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
                    Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                    t.AddClone(page);

                }
                #endregion

                #region 别墅


                else if (item.ytcs[0] == "别墅")
                {
                    IAutoShape text1 = (IAutoShape)page.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, pagenow, item.ytcs[0]);
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
                    Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                    t.AddClone(page);


                }


                #endregion

                #region 大业态


                else
                {
                    IAutoShape text1 = (IAutoShape)page.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, pagenow, item.ytcs[0]);
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
    }
}
