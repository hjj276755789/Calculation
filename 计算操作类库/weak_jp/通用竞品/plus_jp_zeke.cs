﻿using Aspose.Slides;
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
    class plus_jp_zeke : plus_jp_base
    {
        public ISlideCollection _plus_jp_zeke_1(string str, int cjbh)
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

                    #region 竞品分布
                    foreach (var page1 in _plus_jp_dyt_jzgj(item))
                    {
                        t.AddClone(page1);
                    }
                    #endregion
                   



                    #region 格局统计
                    if (item.ytcs == null || item.ytcs[0] != "商铺")
                    {
                        var page2 = temp[1];
                        DataTable dt = new DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.竞争格局名称);
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);

                        dt.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_新开套内均价);

                        dt.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.上周_套内面积);
                        dt.Columns.Add(Base_Config_Cjba.上周_套内均价);

                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套内体量);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套内均价);

                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.本周_套内面积);
                        dt.Columns.Add(Base_Config_Cjba.本周_套内均价);

                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套内体量);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);

                        dt.Columns.Add(Base_Config_Rgsj.本周_成交套数环比);
                        dt.Columns.Add(Base_Config_Rgsj.本周_套内均价环比);
                        dt.Columns.Add(Base_Config_Rgsj.本周_变化原因);

                        IAutoShape text2 = (IAutoShape)page2.Shapes[2];
                        text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc, item.ytcs != null ? item.ytcs[0] : "");


                        dt = GET_JPBA_BX(dt, item);
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_ZeKe_JPBX_Table(page2, dt, 0, null, null,null);
                            t.AddClone(page2);
                        }
                    }
                    else
                    {
                        var page2 = temp[2];
                        DataTable dt = new DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.竞争格局名称);
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);

                        dt.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_新开套内均价);

                        dt.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.上周_建筑面积);
                        dt.Columns.Add(Base_Config_Cjba.上周_套内均价);

                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.本周_建筑面积);
                        dt.Columns.Add(Base_Config_Cjba.本周_套内均价);

                        dt.Columns.Add(Base_Config_Rgsj.本周_成交套数环比);
                        dt.Columns.Add(Base_Config_Rgsj.本周_套内均价环比);
                        dt.Columns.Add(Base_Config_Rgsj.本周_变化原因);

                        IAutoShape text2 = (IAutoShape)page2.Shapes[2];
                        text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc, "商铺");

                        dt = GET_JPBA_BX(dt, item);

                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_ZeKe_JPBX_Table(page2, dt, 0, null, null, "商铺");
                            t.AddClone(page2);
                        }
                        #endregion
                        foreach (var page3 in _plus_jp_dyt_tgtp(item))
                        {
                            t.AddClone(page3);
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



        public System.Data.DataTable GET_JPBA_BX(System.Data.DataTable dt, JP_BA_INFO item)
        {
            if (item.ytcs[0] == "别墅")
            {
                if (item.xfytcs != null&&item.xfytcs.Length>0) { 
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

                        dt.Rows.Add(GET_ROW_BA(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));

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
    }
}
