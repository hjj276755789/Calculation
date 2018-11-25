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
    /// 中交漫山竞品
    /// </summary>
    public class plus_jp_zhongjiao :plus_jp_base
    {
        public ISlideCollection _plus_jp_zhongjiaomanshan_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);

                foreach (var item in param)
                {
                    if (item.qtcs != null && item.qtcs.Length > 0)
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page1 = temp[0];
                        IAutoShape text0_1 = (IAutoShape)page1.Shapes[0];
                        text0_1.TextFrame.Text = string.Format(text0_1.TextFrame.Text, item.bamc);

                        DataTable dt1_0 = new DataTable();
                        dt1_0.Columns.Add(Base_Config_Jzgj.业态);
                        dt1_0.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt1_0.Columns.Add(Base_Config_Cjba.上周_建筑面积);
                        dt1_0.Columns.Add(Base_Config_Cjba.上周_套内面积);
                        dt1_0.Columns.Add(Base_Config_Cjba.上周_套内均价);
                        dt1_0.Columns.Add(Base_Config_Cjba.上周_成交金额);

                        dt1_0.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt1_0.Columns.Add(Base_Config_Cjba.本周_建筑面积);
                        dt1_0.Columns.Add(Base_Config_Cjba.本周_套内面积);
                        dt1_0.Columns.Add(Base_Config_Cjba.本周_套内均价);
                        dt1_0.Columns.Add(Base_Config_Cjba.本周_成交金额);

                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt1_0 = GET_JPXM_ZT(dt1_0, item.jpxmlb);
                            Office_Tables.SetJP_ZHONGJIAOMANSHAN_Table(page1, dt1_0, 1, null, null);
                            t.AddClone(page1);
                        }
                    }
                    else
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page1 = temp[2];

                        IAutoShape text0_1 = (IAutoShape)page1.Shapes[0];
                        text0_1.TextFrame.Text = string.Format(text0_1.TextFrame.Text, item.bamc);
                        DataTable dt2_0 = new DataTable();
                        dt2_0.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt2_0.Columns.Add(Base_Config_Jzgj.业态);

                        dt2_0.Columns.Add(Base_Config_Rgsj.上周_本周到访量);
                        dt2_0.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt2_0.Columns.Add(Base_Config_Cjba.上周_套内均价);
                        dt2_0.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                        dt2_0.Columns.Add(Base_Config_Rgsj.上周_认购套内均价);

                        dt2_0.Columns.Add(Base_Config_Rgsj.本周_本周到访量);
                        dt2_0.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt2_0.Columns.Add(Base_Config_Cjba.本周_套内均价);
                        dt2_0.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt2_0.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt2_0 = GET_JPXM_BX(dt2_0, item.jpxmlb);
                            Office_Tables.SetJP_ZHONGJIAOMANSHAN_1_Table(page1, dt2_0, 2, null, null);
                        }
                        //foreach (var item in item.jpxmlb)
                        //{
                        DataTable dt2_1 = new DataTable();
                        dt2_1.Columns.Add(Base_Config_Jzgj.业态);
                        dt2_1.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间);
                        dt2_1.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt2_1.Columns.Add(Base_Config_Cjba.本周_套内面积);
                        dt2_1.Columns.Add(Base_Config_Cjba.本周_成交金额);
                        dt2_1.Columns.Add(Base_Config_Cjba.本周_套内均价);
                        ;
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt2_1 = GET_JPXM_BX_1(dt2_1, item.jpxmlb);
                            Office_Tables.SetJP_ZHONGJIAOMANSHAN_2_Table(page1, dt2_1, 1, null, null);
                        }
                        t.AddClone(page1);

                        //}
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


        public ISlideCollection _plus_jp_zhongjiaogongyuan_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);

                foreach (var item in param)
                {
                    if (item.qtcs != null && item.qtcs.Length > 0)
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page1 = temp[0];
                        IAutoShape text0_1 = (IAutoShape)page1.Shapes[1];
                        text0_1.TextFrame.Text = string.Format(text0_1.TextFrame.Text, item.bamc);

                        DataTable dt1_0 = new DataTable();

                        var jbz_cjba_spf = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                                            where item.ztcs.Contains(a["zt"])
                                            group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                                            select new
                                            {
                                                zc = s.Key.zc,
                                                zcmc = s.Key.zcmc,
                                                cjje = s.Sum(a => a["cjje"].longs()),
                                                jzmj = s.Sum(a => a["jzmj"].doubls()),
                                            }).OrderBy(m => m.zc).ToList();
                        var jbz_xzys_spf = (from a in Cache_data_xzys.jbz.AsEnumerable()
                                            where item.ztcs.Contains(a["zt"])
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
                        DataTable dt0_1 = new DataTable();
                        dt0_1.Columns.Add("周次");
                        dt0_1.Columns.Add("供应量（万㎡）");
                        dt0_1.Columns.Add("成交量（万㎡）");
                        dt0_1.Columns.Add("建面均价（元/㎡）");
                        foreach (var itemspf in temp_spf)
                        {
                            DataRow dr = dt0_1.NewRow();
                            dr[0] = itemspf.zcmc;
                            dr[1] = itemspf.xzgyl.mj_wf();
                            dr[2] = itemspf.cjmj.mj_wf();
                            dr[3] = itemspf.jmjj.je_y();
                            dt0_1.Rows.Add(dr);
                                
                        }
                        Office_Charts.Chart_gxfx(page1, dt0_1, 2);
                        t.AddClone(page1);
                    }
                    else
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page1 = temp[2];

                        IAutoShape text0_1 = (IAutoShape)page1.Shapes[0];
                        text0_1.TextFrame.Text = string.Format(text0_1.TextFrame.Text, item.bamc);
                        DataTable dt2_0 = new DataTable();
                        dt2_0.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt2_0.Columns.Add(Base_Config_Jzgj.业态);

                        dt2_0.Columns.Add(Base_Config_Rgsj.上周_本周到访量);
                        dt2_0.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt2_0.Columns.Add(Base_Config_Cjba.上周_套内均价);
                        dt2_0.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                        dt2_0.Columns.Add(Base_Config_Rgsj.上周_认购套内均价);

                        dt2_0.Columns.Add(Base_Config_Rgsj.本周_本周到访量);
                        dt2_0.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt2_0.Columns.Add(Base_Config_Cjba.本周_套内均价);
                        dt2_0.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt2_0.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt2_0 = GET_JPXM_BX(dt2_0, item.jpxmlb);
                            Office_Tables.SetJP_ZHONGJIAOMANSHAN_1_Table(page1, dt2_0, 2, null, null);
                        }
                        //foreach (var item in item.jpxmlb)
                        //{
                        DataTable dt2_1 = new DataTable();
                        dt2_1.Columns.Add(Base_Config_Jzgj.业态);
                        dt2_1.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间);
                        dt2_1.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt2_1.Columns.Add(Base_Config_Cjba.本周_套内面积);
                        dt2_1.Columns.Add(Base_Config_Cjba.本周_成交金额);
                        dt2_1.Columns.Add(Base_Config_Cjba.本周_套内均价);
                        ;
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt2_1 = GET_JPXM_BX_1(dt2_1, item.jpxmlb);
                            Office_Tables.SetJP_ZHONGJIAOMANSHAN_2_Table(page1, dt2_1, 1, null, null);
                        }
                        t.AddClone(page1);

                        //}
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
        /// 主题表现 备案数据
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="jpxm"></param>
        /// <returns></returns>
        public System.Data.DataTable GET_JPXM_ZT(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
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
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m =>m["xfyt"].ToString() == item.xfytcs[i]&&item.ztcs.Contains(m["zt"]));
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m =>m["xfyt"].ToString() == item.xfytcs[i] && item.ztcs.Contains(m["zt"]));
                            //本周本案认购数据
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt,temp_cjba_bz, temp_cjba_sz, item));

                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态  
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m =>  m["xfyt"].ToString() == item.ytcs[0] && item.ztcs.Contains(m["zt"]));
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m =>  m["xfyt"].ToString() == item.ytcs[0] && item.ztcs.Contains(m["zt"]));
                        //本周本案认购数据
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else if (item.ytcs[0] == "商务")
                {

                    //商务业态：如果户型参数为空，则使用细分业态参数，若细分业态参数业为空，直接使用主业态：商务
                    if (item.hxcs != null && item.hxcs.Length > 0)
                    {
                        for (int i = 0; i < item.hxcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            //var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["xfyt"].ToString() == item.hxcs[0] && item.ztcs.Contains(m["zt"]));

                           // var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["xfyt"].ToString() == item.hxcs[0] && item.ztcs.Contains(m["zt"]));

                            #endregion

                            dt.Rows.Add(GET_ROW(item.hxcs[i], dr1, dt,  temp_cjba_bz, temp_cjba_sz, item));
                        }
                    }
                    else if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态

                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["xfyt"].ToString() == item.xfytcs[0] && item.ztcs.Contains(m["zt"]));
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["xfyt"].ToString() == item.xfytcs[0] && item.ztcs.Contains(m["zt"]));

                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_cjba_bz, temp_cjba_sz, item));
                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态

                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["xfyt"].ToString() == item.ytcs[0] && item.ztcs.Contains(m["zt"]));
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["xfyt"].ToString() == item.ytcs[0] && item.ztcs.Contains(m["zt"]));

                      

                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else if (item.ytcs[0] == "商业")
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m =>  m["xfyt"].ToString() == item.ytcs[0] && item.ztcs.Contains(m["zt"]));
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m =>  m["xfyt"].ToString() == item.ytcs[0] && item.ztcs.Contains(m["zt"]));


                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt,  temp_cjba_bz, temp_cjba_sz, item));
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                  
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m =>  item.ytcs.Contains(m["yt"].ToString()) && item.ztcs.Contains(m["zt"]));

                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m =>  item.ytcs.Contains(m["yt"].ToString()) && item.ztcs.Contains(m["zt"]));
                  
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_cjba_bz, temp_cjba_sz, item));


                }
            }


            return dt;
        }

        /// <summary>
        /// 竞品项目表现 本案与认购数据
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
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[0]);
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[0]);

                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[0]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[0]);
                            //本周本案认购数据
                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
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
                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else if (item.ytcs[0] == "商务")
                {

                    //商务业态：如果户型参数为空，则使用细分业态参数，若细分业态参数业为空，直接使用主业态：商务
                    if (item.hxcs != null && item.hxcs.Length > 0)
                    {
                        for (int i = 0; i < item.hxcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[0]);

                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[0]);


                            //本周本案认购数据
                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.hxcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                        }
                    }
                    else if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态

                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[0]);
                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[0]);

                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[0]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[0]);


                            //本周本案认购数据
                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);


                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);

                        //本周本案认购数据
                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else if (item.ytcs[0] == "商业")
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态

                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);


                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);

                    //本周本案认购数据
                    var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));

                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    //本周本案认购数据
                    var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));


                }
            }


            return dt;
        }
        public System.Data.DataTable GET_JPBA_BX(System.Data.DataTable dt, JP_BA_INFO item)
        {
           
                if (item.ytcs[0] == "别墅")
                {
                    if (item.xfytcs != null&&item.xfytcs.Contains("别墅"))
                    { 
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {

                            DataRow dr1 = dt.NewRow();
                            #region 数据准备
                            //竞品业态
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                            #endregion

                            dt.Rows.Add(GET_ROW_BA(item.xfytcs[i], dr1, dt, null, null, temp_cjba_bz, temp_cjba_sz, item));

                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态  
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                        //本周本案认购数据
                        #endregion

                        dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, null, null, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else if (item.ytcs[0] == "商务")
                {

                    //商务业态：如果户型参数为空，则使用细分业态参数，若细分业态参数业为空，直接使用主业态：商务
                    if (item.hxcs != null && item.hxcs.Length > 0)
                    {
                        for (int i = 0; i < item.hxcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[0]);

                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[0]);


                            //本周本案认购数据
                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW_BA(item.hxcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                        }
                    }
                    else if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态

                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[0]);
                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[0]);

                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[0]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[0]);


                            //本周本案认购数据
                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                            #endregion

                            dt.Rows.Add(GET_ROW_BA(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);


                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);

                        //本周本案认购数据
                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                        #endregion

                        dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else if (item.ytcs[0] == "商业")
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态

                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);


                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);

                    //本周本案认购数据
                    var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                    #endregion

                    dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
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



        public System.Data.DataTable GET_JPXM_BX_1(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {
                if (item.ytcs == null || item.ytcs.Length <= 0)
                {
                    Base_Log.Log("业态参数为空！跳过！竞品项目ID：" + item.id);
                    continue;
                }
                if (item.ytcs[0] == "别墅")
                {
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {
                        //这里根据要求来设置（若不需要计算面积区间，这里需要注释）
                        if (item.zlmjqj == null || item.zlmjqj.Length <= 0)
                        {
                            Base_Log.Log("主力面积区间为空！跳过！竞品项目ID：" + item.id);
                            continue;
                        }
                        foreach (var mjitem in item.zlmjqj)
                        {
                            #region 计算主力面积区间
                            string[] p = mjitem.Split('-');
                            DataTable dt_bz = new DataTable();
                            DataTable dt_sz = new DataTable();
                            if (p[1] != "∞")
                            {
                                dt_bz = Cache_data_cjjl.bz.Select("tnmj >=" + p[0] + " and tnmj<=" + p[1]).CopyToDataTable();
                                dt_sz = Cache_data_cjjl.sz.Select("tnmj >=" + p[0] + " and tnmj<=" + p[1]).CopyToDataTable();
                            }
                            else
                            {
                                dt_bz = Cache_data_cjjl.bz.Select("tnmj >=" + p[0]).CopyToDataTable();
                                dt_sz = Cache_data_cjjl.sz.Select("tnmj >=" + p[0]).CopyToDataTable();
                            }
                            DataRow dr1 = dt.NewRow();
                            JP_JPXM_INFO jp = new JP_JPXM_INFO();
                            jp = item;
                            jp.zlmjqj = new string[] { mjitem };


                            #endregion

                            #region 数据准备
                            //竞品业态

                            var temp_cjba_bz = dt_bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_cjba_bz, jp));
                        }

                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    for (int i = 0; i < item.hxcs.Length; i++)
                    {
                        //这里根据要求来设置（若不需要计算面积区间，这里需要注释）
                        if (item.zlmjqj == null || item.zlmjqj.Length <= 0)
                        {
                            Base_Log.Log("主力面积区间为空！跳过！竞品项目ID：" + item.id);
                            continue;
                        }
                        foreach (var mjitem in item.zlmjqj)
                        {
                            #region 计算主力面积区间
                            string[] p = mjitem.Split('-');
                            DataTable dt_bz = new DataTable();
                            DataTable dt_sz = new DataTable();
                            if (p[1] != "∞")
                            {
                                dt_bz = Cache_data_cjjl.bz.Select("tnmj >=" + p[0] + " and tnmj<=" + p[1]).CopyToDataTable();
                                dt_sz = Cache_data_cjjl.sz.Select("tnmj >=" + p[0] + " and tnmj<=" + p[1]).CopyToDataTable();
                            }
                            else
                            {
                                dt_bz = Cache_data_cjjl.bz.Select("tnmj >=" + p[0]).CopyToDataTable();
                                dt_sz = Cache_data_cjjl.sz.Select("tnmj >=" + p[0]).CopyToDataTable();
                            }
                            DataRow dr1 = dt.NewRow();
                            JP_JPXM_INFO jp = new JP_JPXM_INFO();
                            jp = item;
                            jp.zlmjqj = new string[] { mjitem };


                            #endregion

                            #region 数据准备
                            //竞品业态

                            var temp_cjba_bz = dt_bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);

                          
                            #endregion

                            dt.Rows.Add(GET_ROW(item.hxcs[0], dr1, dt,  temp_cjba_bz,  jp));
                        }
                    }
                }
                else
                {

                    //这里根据要求来设置（若不需要计算面积区间，这里需要注释）
                    if (item.zlmjqj == null || item.zlmjqj.Length <= 0)
                    {
                        Base_Log.Log("主力面积区间为空！跳过！竞品项目ID：" + item.id);
                        continue;
                    }
                    foreach (var mjitem in item.zlmjqj)
                    {
                        #region 计算主力面积区间
                        string[] p = mjitem.Split('-');
                        DataTable dt_bz = new DataTable();
                        DataTable dt_sz = new DataTable();
                        if (p[1] != "∞")
                        {
                            dt_bz = Cache_data_cjjl.bz.Select("tnmj >=" + p[0] + " and tnmj<=" + p[1]).CopyToDataTable();
                            dt_sz = Cache_data_cjjl.sz.Select("tnmj >=" + p[0] + " and tnmj<=" + p[1]).CopyToDataTable();
                        }
                        else
                        {
                            dt_bz = Cache_data_cjjl.bz.Select("tnmj >=" + p[0]).CopyToDataTable();
                            dt_sz = Cache_data_cjjl.sz.Select("tnmj >=" + p[0]).CopyToDataTable();
                        }
                        DataRow dr1 = dt.NewRow();
                        JP_JPXM_INFO jp = new JP_JPXM_INFO();
                        jp = item;
                        jp.zlmjqj = new string[] { mjitem };


                        #endregion

                        #region 数据准备
                        //竞品业态
                        var temp_cjba_bz = dt_bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_cjba_bz, jp));
                    }
                }

            }
            return dt;
        }

        public  DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                             EnumerableRowCollection<DataRow> temp_cjba_bz,
                             EnumerableRowCollection<DataRow> temp_cjba_sz,
                             JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {

                try
                {
                    if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Cjba.本周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.本周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()).je_wy() : 0; }; break;
                            case Base_Config_Cjba.本周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()).mj_wf() : 0; }; break;
                            case Base_Config_Cjba.本周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls()).mj_wf() : 0; }; break;
                            case Base_Config_Cjba.本周_建面均价:
                                {

                                    if ((temp_cjba_bz != null && temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Cjba.本周_套内均价:
                                {
                                    if ((temp_cjba_bz != null && temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Cjba.上周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.上周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()).je_wy() : 0; }; break;
                            case Base_Config_Cjba.上周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()).mj_wf() : 0; }; break;
                            case Base_Config_Cjba.上周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls()).mj_wf() : 0; }; break;
                            case Base_Config_Cjba.上周_建面均价:
                                {
                                    if ((temp_cjba_sz != null && temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Cjba.上周_套内均价:
                                {
                                    if ((temp_cjba_sz != null && temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;


                            case Base_Config_Cjba.本周_备案套数环比:
                                {
                                    dr1[dt.Columns[j].ColumnName] = ((temp_cjba_bz.Sum(m => m["ts"].ints()) - temp_cjba_sz.Sum(m => m["ts"].ints())) / temp_cjba_sz.Sum(m => m["ts"].ints())).doubls().ss_bfb();
                                }; break;
                            case Base_Config_Cjba.本周_套内均价环比:
                                {
                                    long bz_cjje = temp_cjba_bz.Sum(m => m["cjje"].ints());
                                    long bz_tnmj = temp_cjba_bz.Sum(m => m["tnmj"].ints());
                                    long sz_cjje = temp_cjba_sz.Sum(m => m["cjje"].ints());
                                    long sz_tnmj = temp_cjba_sz.Sum(m => m["tnmj"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_cjje / bz_tnmj - sz_cjje / sz_tnmj) / (sz_cjje / sz_tnmj)).doubls().ss_bfb();
                                }; break;
                            case Base_Config_Cjba.本周_套内面积环比:
                                {
                                    long bz_tnmj = temp_cjba_bz.Sum(m => m["tnmj"].ints());
                                    long sz_tnmj = temp_cjba_sz.Sum(m => m["tnmj"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_tnmj - sz_tnmj) / (sz_tnmj)).doubls().ss_bfb();
                                }; break;
                            case Base_Config_Cjba.本周_套均总价环比:
                                {
                                    long bz_cjje = temp_cjba_bz.Sum(m => m["cjje"].ints());
                                    long bz_ts = temp_cjba_bz.Sum(m => m["ts"].ints());
                                    long sz_cjje = temp_cjba_sz.Sum(m => m["cjje"].ints());
                                    long sz_ts = temp_cjba_sz.Sum(m => m["ts"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_cjje / bz_ts - sz_cjje / sz_ts) / (sz_cjje / sz_ts)).doubls().ss_bfb();
                                }; break;

                            case Base_Config_Cjba.本周_建筑面积环比:
                                {
                                    long bz_cjje = temp_cjba_bz.Sum(m => m["cjje"].ints());
                                    long bz_jzmj = temp_cjba_bz.Sum(m => m["jzmj"].ints());
                                    long sz_cjje = temp_cjba_sz.Sum(m => m["cjje"].ints());
                                    long sz_jzmj = temp_cjba_sz.Sum(m => m["jzmj"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_cjje / bz_jzmj - sz_cjje / sz_jzmj) / (sz_cjje / sz_jzmj)).doubls().ss_bfb();
                                }; break;
                            case Base_Config_Cjba.本周_建面均价环比:
                                {
                                    long bz_cjje = temp_cjba_bz.Sum(m => m["cjje"].ints());
                                    long bz_jzmj = temp_cjba_bz.Sum(m => m["jzmj"].ints());
                                    long sz_cjje = temp_cjba_sz.Sum(m => m["cjje"].ints());
                                    long sz_jzmj = temp_cjba_sz.Sum(m => m["jzmj"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_cjje / bz_jzmj - sz_cjje / sz_jzmj) / (sz_cjje / sz_jzmj)).doubls().ss_bfb();
                                }; break;
                            case Base_Config_Cjba.本周_成交金额环比:
                                {
                                    long bz_cjje = temp_cjba_bz.Sum(m => m["cjje"].ints());
                                    long sz_cjje = temp_cjba_sz.Sum(m => m["cjje"].ints());
                                    dr1[dt.Columns[j].ColumnName] = ((bz_cjje - sz_cjje) / sz_cjje).doubls().ss_bfb();
                                }; break;



                            default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                        }


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
                                    dr1[dt.Columns[j].ColumnName] = item != null && item.ztcs != null ? string.Join(",", item.ztcs) : "";
                                }; break;
                            case Base_Config_Jzgj.竞争格局_主力面积区间:
                                {
                                    dr1[dt.Columns[j].ColumnName] = string.Join(",", item.zlmjqj);
                                }; break;
                            case Base_Config_Jzgj.竞争格局名称:
                                {
                                    dr1[dt.Columns[j].ColumnName] = item.jzgjmc;
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

        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                          EnumerableRowCollection<DataRow> temp_cjba_bz,
                          JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {

                try
                {
                    if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Cjba.本周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.本周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()).je_wy() : 0; }; break;
                            case Base_Config_Cjba.本周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()).mj_wf() : 0; }; break;
                            case Base_Config_Cjba.本周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls()).mj() : 0; }; break;
                            case Base_Config_Cjba.本周_建面均价:
                                {

                                    if ((temp_cjba_bz != null && temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Cjba.本周_套内均价:
                                {
                                    if ((temp_cjba_bz != null && temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;

                            default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                        }


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
                                    dr1[dt.Columns[j].ColumnName] = item != null && item.ztcs != null ? string.Join(",", item.ztcs) : "";
                                }; break;
                            case Base_Config_Jzgj.竞争格局_主力面积区间:
                                {
                                    dr1[dt.Columns[j].ColumnName] = string.Join(",", item.zlmjqj);
                                }; break;
                            case Base_Config_Jzgj.竞争格局名称:
                                {
                                    dr1[dt.Columns[j].ColumnName] = item.jzgjmc;
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
