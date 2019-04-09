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

    /// <summary>
    /// 北大资源
    /// </summary>
    public class plus_jp_beidaziyuan :plus_jp_base
    {
        private static DataTable byba;
        //private static DataTable bycj;
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
        public class Base_Config_Rgsj_BY
        {

            public const string 本月_认购套数 = "by_ts";
            public static string[] _认购数据 = { "by_ts", };
        }

        public plus_jp_beidaziyuan()
        {
            Base_date.init_yb(Base_date.bn, Base_date.GET_Y_FROM_Z(Base_date.bn, Base_date.bz));
            byba = ZB_Data_CJBA_DataProvider.GET_ZB(Base_date.by_First, Base_date.bz_Last);
        }

        public ISlideCollection _plus_jp_beidaziyuan_1(string str, int cjbh)
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
                    if (item.qtcs!="住宅") {
                        var page2 = temp[1];
                        #region 格局统计
                        DataTable dt = new DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);
                        dt.Columns.Add("在售楼栋");//
                        dt.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间);


                        dt.Columns.Add(Base_Config_Rgsj.上周_本周到访量);
                        dt.Columns.Add(Base_Config_Cjba.上周_备案套数);

                        dt.Columns.Add(Base_Config_Rgsj.本周_本周到访量);
                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.本周_建面均价);

                        dt.Columns.Add(Base_Config_Cjba_BY.本月_备案套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_营销动作);

                        IAutoShape text2 = (IAutoShape)page2.Shapes[0];
                        text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc);
                        #endregion
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_BEIDAZIYUAN_PT_Table(page2, dt, 2, null, null);
                            t.AddClone(page2);
                        }
                    }
                    else {
                        var page2 = temp[0];
                        #region 格局统计
                        DataTable dt = new DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);
                        dt.Columns.Add("在售楼栋");//
                        dt.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间);


                        dt.Columns.Add(Base_Config_Rgsj.上周_本周到访量);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);

                        dt.Columns.Add(Base_Config_Rgsj.本周_本周到访量);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);

                        dt.Columns.Add("当月累计认购");//认购无时间字段，无法获取本月认购。
                        dt.Columns.Add("剩余套数");
                        dt.Columns.Add(Base_Config_Rgsj.本周_营销动作);

                        IAutoShape text2 = (IAutoShape)page2.Shapes[0];
                        text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc);
                        #endregion
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_BEIDAZIYUAN_PT_Table(page2, dt, 2, null, null);
                            t.AddClone(page2);
                        }
                    }
                    foreach (var page3 in _plus_jp_dyt_tgtp(item))
                    {
                        t.AddClone(page3);
                    }

                }
                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log(e.Message);
                Base_Log.Log(e.StackTrace);
                return null;
            }
        }

        public System.Data.DataTable GET_JPXM_BX(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {
                if (item.ytcs == null || item.ytcs.Length <= 0) {
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

                            DataRow dr1 = dt.NewRow();
                            JP_JPXM_INFO jp = new JP_JPXM_INFO();
                            jp = item;
                            jp.zlmjqj = new string[] { mjitem };


                            #endregion

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                            var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(string.Join(",", item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, null, null, jp));
                        }

                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    if (item.hxcs != null && item.hxcs.Length > 0)
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
                                DataRow dr1 = dt.NewRow();
                                JP_JPXM_INFO jp = new JP_JPXM_INFO();
                                jp = item;
                                jp.zlmjqj = new string[] { mjitem };
                               

                                #endregion

                                #region 数据准备
                                //竞品业态
                                var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                                var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);

                                var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                                var temp_cjba_sz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                                var tempby = byba.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                                //本周本案认购数据
                                var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                                var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                                #endregion

                                dt.Rows.Add(GET_ROW(item.hxcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, tempby, jp));
                            }
                        }
                    }
                    else if (item.xfytcs != null && item.xfytcs.Length > 0)
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
                                DataRow dr1 = dt.NewRow();
                                JP_JPXM_INFO jp = new JP_JPXM_INFO();
                                jp = item;
                                jp.zlmjqj = new string[] { mjitem };


                                #endregion

                                #region 数据准备
                                //竞品业态
                                var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                                var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]) ;

                                var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                                var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]) ;
                                var tempby = byba.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                                //本周本案认购数据
                                var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                                var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                                #endregion

                                dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, tempby,jp));
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
                            DataRow dr1 = dt.NewRow();
                            JP_JPXM_INFO jp = new JP_JPXM_INFO();
                            jp = item;
                            jp.zlmjqj = new string[] { mjitem };


                            #endregion

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));

                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                            var tempby = byba.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains( m["yt"].ToString()));
                            //本周本案认购数据
                            var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(string.Join(",", item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, tempby, jp));
                        }
                    }

                }
                else if (item.ytcs[0] == "商铺")
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
                        
                        DataRow dr1 = dt.NewRow();
                        JP_JPXM_INFO jp = new JP_JPXM_INFO();
                        jp = item;
                        jp.zlmjqj = new string[] { mjitem };


                        #endregion

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString())) ;

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString())) ;
                        var tempby = byba.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(string.Join(",", item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, tempby, jp));
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
                        
                        DataRow dr1 = dt.NewRow();
                        JP_JPXM_INFO jp = new JP_JPXM_INFO();
                        jp = item;
                        jp.zlmjqj = new string[] { mjitem };


                        #endregion

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(string.Join(",", item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, null, null, jp));
                    }
                }

            }
            return dt;
        }



        public virtual DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                        DataRow temp_ba_bz,
                        DataRow temp_ba_sz,
                        EnumerableRowCollection<DataRow> temp_cjba_bz,
                        EnumerableRowCollection<DataRow> temp_cjba_sz,
                        EnumerableRowCollection<DataRow> temp_cjba_by,
                        JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {

                try
                {


                    if (Base_Config_Rgsj._认购数据.Contains(dt.Columns[j].ColumnName))
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Rgsj.上周_新开套数:
                            case Base_Config_Rgsj.上周_新开销售套数:
                            case Base_Config_Rgsj.上周_新开建面均价:
                            case Base_Config_Rgsj.上周_新开套内均价:
                            case Base_Config_Rgsj.上周_认购套数:
                            case Base_Config_Rgsj.上周_认购套内体量:
                            case Base_Config_Rgsj.上周_认购套内均价:
                            case Base_Config_Rgsj.上周_认购建面体量:
                            case Base_Config_Rgsj.上周_认购建面均价:
                            case Base_Config_Rgsj.上周_认购金额:
                            case Base_Config_Rgsj.上周_本周到访量:
                                {
                                    if (temp_ba_sz != null)
                                    {
                                        dr1[dt.Columns[j].ColumnName] = temp_ba_sz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                    }
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Rgsj.本周_新开套数:
                            case Base_Config_Rgsj.本周_新开销售套数:
                            case Base_Config_Rgsj.本周_新开建面均价:
                            case Base_Config_Rgsj.本周_新开套内均价:
                            case Base_Config_Rgsj.本周_认购套数:
                            case Base_Config_Rgsj.本周_认购套内体量:
                            case Base_Config_Rgsj.本周_认购套内均价:
                            case Base_Config_Rgsj.本周_认购建面体量:
                            case Base_Config_Rgsj.本周_认购建面均价:
                            case Base_Config_Rgsj.本周_认购金额:
                            case Base_Config_Rgsj.本周_本周到访量:

                                {
                                    if (temp_ba_bz != null)
                                    {
                                        dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                    }
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Rgsj.本周_认购套数环比: { dr1[dt.Columns[j].ColumnName] = temp_ba_sz["rgts"] != null && temp_ba_sz["rgts"].ints() != 0 ? ((temp_ba_bz["rgts"].ints() - temp_ba_sz["rgts"].ints()) / temp_ba_sz["rgts"].ints()).doubls().ss_bfb() : "0%"; }; break;
                            case Base_Config_Rgsj.本周_认购金额环比: { dr1[dt.Columns[j].ColumnName] = temp_ba_sz["rgje"] != null && temp_ba_sz["rgje"].ints() != 0 ? ((temp_ba_bz["rgts"].ints() - temp_ba_sz["rgts"].ints()) / temp_ba_sz["rgts"].ints()).doubls().ss_bfb() : "0%"; }; break;
                            case Base_Config_Rgsj.本周_认购建筑面积环比: { dr1[dt.Columns[j].ColumnName] = temp_ba_sz["rgjmtl"] != null && temp_ba_sz["rgjmtl"].ints() != 0 ? ((temp_ba_bz["rgjmtl"].ints() - temp_ba_sz["rgjmtl"].ints()) / temp_ba_sz["rgjmtl"].ints()).doubls().ss_bfb() : "0%"; }; break;
                            case Base_Config_Rgsj.本周_认购套内面积环比: { }; break;
                            case Base_Config_Rgsj.本周_认购建面均价环比: { dr1[dt.Columns[j].ColumnName] = temp_ba_sz["rgjmjj"] != null && temp_ba_sz["rgjmjj"].ints() != 0 ? ((temp_ba_bz["rgjmjj"].ints() - temp_ba_sz["rgjmjj"].ints()) / temp_ba_sz["rgjmjj"].ints()).doubls().ss_bfb() : "0%"; }; break;
                            case Base_Config_Rgsj.本周_认购套内均价环比: { }; break;
                            case Base_Config_Rgsj.本周_认购套均总价环比: { }; break;
                            default:
                                {
                                    try
                                    {
                                        if (temp_ba_bz != null)
                                        {
                                            dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName];
                                        }
                                        else
                                            dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                    catch (Exception)
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "-";
                                    }

                                }; break;

                        }
                    }

                    else if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Cjba.本周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.本周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) : 0; }; break;
                            case Base_Config_Cjba.本周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()) : 0; }; break;
                            case Base_Config_Cjba.本周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].ints()) : 0; }; break;
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
                            case Base_Config_Cjba.上周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) : 0; }; break;
                            case Base_Config_Cjba.上周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()) : 0; }; break;
                            case Base_Config_Cjba.上周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].ints()) : 0; }; break;
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
                    else if (Base_Config_Cjba_BY._备案数据.Contains(dt.Columns[j].ColumnName))
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Cjba_BY.本月_备案套数:
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_cjba_by.Sum(m => m["ts"].ints());
                                };break;

                        }
                    }
                    else if (Base_Config_Rgsj_BY._认购数据.Contains(dt.Columns[j].ColumnName))
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Rgsj_BY.本月_认购套数:
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_cjba_by.Sum(m => m["rgts"].ints());
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
