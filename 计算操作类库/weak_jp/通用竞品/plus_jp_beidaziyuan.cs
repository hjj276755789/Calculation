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
    /// 北大资源
    /// </summary>
    public class plus_jp_beidaziyuan :plus_jp_base
    {
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

                    //foreach (var page1 in _plus_jp_dyt_jzgj(item))
                    //{
                    //    t.AddClone(page1);
                    //}

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

                    dt.Columns.Add("当月累计认购");
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
                        if(item.zlmjqj==null||item.zlmjqj.Length<=0)
                        { 
                            Base_Log.Log("主力面积区间为空！跳过！竞品项目ID："+item.id);
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
                                dt_bz = Cache_data_cjjl.bz.Select("jzmj >=" + p[0] + " and jzmj<=" + p[1]).CopyToDataTable();
                                dt_sz = Cache_data_cjjl.sz.Select("jzmj >=" + p[0] + " and jzmj<=" + p[1]).CopyToDataTable();
                            }
                            else
                            {
                                dt_bz = Cache_data_cjjl.bz.Select("jzmj >=" + p[0]).CopyToDataTable();
                                dt_sz = Cache_data_cjjl.sz.Select("jzmj >=" + p[0]).CopyToDataTable();
                            }
                            DataRow dr1 = dt.NewRow();
                            JP_JPXM_INFO jp = new JP_JPXM_INFO();
                            jp = item;
                            jp.zlmjqj = new string[] { mjitem };
                            

                            #endregion

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_cjba_bz = dt_bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_cjba_sz = dt_sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                            var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, jp));
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
                                dt_bz = Cache_data_cjjl.bz.Select("jzmj >=" + p[0] + " and jzmj<=" + p[1]).CopyToDataTable();
                                dt_sz = Cache_data_cjjl.sz.Select("jzmj >=" + p[0] + " and jzmj<=" + p[1]).CopyToDataTable();
                            }
                            else
                            {
                                dt_bz = Cache_data_cjjl.bz.Select("jzmj >=" + p[0]).CopyToDataTable();
                                dt_sz = Cache_data_cjjl.sz.Select("jzmj >=" + p[0]).CopyToDataTable();
                            }
                            DataRow dr1 = dt.NewRow();
                            JP_JPXM_INFO jp = new JP_JPXM_INFO();
                            jp = item;
                            jp.zlmjqj = new string[] { mjitem };


                            #endregion

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_cjba_bz = dt_bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);

                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_cjba_sz = dt_sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                            //本周本案认购数据
                            var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.hxcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, jp));
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
                            dt_bz = Cache_data_cjjl.bz.Select("jzmj >=" + p[0] + " and jzmj<=" + p[1]).CopyToDataTable();
                            dt_sz = Cache_data_cjjl.sz.Select("jzmj >=" + p[0] + " and jzmj<=" + p[1]).CopyToDataTable();
                        }
                        else
                        {
                            dt_bz = Cache_data_cjjl.bz.Select("jzmj >=" + p[0]).CopyToDataTable();
                            dt_sz = Cache_data_cjjl.sz.Select("jzmj >=" + p[0]).CopyToDataTable();
                        }
                        DataRow dr1 = dt.NewRow();
                        JP_JPXM_INFO jp = new JP_JPXM_INFO();
                        jp = item;
                        jp.zlmjqj = new string[] { mjitem };


                        #endregion

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_cjba_bz = dt_bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_cjba_sz = dt_sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, jp));
                    }
                }

            }
            return dt;
        }
    }
}
