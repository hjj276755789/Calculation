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
    public class plus_jp_yangguangcheng :plus_jp_base
    {
        public ISlideCollection _plus_jp_yangguangcheng_1(string str, int cjbh)
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
                    IAutoShape text0_1 = (IAutoShape)page1.Shapes[0];
                    text0_1.TextFrame.Text = string.Format(text0_1.TextFrame.Text, item.bamc);

                    DataTable dt1_0 = new DataTable();
                    dt1_0.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt1_0.Columns.Add(Base_Config_Jzgj.业态);
                    dt1_0.Columns.Add(Base_Config_Rgsj.上周_本周到访量);
                    dt1_0.Columns.Add(Base_Config_Rgsj.本周_本周到访量);
                    dt1_0.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间);
                    dt1_0.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                    dt1_0.Columns.Add(Base_Config_Cjba.本周_备案套数);
                    dt1_0.Columns.Add(Base_Config_Cjba.本周_建面均价);
                    dt1_0.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                    dt1_0.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);

                    dt1_0.Columns.Add(Base_Config_Rgsj.本周_优惠);
                    dt1_0.Columns.Add("认购库存套数");
                    dt1_0.Columns.Add("本周动态");
                    dt1_0.Columns.Add(Base_Config_Rgsj.本周_活动);
                    dt1_0.Columns.Add(Base_Config_Rgsj.本周_营销动作);
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0) { 
                        dt1_0 = GET_JPXM_BX_1(dt1_0, item.jpxmlb);
                        Office_Tables.SetJP_YANGGUANGCHENG_Table(page1, dt1_0, 1, null, null);
                        t.AddClone(page1);

                        foreach (var jpitem in item.jpxmlb)
                        {
                            var tp1 = new Presentation(str);
                            var temp1 = tp1.Slides;
                            var page2 = temp1[1];
                            IAutoShape text1_1 = (IAutoShape)page2.Shapes[0];
                            text1_1.TextFrame.Text = string.Format(text1_1.TextFrame.Text,string.Join( ",",jpitem.lpcs));
                            t.AddClone(page2);
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
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);

                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_cjba_bz, jp));
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
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);

                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();

                            #endregion

                            dt.Rows.Add(GET_ROW(item.hxcs[0], dr1, dt, temp_rg_bz, temp_cjba_bz, jp));
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
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains( m["yt"].ToString()) );

                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz,temp_cjba_bz, jp));
                    }
                }

            }
            return dt;
        }

    }
}
