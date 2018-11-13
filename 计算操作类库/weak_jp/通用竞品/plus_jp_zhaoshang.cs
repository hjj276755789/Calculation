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
    public class plus_jp_zhaoshang : plus_jp_base
    {
        public ISlideCollection _plus_jp_zhaoshang_1(string str, int cjbh)
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
                    int[] index1 = { 1, 0 };
                    int[] index2 = { 2, 3 };
                    int[] index3 = { 4, 2,1 };
                    int[] index4 = { 1, 0 };
                    Base_Log.Log("主团周度排名开始");
                    foreach (var qypm in ztzdpm(str, index1, index2, index3, index4, item.qycs[0]))
                    {
                        t.AddClone(qypm);
                    }
                    Base_Log.Log("主团周度排名结束");

                    #region 格局统计
                    Base_Log.Log("格局统计开始");

                    var page2 = temp[3];
                    DataTable dt = new DataTable();
                    dt.Columns.Add(Base_Config_Jzgj.竞争格局名称);
                    dt.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt.Columns.Add(Base_Config_Jzgj.业态);

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
                    IAutoShape text2 = (IAutoShape)page2.Shapes[1];
                    text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc, item.ytcs[0]);
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_BX(dt, item.jpxmlb);
                        Office_Tables.SetJP_RUIAN_JPBX_Table(page2, dt, 2, null, null);
                        t.AddClone(page2);
                    }
                    Base_Log.Log("格局统计结束");
                    Base_Log.Log("近期动作开始");

                    #endregion
                    #region 近期动作
                    var page3 = temp[4];
                    IAutoShape text3 = (IAutoShape)page3.Shapes[0];
                    text3.TextFrame.Text = string.Format(text3.TextFrame.Text, item.bamc, item.ytcs[0]);

                    DataTable dt1 = new DataTable();
                    dt1.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt1.Columns.Add(Base_Config_Rgsj.本周_优惠);
                    dt1.Columns.Add(Base_Config_Rgsj.本周_营销动作);
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt1 = GET_JPXM_BX(dt1, item.jpxmlb);
                        Office_Tables.SetTable(page3, dt1, 1, null, null);
                    }
                    t.AddClone(page3);
                    Base_Log.Log("近期动作开始");
                    #endregion



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
