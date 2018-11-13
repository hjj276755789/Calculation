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
    public class plus_jp_xuhuicheng : plus_jp_base_jpmb1
    {
        public ISlideCollection _plus_jp_xuhuicheng_1(string str, int cjbh)
        {
            try
            {

                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);
                var temp1_1 = new Presentation(str).Slides;
                var page1_1 = temp1_1[0];
                string[] zt1_1 = { "大竹林", "照母山", "礼嘉" };
                DataTable dt1_1 = JSZZTSCBX(zt1_1);
                Office_Charts.Chart_gxfx(page1_1, dt1_1, 1);
                t.AddClone(page1_1);
                var temp1_2 = new Presentation(str).Slides;
                var page1_2 = temp1_2[0];
                string[] zt1_2 = { "巴南区" };
                DataTable dt1_2 = JSZQYSCBX(zt1_2);
                Office_Charts.Chart_gxfx(page1_2, dt1_2, 1);
                t.AddClone(page1_2);

                foreach (var item in param)
                {

                    if (string.IsNullOrEmpty(item.qtcs)) {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        #region 持销项目销售
                        var page2 = temp[1];
                        DataTable dt = new DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.业态);
                        dt.Columns.Add(Base_Config_Jzgj.组团);
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);         
                        dt.Columns.Add(Base_Config_Rgsj.上上上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上上上周_认购套内均价);
                        dt.Columns.Add(Base_Config_Rgsj.上上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上上周_认购套内均价);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套内均价);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);
                        dt.Columns.Add(Base_Config_Rgsj.本周_变化原因);
                  
                        IAutoShape text2 = (IAutoShape)page2.Shapes[0];
                        text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc);
                    
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_XUHUICHENG_CHIXUXIAOSHOUXIANGMU_Table(page2, dt, 1, null, null);
                            t.AddClone(page2);
                        }
                        #endregion
                    }
                    else
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page3 = temp[2];
                        t.AddClone(page3);
                        var page4 = temp[3];
                        DataTable dt4 = new DataTable();
                        dt4.Columns.Add("kfs");
                        dt4.Columns.Add("hj");
                        dt4.Columns.Add("sssz_cjje");
                        dt4.Columns.Add("ssz_cjje");
                        dt4.Columns.Add("sz_cjje");
                        dt4.Columns.Add("bz_cjje");
                        dt4 = GET_JPXM_ZT_CJJE(dt4, item.jpxmlb);
                        dt4 = GET_JPBA_CJJE(dt4, item);
                        Office_Tables.SetJP_XUHUICHENG_XIAOSHOUE_Table(page4, dt4, 1, null, null);
                        IAutoShape text4 = (IAutoShape)page4.Shapes[0];
                        text4.TextFrame.Text = string.Format(text4.TextFrame.Text, item.bamc);
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
                            Office_Tables.SetJP_XUHUICHENG_XIAOSHOUE_Table(page5, dt5, 0, null, null);
                            IAutoShape text5 = (IAutoShape)page5.Shapes[1];
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

        public DataTable GET_JPXM_BX(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {

                if (item.ytcs[0] == "别墅")
                {
                    if(item.xfytcs!=null&& item.xfytcs.Length > 0) { 
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {

                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                            var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                            var temp_ba_ssz = temp_rgsj_ssz.FirstOrDefault();
                            var temp_ba_sssz = temp_rgsj_sssz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));

                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        var temp_ba_ssz = temp_rgsj_ssz.FirstOrDefault();
                        var temp_ba_sssz = temp_rgsj_sssz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
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
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        var temp_ba_ssz = temp_rgsj_ssz.FirstOrDefault();
                        var temp_ba_sssz = temp_rgsj_sssz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.hxcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                    var temp_ba_ssz = temp_rgsj_ssz != null && temp_rgsj_ssz.Count() > 0 ? temp_rgsj_ssz.FirstOrDefault() : null;
                    var temp_ba_sssz = temp_rgsj_sssz != null && temp_rgsj_sssz.Count() > 0 ? temp_rgsj_sssz.FirstOrDefault() : null;
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                }


            }


            return dt;
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

     
    }
}
