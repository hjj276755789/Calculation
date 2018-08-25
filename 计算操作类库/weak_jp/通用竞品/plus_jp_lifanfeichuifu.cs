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
    /// 力帆翡翠府
    /// </summary>
    public class plus_jp_lifanfeichuifu :plus_jp_base
    {
        public ISlideCollection _plus_jp_biguiyuan_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);

                foreach (var item in jbzzs(str))
                {
                    t.AddClone(item);
                }

                foreach (var item in param)
                {

                    var tp = new Presentation(str);
                    var temp = tp.Slides;

                    #region 竞品分布
                    var page1 = temp[1];
                    IAutoShape text1 = (IAutoShape)page1.Shapes[1];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc);
                    #endregion
                    t.AddClone(page1);



                    #region 市场表现
                    var page2 = temp[2];
                    DataTable dt = new DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);
                        dt.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间);

                        dt.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_新开套内均价);
                        dt.Columns.Add(Base_Config_Rgsj.本周到访量);

                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购建面体量);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购金额);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购建面均价);

                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面体量);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购金额);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);
                        dt.Columns.Add(Base_Config_Rgsj.营销动作);
                   
                    IAutoShape text2 = (IAutoShape)page2.Shapes[2];
                    text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc);
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_BX(dt, item.jpxmlb);
                        Office_Tables.SetJP_LiFanFeiCuiFu_JPBX_Table(page2, dt, 5, null, null);
                        t.AddClone(page2);
                    }
                    #endregion

                    #region 加推计划 ---不做


                    #endregion
                 

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

        public ISlideCollection jbzzs(string str)
        {
            var p = new Presentation();
            var t = p.Slides;
            t.RemoveAt(0);
            var tp = new Presentation(str);
            var temp = tp.Slides;



            string[] zt = { "蔡家"};
            #region  P1
            var page1 = temp[0];
            IAutoShape text2 = (IAutoShape)page1.Shapes[0];
            text2.TextFrame.Text = string.Format(text2.TextFrame.Text, string.Join("、", zt));
            IAutoShape text3 = (IAutoShape)page1.Shapes[0];
            text3.TextFrame.Text = string.Format(text3.TextFrame.Text, string.Join("、", zt));
            var dt1_1 = from a in Cache_data_xzys.jbz.AsEnumerable()
                        where zt.Contains(a["zt"])
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            xzgyl = s.Sum(m => m["jzmj"].doubls())+s.Sum(m=>m["fzzmj"].doubls())
                        };
            var dt1_2 = from a in Cache_data_cjjl.jbz.AsEnumerable()
                        where zt.Contains(a["zt"])
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            cjje = s.Sum(m => m["cjje"].longs()),
                            jzmj = s.Sum(m => m["jzmj"].doubls())
                        };
            DataTable dt1 = new DataTable();
            dt1.Columns.Add("周次");
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 7), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 6), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 5), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 4), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1), typeof(double));
            dt1.Columns.Add(Base_date.GET_ZCMC(Base_date.bn, Base_date.bz), typeof(double));
            DataRow dr1 = dt1.NewRow();
            DataRow dr2 = dt1.NewRow();
            DataRow dr3 = dt1.NewRow();
            dr1[0] = "供应体量";
            dr2[0] = "成交体量";
            dr3[0] = "建面均价";

            for (int i = 0; i < 8; i++)
            {
                var xzys = dt1_1.FirstOrDefault(m => m.zc.ints() == (Base_date.bz - (7 - i)));
                var cjba = dt1_2.FirstOrDefault(m => m.zc.ints() == (Base_date.bz - (7 - i)));
                dr1[i + 1] = (xzys != null) ? xzys.xzgyl.mj_wf() : 0;
                dr2[i + 1] = cjba != null ? cjba.jzmj.mj_wf() : 0;
                dr3[i + 1] = cjba != null ? (cjba.cjje / cjba.jzmj).je_y() : 0;
                if (i == 7)
                {
                    IAutoShape text1 = (IAutoShape)page1.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, (xzys != null) ? xzys.xzgyl.mj_wf() : 0, cjba != null ? cjba.jzmj.mj_wf() : 0, cjba != null ? (cjba.cjje / cjba.jzmj).je_y() : 0);
                }
            }
            dt1.Rows.Add(dr1);
            dt1.Rows.Add(dr2);
            dt1.Rows.Add(dr3);
            Office_Charts.Chart_jp_langshi_chart1(page1, dt1, 4);
            t.AddClone(page1);
            #endregion
           
            return t;
        }
    }
}
