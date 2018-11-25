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
    /// 阳光100阿尔勒
    /// </summary>
    public class plus_jp_yangguang100aerle :plus_jp_base
    {
        public ISlideCollection _plus_jp_yangguang100aerle_1(string str, int cjbh)
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
                    IAutoShape text0_1 = (IAutoShape)page1.Shapes[1];
                    text0_1.TextFrame.Text = string.Format(text0_1.TextFrame.Text, item.bamc, item.ytcs[0]);

                    var page2 = temp[1];
                    DataTable dt2_0 = new DataTable();
                    dt2_0.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt2_0.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                    dt2_0.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                    dt2_0.Columns.Add(Base_Config_Rgsj.本周_新开建面均价);

                    dt2_0.Columns.Add(Base_Config_Rgsj.本周_本周来电);
                    dt2_0.Columns.Add(Base_Config_Rgsj.本周_本周到访量);

                    dt2_0.Columns.Add(Base_Config_Cjba.上周_备案套数);
                    dt2_0.Columns.Add(Base_Config_Cjba.上周_建面均价);
                    dt2_0.Columns.Add(Base_Config_Cjba.上周_套均总价);
                    dt2_0.Columns.Add(Base_Config_Cjba.上周_建筑面积);
                    dt2_0.Columns.Add(Base_Config_Cjba.上周_成交金额);

                    dt2_0.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                    dt2_0.Columns.Add(Base_Config_Rgsj.上周_认购建面均价);
                    dt2_0.Columns.Add(Base_Config_Rgsj.上周_认购金额);

                    dt2_0.Columns.Add(Base_Config_Cjba.本周_备案套数);
                    dt2_0.Columns.Add(Base_Config_Cjba.本周_建面均价);
                    dt2_0.Columns.Add(Base_Config_Cjba.本周_套均总价);
                    dt2_0.Columns.Add(Base_Config_Cjba.本周_建筑面积);
                    dt2_0.Columns.Add(Base_Config_Cjba.本周_成交金额);
                                                    
                    dt2_0.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                    dt2_0.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);
                    dt2_0.Columns.Add(Base_Config_Rgsj.本周_认购金额);
                    dt2_0.Columns.Add("剩余套数");
                    dt2_0.Columns.Add(Base_Config_Rgsj.本周_变化原因);
                    dt2_0 = GET_JPBA_BX(dt2_0, item);
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        IAutoShape text2 = (IAutoShape)page2.Shapes[1];
                        text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc, item.ytcs[0]);
                        dt2_0 = GET_JPXM_BX(dt2_0, item.jpxmlb);
                        Office_Tables.SetJP_CHONGQING18TI_Table(page2, dt2_0, 2, null, null);
                        t.AddClone(page2);
                    }


                    var page3 = temp[2];
                    IAutoShape text3 = (IAutoShape)page3.Shapes[1];
                    text3.TextFrame.Text = string.Format(text3.TextFrame.Text, item.bamc, item.ytcs[0]);

                    DataTable dt1 = new DataTable();
                    dt1.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt1.Columns.Add(Base_Config_Rgsj.本周_优惠);
                    dt1.Columns.Add(Base_Config_Rgsj.本周_活动);
                    dt1.Columns.Add(Base_Config_Rgsj.本周_优惠);
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt1 = GET_JPXM_BX(dt1, item.jpxmlb);
                        Office_Tables.SetTable(page3, dt1, 2, null, null);
                    }
                    t.AddClone(page3);

                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        foreach (var jpitem in item.jpxmlb)
                        {
                            var tp1 = new Presentation(str);
                            var temp1 = tp1.Slides;
                            var page4 = temp1[3];
                            DataTable dttemp = Cache_data_cjjl.jbz.Select("zc>=" + (Base_date.bz - 3)).CopyToDataTable();
                            var data = from a in dttemp.AsEnumerable()
                                       where a["lpmc"].ToString() == jpitem.lpcs[0] && a["yt"].ToString() == jpitem.ytcs[0]
                                       group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                                       select new
                                       {
                                           zc = s.Key.zc,
                                           zcmc = s.Key.zcmc,
                                           ts = s.Sum(m => m["ts"].ints()),
                                           jmjj = s.Sum(m => m["cjje"].longs()) / s.Sum(m => m["jzmj"].doubls())
                                       };
                            DataTable dt4_0 = new DataTable();
                            dt4_0.Columns.Add("周次名称");
                            dt4_0.Columns.Add("精装成交套数");
                            dt4_0.Columns.Add("精装成交均价");
                            foreach (var tempitem in data)
                            {
                                DataRow dr = dt4_0.NewRow();
                                dr["周次名称"] = tempitem.zcmc;
                                dr["精装成交套数"] = tempitem.ts;
                                dr["精装成交均价"] = tempitem.jmjj.je_y();
                                dt4_0.Rows.Add(dr);
                            }
                            IAutoShape text4_0 = (IAutoShape)page4.Shapes[0];
                            text4_0.TextFrame.Text = string.Format(text4_0.TextFrame.Text, item.bamc, item.ytcs[0]);
                            IAutoShape text4_1 = (IAutoShape)page4.Shapes[1];
                            text4_1.TextFrame.Text = string.Format(text4_1.TextFrame.Text, jpitem.lpcs[0], jpitem.ytcs[0]);
                            //无法将类型为“Aspose.Slides.OleObjectFrame”的对象强制转换为类型“Aspose.Slides.Charts.IChart”。
                            //Office_Charts.Chart_gxfx(pag4, dt4_0, 2);
                            t.AddClone(page4);

                            var pag5 = temp1[4];
                            IAutoShape text5_0 = (IAutoShape)pag5.Shapes[1];
                            text5_0.TextFrame.Text = string.Format(text5_0.TextFrame.Text, item.bamc, item.ytcs[0]);
                            IAutoShape text5_1 = (IAutoShape)pag5.Shapes[2];
                            text5_1.TextFrame.Text = string.Format(text5_1.TextFrame.Text, jpitem.lpcs[0], jpitem.ytcs[0]);
                            t.AddClone(pag5);
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
    }
}
