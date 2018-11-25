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
    public  class plus_jp_lvdi:plus_jp_base
    {

        public ISlideCollection _plus_jp_lvdi_1(string str, int cjbh)
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

                    if (item.ytcs[0] != "商务")
                    {

                        #region 格局统计
                        var page1 = temp[0];
                        DataTable dt = new DataTable();

                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);

                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套内均价);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);

                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面均价环比);

                        dt.Columns.Add("bzcl");
                        dt.Columns.Add();
                        dt.Columns.Add("bybajmjj");
                        dt.Columns.Add(Base_Config_Rgsj.本周_营销动作);

                        IAutoShape text2 = (IAutoShape)page1.Shapes[0];
                        text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc);
                        #endregion
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_LVDI_PUTONG_Table(page1, dt, 1, null, null);
                            t.AddClone(page1);
                        }

                        foreach (var page3 in _plus_jp_dyt_tgtp(item))
                        {
                            t.AddClone(page3);
                        }
                    }
                    else
                    {
                        #region 格局统计
                        var page2 = temp[1];
                        DataTable dt = new DataTable();

                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);

                        dt.Columns.Add(Base_Config_Rgsj.上上上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上上上周_认购建面均价);
                        dt.Columns.Add(Base_Config_Cjba.上上上周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.上上上周_建面均价);

                        dt.Columns.Add(Base_Config_Rgsj.上上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上上周_认购建面均价);
                        dt.Columns.Add(Base_Config_Cjba.上上周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.上上周_建面均价);

                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购建面均价);
                        dt.Columns.Add(Base_Config_Cjba.上周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.上周_建面均价);

                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);
                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.本周_建面均价);

                        
                        dt.Columns.Add(Base_Config_Rgsj.本周_营销动作);

                        IAutoShape text2 = (IAutoShape)page2.Shapes[0];
                        text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc);

                        
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_LVDI_PUTONG_Table(page2, dt, 1, null, null);
                            t.AddClone(page2);
                        }
                        #endregion
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
                #region 别墅
               

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
                            var rgsj_bz = temp_rgsj_bz.FirstOrDefault();
                            var rgsj_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, rgsj_bz, rgsj_sz, temp_cjba_bz, temp_cjba_sz, item));

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
                        var rgsj_bz = temp_rgsj_bz.FirstOrDefault();
                        var rgsj_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, rgsj_bz, rgsj_sz, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                #endregion
                else if (item.ytcs[0] == "商务")
                {
                    for (int i = 0; i < item.hxcs.Length; i++)
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        //var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        //var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                        var temp_cjba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                        var temp_cjba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                        //本周本案认购数据
                        //var temp_rgsj_bz = temp_rgsj_bz.FirstOrDefault();
                        //var temp_rgsj_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, null, null, null, null, temp_cjba_bz, temp_cjba_sz, temp_cjba_ssz, temp_cjba_sssz, item));
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
                    var rgsj_bz = temp_rgsj_bz.FirstOrDefault();
                    var rgsj_sz = temp_rgsj_sz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, rgsj_bz, rgsj_sz, temp_cjba_bz, temp_cjba_sz, item));
                }


            }


            return dt;
        }


        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                        DataRow temp_rgsj_bz,
                        DataRow temp_rgsj_sz,
                        DataRow temp_rgsj_ssz,
                        DataRow temp_rgsj_sssz,
                        EnumerableRowCollection<DataRow> temp_cjba_bz,
                        EnumerableRowCollection<DataRow> temp_cjba_sz,
                        EnumerableRowCollection<DataRow> temp_cjba_ssz,
                        EnumerableRowCollection<DataRow> temp_cjba_sssz,
                        JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {

                try
                {

                    #region 认购
                    

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
                                {
                                    if (temp_rgsj_sz != null)
                                    {
                                        dr1[dt.Columns[j].ColumnName] = temp_rgsj_sz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
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
                                {
                                    if (temp_rgsj_bz != null)
                                    {
                                        dr1[dt.Columns[j].ColumnName] = temp_rgsj_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                    }
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Rgsj.上上周_新开套数:
                            case Base_Config_Rgsj.上上周_新开销售套数:
                            case Base_Config_Rgsj.上上周_新开建面均价:
                            case Base_Config_Rgsj.上上周_新开套内均价:
                            case Base_Config_Rgsj.上上周_认购套数:
                            case Base_Config_Rgsj.上上周_认购套内体量:
                            case Base_Config_Rgsj.上上周_认购套内均价:
                            case Base_Config_Rgsj.上上周_认购建面体量:
                            case Base_Config_Rgsj.上上周_认购建面均价:
                            case Base_Config_Rgsj.上上周_认购金额:
                                {
                                    if (temp_rgsj_ssz != null)
                                    {
                                        dr1[dt.Columns[j].ColumnName] = temp_rgsj_ssz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                    }
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Rgsj.上上上周_新开套数:
                            case Base_Config_Rgsj.上上上周_新开销售套数:
                            case Base_Config_Rgsj.上上上周_新开建面均价:
                            case Base_Config_Rgsj.上上上周_新开套内均价:
                            case Base_Config_Rgsj.上上上周_认购套数:
                            case Base_Config_Rgsj.上上上周_认购套内体量:
                            case Base_Config_Rgsj.上上上周_认购套内均价:
                            case Base_Config_Rgsj.上上上周_认购建面体量:
                            case Base_Config_Rgsj.上上上周_认购建面均价:
                            case Base_Config_Rgsj.上上上周_认购金额:
                                {
                                    if (temp_rgsj_sssz != null)
                                    {
                                        dr1[dt.Columns[j].ColumnName] = temp_rgsj_ssz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                    }
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Rgsj.本周_认购套数环比: { dr1[dt.Columns[j].ColumnName] = temp_rgsj_sz["rgts"] != null && temp_rgsj_sz["rgts"].ints() != 0 ? ((temp_rgsj_bz["rgts"].ints() - temp_rgsj_sz["rgts"].ints()) / temp_rgsj_sz["rgts"].ints()).doubls().ss_bfb() : "0%"; }; break;
                            case Base_Config_Rgsj.本周_认购金额环比: { dr1[dt.Columns[j].ColumnName] = temp_rgsj_sz["rgje"] != null && temp_rgsj_sz["rgje"].ints() != 0 ? ((temp_rgsj_bz["rgts"].ints() - temp_rgsj_sz["rgts"].ints()) / temp_rgsj_sz["rgts"].ints()).doubls().ss_bfb() : "0%"; }; break;
                            case Base_Config_Rgsj.本周_认购建筑面积环比: { dr1[dt.Columns[j].ColumnName] = temp_rgsj_sz["rgjmtl"] != null && temp_rgsj_sz["rgjmtl"].ints() != 0 ? ((temp_rgsj_bz["rgjmtl"].ints() - temp_rgsj_sz["rgjmtl"].ints()) / temp_rgsj_sz["rgjmtl"].ints()).doubls().ss_bfb() : "0%"; }; break;
                            case Base_Config_Rgsj.本周_认购套内面积环比: { }; break;
                            case Base_Config_Rgsj.本周_认购建面均价环比: { dr1[dt.Columns[j].ColumnName] = temp_rgsj_sz["rgjmjj"] != null && temp_rgsj_sz["rgjmjj"].ints() != 0 ? ((temp_rgsj_bz["rgjmjj"].ints() - temp_rgsj_sz["rgjmjj"].ints()) / temp_rgsj_sz["rgjmjj"].ints()).doubls().ss_bfb() : "0%"; }; break;
                            case Base_Config_Rgsj.本周_认购套内均价环比: { }; break;
                            case Base_Config_Rgsj.本周_认购套均总价环比: { }; break;
                            default:
                                {
                                    try
                                    {
                                        if (temp_rgsj_bz != null)
                                        {
                                            dr1[dt.Columns[j].ColumnName] = temp_rgsj_bz[dt.Columns[j].ColumnName];
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
                    #endregion
                    #region 备案

                    
                    else if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            #region 本周

                            
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
                            #endregion
                            #region 上周

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
                            #endregion
                            #region 上上周
                            

                            case Base_Config_Cjba.上上周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_ssz != null ? temp_cjba_ssz.Sum(m => m[Base_Config_Cjba.上上周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.上上周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_ssz != null ? temp_cjba_ssz.Sum(m => m[Base_Config_Cjba.上上周_成交金额._ConfigCjbaMc()].longs()) : 0; }; break;
                            case Base_Config_Cjba.上上周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_ssz != null ? temp_cjba_ssz.Sum(m => m[Base_Config_Cjba.上上周_建筑面积._ConfigCjbaMc()].doubls()) : 0; }; break;
                            case Base_Config_Cjba.上上周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_ssz != null ? temp_cjba_ssz.Sum(m => m[Base_Config_Cjba.上上周_套内面积._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.上上周_建面均价:
                                {
                                    if ((temp_cjba_ssz != null && temp_cjba_ssz.Sum(m => m[Base_Config_Cjba.上上周_备案套数._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_ssz.Sum(m => m[Base_Config_Cjba.上上周_备案套数._ConfigCjbaMc()].longs()) / temp_cjba_ssz.Sum(m => m[Base_Config_Cjba.上上周_备案套数._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Cjba.上上周_套内均价:
                                {
                                    if ((temp_cjba_ssz != null && temp_cjba_ssz.Sum(m => m[Base_Config_Cjba.上上周_备案套数._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_ssz.Sum(m => m[Base_Config_Cjba.上上周_备案套数._ConfigCjbaMc()].longs()) / temp_cjba_ssz.Sum(m => m[Base_Config_Cjba.上上周_备案套数._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            #endregion
                            #region 上上上周
                            case Base_Config_Cjba.上上上周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sssz != null ? temp_cjba_sssz.Sum(m => m[Base_Config_Cjba.上上上周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.上上上周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sssz != null ? temp_cjba_sssz.Sum(m => m[Base_Config_Cjba.上上上周_成交金额._ConfigCjbaMc()].longs()) : 0; }; break;
                            case Base_Config_Cjba.上上上周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sssz != null ? temp_cjba_sssz.Sum(m => m[Base_Config_Cjba.上上上周_建筑面积._ConfigCjbaMc()].doubls()) : 0; }; break;
                            case Base_Config_Cjba.上上上周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sssz != null ? temp_cjba_sssz.Sum(m => m[Base_Config_Cjba.上上上周_套内面积._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.上上上周_建面均价:
                                {
                                    if ((temp_cjba_sssz != null && temp_cjba_sssz.Sum(m => m[Base_Config_Cjba.上上上周_备案套数._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_sssz.Sum(m => m[Base_Config_Cjba.上上上周_备案套数._ConfigCjbaMc()].longs()) / temp_cjba_sssz.Sum(m => m[Base_Config_Cjba.上上上周_备案套数._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            case Base_Config_Cjba.上上上周_套内均价:
                                {
                                    if ((temp_cjba_sssz != null && temp_cjba_sssz.Sum(m => m[Base_Config_Cjba.上上上周_备案套数._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_sssz.Sum(m => m[Base_Config_Cjba.上上上周_备案套数._ConfigCjbaMc()].longs()) / temp_cjba_sssz.Sum(m => m[Base_Config_Cjba.上上上周_备案套数._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "0";
                                    }
                                }; break;
                            #endregion

                            #region 其他

                            
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
                            #endregion


                            default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                        }


                    }
                    #endregion
                    #region 竞争格局
                  
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
                                    dr1[dt.Columns[j].ColumnName] = item.zlmjqj;
                                }; break;
                            case Base_Config_Jzgj.竞争格局名称:
                                {
                                    dr1[dt.Columns[j].ColumnName] = item.jzgjmc;
                                }; break;
                        }

                    }

                    #endregion
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
