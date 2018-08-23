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
    public class plus_jp_shouchuang : plus_jp_base
    {
        public ISlideCollection _plus_jp_jiangbeizuizhiye_1(string str, int cjbh)
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
                    var page = temp[0];
                    #region 格局统计
                    DataTable dt = new DataTable();
                    dt.Columns.Add(Base_Config_Jzgj.竞争格局名称);
                    dt.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间);
                    dt.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);
                    dt.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                    dt.Columns.Add(Base_Config_Rgsj.本周_新开销售套数);
                    dt.Columns.Add(Base_Config_Rgsj.新开套内均价);

                    dt.Columns.Add("sssz_ts");//上上上周_备案套数
                    dt.Columns.Add("sssz_tnjj");//上上上周_套内均价

                    dt.Columns.Add("ssz_ts");//上上周_备案套数
                    dt.Columns.Add("ssz_tnjj");//上上周_套内均价

                    dt.Columns.Add(Base_Config_Cjba.上周_备案套数);
                    dt.Columns.Add(Base_Config_Cjba.上周_建面均价);

                    dt.Columns.Add(Base_Config_Cjba.本周_备案套数);
                    dt.Columns.Add(Base_Config_Cjba.本周_建面均价);

                    dt.Columns.Add(Base_Config_Rgsj.变化原因);
                    dt.Columns.Add(Base_Config_Rgsj.营销动作);
                    #endregion
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_BX(dt, item.jpxmlb);
                        Office_Tables.SetJP_ShouChuang_JPBX_Table(page, dt, 1, null, null);
                        t.AddClone(page);
                    }
                }
                return t;
            }
            catch (Exception)
            {
                return null;
            }

        }

        /// <summary>
        /// 竞品表现
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
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {

                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_ssz = (Cache_data_cjjl.jbz.Select("zc=" + (Base_date.bz - 2)).CopyToDataTable()).AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_sssz = (Cache_data_cjjl.jbz.Select("zc=" + (Base_date.bz - 3)).CopyToDataTable()).AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_cjba_bz, temp_cjba_sz, temp_cjba_ssz, temp_cjba_sssz, item));

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
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[i]);
                        var temp_cjba_ssz = (Cache_data_cjjl.jbz.Select("zc=" + (Base_date.bz - 2)).CopyToDataTable()).AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[i]);
                        var temp_cjba_sssz = (Cache_data_cjjl.jbz.Select("zc=" + (Base_date.bz - 3)).CopyToDataTable()).AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[i]);

                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion
                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_cjba_bz, temp_cjba_sz, temp_cjba_ssz, temp_cjba_sssz, item));
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                    var temp_cjba_ssz = (Cache_data_cjjl.jbz.Select("zc=" + (Base_date.bz - 2)).CopyToDataTable()).AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sssz = (Cache_data_cjjl.jbz.Select("zc=" + (Base_date.bz - 3)).CopyToDataTable()).AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);


                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_cjba_bz, temp_cjba_sz, temp_cjba_ssz, temp_cjba_sssz, item));
                }


            }


            return dt;
        }


        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                               DataRow temp_ba_bz,
                               EnumerableRowCollection<DataRow> temp_cjba_bz,
                               EnumerableRowCollection<DataRow> temp_cjba_sz,
                               EnumerableRowCollection<DataRow> temp_cjba_ssz,
                               EnumerableRowCollection<DataRow> temp_cjba_sssz,
                               JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {


                if (Base_Config_Rgsj._认购数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Rgsj.本周_新开销售套数:
                        case Base_Config_Rgsj.本周_新开套数:
                        case Base_Config_Rgsj.本周_认购套数:
                        case Base_Config_Rgsj.本周_认购套内均价:
                        case Base_Config_Rgsj.本周_认购建面均价:
                        case Base_Config_Rgsj.本周_认购套内体量:
                        case Base_Config_Rgsj.本周_认购建面体量:
                        case Base_Config_Rgsj.本周_认购金额:
                            {
                                if (temp_ba_bz != null)
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                }
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                        default:
                            {
                                if (temp_ba_bz != null)
                                {
                                     dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName];  
                                }
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                    }
                }
                else if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                {
                    if (temp_cjba_bz != null)
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Cjba.本周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()); }; break;
                            case Base_Config_Cjba.本周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()); }; break;
                            case Base_Config_Cjba.本周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()); }; break;
                            case Base_Config_Cjba.本周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].ints()); }; break;
                            case Base_Config_Cjba.本周_建面均价: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()) != 0 ? (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls())).je_y() : 0;}; break;
                            case Base_Config_Cjba.本周_套内均价: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls()) != 0 ? (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls())).je_y() : 0; }; break;

                        }
                    }

                    if (temp_cjba_sz != null)
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Cjba.上周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_备案套数._ConfigCjbaMc()].ints()); }; break;
                            case Base_Config_Cjba.上周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()); }; break;
                            case Base_Config_Cjba.上周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()); }; break;
                            case Base_Config_Cjba.上周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].ints()); }; break;
                            case Base_Config_Cjba.上周_建面均价: {
                                    var cjje = temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs());
                                    var jzmj = temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls());
                                    //var jmjj = temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()) != 0 ? (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls())).je_y() : 0;
                                    dr1[dt.Columns[j].ColumnName] = jzmj!=0? (cjje / jzmj).je_y():0;
                                }; break;
                            case Base_Config_Cjba.上周_套内均价: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls()) != 0 ? (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls())).je_y() : 0; }; break;
                        }
                    }
                }
                else if (Base_Config_Jzgj._竞争格局参数名称.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Jzgj.项目名称:{ dr1[dt.Columns[j].ColumnName] = item.lpcs[0]; }; break;
                        case Base_Config_Jzgj.业态:{ dr1[dt.Columns[j].ColumnName] = yt; }; break;
                        case Base_Config_Jzgj.竞争格局名称: { dr1[dt.Columns[j].ColumnName] = item.jzgjmc; }; break;
                        case Base_Config_Jzgj.竞争格局_主力面积区间: { dr1[dt.Columns[j].ColumnName] = item.zlmjqj; }; break;
                        default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                    }

                }
                else
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case "sssz_ts": { dr1[dt.Columns[j].ColumnName] = temp_cjba_sssz.Sum(m => m["ts"].ints()); }; break;
                        case "sssz_tnjj": { dr1[dt.Columns[j].ColumnName] = temp_cjba_sssz.Sum(m => m["jzmj"].doubls()) != 0 ? (temp_cjba_sssz.Sum(m => m["cjje"].longs()) / temp_cjba_sssz.Sum(m => m["jzmj"].doubls())).je_y() : 0; }; break;
                        case "ssz_ts": { dr1[dt.Columns[j].ColumnName] = temp_cjba_ssz.Sum(m => m["ts"].doubls()); }; break;
                        case "ssz_tnjj": { dr1[dt.Columns[j].ColumnName] = temp_cjba_ssz.Sum(m => m["jzmj"].doubls()) != 0 ? (temp_cjba_ssz.Sum(m => m["cjje"].longs()) / temp_cjba_ssz.Sum(m => m["jzmj"].doubls())).je_y() : 0; }; break;

                    }
                }
            }

            return dr1;
        }
    }
}
