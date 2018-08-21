using Aspose.Slides;
using Aspose.Slides.Charts;
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
    public class plus_jp_dongyuandichan : weak
    {
        public ISlideCollection _plus_jp_dongyuandichan_1(string str, int cjbh)
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
                    #region 格局统计


                    var page = temp[1];
                    IAutoShape text = (IAutoShape)page.Shapes[2];
                    text.TextFrame.Text = string.Format(text.TextFrame.Text, item.bamc, item.ytcs[0]);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add(Base_Config_Rgsj.项目名称);
                    dt.Columns.Add(Base_Config_Rgsj.业态);
                    dt.Columns.Add(Base_Config_Cjba.备案套数);
                    dt.Columns.Add(Base_Config_Cjba.建面均价);
                    dt.Columns.Add(Base_Config_Rgsj.认购套数);
                    dt.Columns.Add(Base_Config_Rgsj.认购建面均价);
                    dt.Columns.Add(Base_Config_Rgsj.营销动作);

                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_BX(dt, item.jpxmlb);
                        Office_Tables.SetJP_RUIAN_JPBX_Table(page, dt.AsEnumerable().OrderBy(m => m["jzgjmc"]).CopyToDataTable(), 4, null, null);
                        t.AddClone(page);
                    }
                    #endregion

                }
            }
            catch (Exception e)
            {

            }
            return null;
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
                        dr1["jzgjmc"] = item.qycs[0];//竞争业态
                        dr1["lpmc"] = item.lpcs[0];//竞争楼盘名称

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion

                        #region 本周认购数据
                        //if (temp_ba_bz != null)
                        //{
                        //    dr1["xkts"] = temp_ba_bz["xkts"];
                        //    dr1["xkxsts"] = temp_ba_bz["xkxsts"];
                        //    dr1["xktnjj"] = temp_ba_bz["xktnjj"];
                        //    dr1["bzrgts"] = temp_ba_bz["rgts"];
                        //    dr1["bzrgtnjj"] = temp_ba_bz["rgtnjj"];

                        //}
                        //else
                        //{
                        //    dr1["xkts"] = "";
                        //    dr1["xkxsts"] = "";
                        //    dr1["xktnjj"] = "";
                        //    dr1["bzrgts"] = 0;
                        //    dr1["bzrgtnjj"] = 0;
                        //}

                        //#endregion

                        //#region 本周成交数据
                        //if (temp_ba_bz != null)
                        //{
                        //    dr1["bzbats"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                        //    dr1["bzbatnjj"] = temp_cjba_bz.Sum(n => n["tnmj"].ints()) != 0 ? (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].doubls())).je_y() : 0;
                        //}
                        //else
                        //{
                        //    dr1["bzbats"] = 0;
                        //    dr1["bzbatnjj"] = 0;
                        //}
                        #endregion

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (Base_Config_Rgsj._认购数据.Contains(dt.Columns[j].ColumnName))
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName];
                            else
                            {
                                switch (dt.Columns[j].ColumnName)
                                {
                                    case Base_Config_Cjba.备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.备案套数].ints()); }; break;
                                    case Base_Config_Cjba.成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.成交金额].longs()); }; break;
                                    case Base_Config_Cjba.建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.建筑面积].doubls()); }; break;
                                    case Base_Config_Cjba.套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.套内面积].ints()); }; break;
                                    case Base_Config_Cjba.建面均价: { dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.成交金额].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.建筑面积].doubls())).je_y(); }; break;
                                    case Base_Config_Cjba.套内均价: { dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.成交金额].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.套内面积].doubls())).je_y(); }; break;
                                    default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                                }
                            }
                        }
                        dt.Rows.Add(dr1);

                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    for (int i = 0; i < item.hxcs.Length; i++)
                    {
                        DataRow dr1 = dt.NewRow();
                        dr1["jzgjmc"] = item.qycs[0];//竞争业态
                        dr1["lpmc"] = item.lpcs[0];//竞争楼盘名称
                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);

                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion

                        //#region 本周认购数据
                        //if (temp_ba_bz != null)
                        //{
                        //    dr1["xkts"] = temp_ba_bz["xkts"];
                        //    dr1["xkxsts"] = temp_ba_bz["xkxsts"];
                        //    dr1["xktnjj"] = temp_ba_bz["xktnjj"];
                        //    dr1["bzrgts"] = temp_ba_bz["rgts"];
                        //    dr1["bzrgtnjj"] = temp_ba_bz["rgtnjj"];

                        //}
                        //else
                        //{
                        //    dr1["xkts"] = "";
                        //    dr1["xkxsts"] = "";
                        //    dr1["xktnjj"] = "";
                        //    dr1["bzrgts"] = 0;
                        //    dr1["bzrgtnjj"] = 0;
                        //}

                        //#endregion

                        //#region 本周成交数据
                        //if (temp_ba_bz != null)
                        //{
                        //    dr1["bzbats"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                        //    dr1["bzbatnjj"] = temp_cjba_bz.Sum(n => n["tnmj"].ints()) != 0 ? (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].doubls())).je_y() : 0;

                        //}
                        //else
                        //{
                        //    dr1["bzbats"] = 0;
                        //    dr1["bzbatnjj"] = 0;
                        //}
                        //#endregion


                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (Base_Config_Rgsj._认购数据.Contains(dt.Columns[j].ColumnName))
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName];
                            else
                            {
                                switch (dt.Columns[j].ColumnName)
                                {
                                    case Base_Config_Cjba.备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.备案套数].ints()); }; break;
                                    case Base_Config_Cjba.成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.成交金额].longs()); }; break;
                                    case Base_Config_Cjba.建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.建筑面积].doubls()); }; break;
                                    case Base_Config_Cjba.套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.套内面积].ints()); }; break;
                                    case Base_Config_Cjba.建面均价: { dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.成交金额].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.建筑面积].doubls())).je_y(); }; break;
                                    case Base_Config_Cjba.套内均价: { dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.成交金额].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.套内面积].doubls())).je_y(); }; break;
                                    default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                                }
                            }
                        }

                        dt.Rows.Add(dr1);
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();
                    dr1["jzgjmc"] = item.jzgjmc;//竞争业态
                    dr1["lpmc"] = item.lpcs[0];//竞争楼盘名称

                    #region 数据准备
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    #endregion

                    //#region 本周认购数据
                    //if (temp_ba_bz != null)
                    //{
                    //    dr1["xkts"] = temp_ba_bz["xkts"];
                    //    dr1["xkxsts"] = temp_ba_bz["xkxsts"];
                    //    dr1["xktnjj"] = temp_ba_bz["xktnjj"];
                    //    dr1["bzrgts"] = temp_ba_bz["rgts"];
                    //    dr1["bzrgtnjj"] = temp_ba_bz["rgtnjj"];

                    //}
                    //else
                    //{
                    //    dr1["xkts"] = "";
                    //    dr1["xkxsts"] = "";
                    //    dr1["xktnjj"] = "";
                    //    dr1["bzrgts"] = 0;
                    //    dr1["bzrgtnjj"] = 0;
                    //}

                    //#endregion

                    //#region 本周成交数据
                    //if (temp_ba_bz != null)
                    //{
                    //    dr1["bzbats"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                    //    dr1["bzbatnjj"] = temp_cjba_bz.Sum(n => n["tnmj"].ints()) != 0 ? (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].doubls())).je_y() : 0;

                    //}
                    //else
                    //{
                    //    dr1["bzbats"] = 0;
                    //    dr1["bzbatnjj"] = 0;
                    //}
                    //#endregion
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (Base_Config_Rgsj._认购数据.Contains(dt.Columns[j].ColumnName))
                            dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName];
                        else
                        {
                            switch (dt.Columns[j].ColumnName)
                            {
                                case Base_Config_Cjba.备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.备案套数].ints()); }; break;
                                case Base_Config_Cjba.成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.成交金额].longs()); }; break;
                                case Base_Config_Cjba.建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.建筑面积].doubls()); }; break;
                                case Base_Config_Cjba.套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz.Sum(m => m[Base_Config_Cjba.套内面积].ints()); }; break;
                                case Base_Config_Cjba.建面均价: { dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.成交金额].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.建筑面积].doubls())).je_y(); }; break;
                                case Base_Config_Cjba.套内均价: { dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.成交金额].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.套内面积].doubls())).je_y(); }; break;
                                default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                            }
                        }
                    }

                    dt.Rows.Add(dr1);
                }


            }


            return dt;
        }

    }
}
