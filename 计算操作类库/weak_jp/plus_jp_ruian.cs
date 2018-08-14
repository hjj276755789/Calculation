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
    /// 瑞安竞品
    /// </summary>
    public class plus_jp_ruian :plus_jp_base
    {
        public ISlideCollection _plus_jp_ruian_1(string str, int cjbh)
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
                    IAutoShape text = (IAutoShape)page.Shapes[2];
                    text.TextFrame.Text = string.Format(text.TextFrame.Text, item.bamc, item.ytcs[0]);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add("qy");
                    dt.Columns.Add("lpmc");
                    dt.Columns.Add("yt");

                    dt.Columns.Add("xkts");
                    dt.Columns.Add("xkxsts");
                    dt.Columns.Add("xktnjj");

                    dt.Columns.Add("szbats");
                    dt.Columns.Add("szbatnjj");
                    dt.Columns.Add("szrgts");
                    dt.Columns.Add("szrgtnjj");

                    dt.Columns.Add("bzbats");
                    dt.Columns.Add("bzbatnjj");
                    dt.Columns.Add("bzrgts");
                    dt.Columns.Add("bzrgtnjj");

                    dt.Columns.Add("thb");
                    dt.Columns.Add("jghb");
                    dt.Columns.Add("bhyy");
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_BX(dt, item.jpxmlb);
                        Office_Tables.SetJP_RUIAN_JPBX_Table(page, dt.AsEnumerable().OrderBy(m=>m["qy"]).CopyToDataTable(), 4, null, null);
                        t.AddClone(page);
                    }


                    var page1 = temp[1];
                    IAutoShape text1 = (IAutoShape)page1.Shapes[1];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                    System.Data.DataTable dt1 = new System.Data.DataTable();
                    dt1.Columns.Add("xm");
                    dt1.Columns.Add("yh");
                    dt1.Columns.Add("yxdz");
                    dt1.Columns.Add("xzjtyj");
                    dt1.Columns.Add("bkfs");
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_JQDZ(dt1, item.jpxmlb);
                        Office_Tables.SetJP_RUIAN_JQHD_Table(page1, dt, 3, null, null);
                        t.AddClone(page1);
                    }
                }

                foreach (var item in _plus_jp_dyt_tgtp(cjbh))
                {
                    if (item != null)
                        t.AddClone(item);
                }
                return t;
            }
            catch(Exception e)
            {
                Base_Log.Log(e.Message);
                return null;
            }

        }


        public DataTable GET_JPXM_BX(DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {
                if (item.ytcs[0] == "别墅")
                {
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {

                        DataRow dr1 = dt.NewRow();
                        dr1[0] = item.qycs[0];//竞争业态
                        dr1[1] = item.lpcs[0];//竞争楼盘名称
                        dr1[2] = item.xfytcs[i];//竞争业态

                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                        //上周本案认购数据
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion

                        #region 上周认购数据
                        if (temp_ba_bz != null)
                        {
                            dr1["szrgts"] = temp_ba_sz["xkts"];
                            dr1["szrgtnjj"] = temp_ba_sz["xktnjj"];

                        }
                        else
                        {
                            dr1["szrgts"] = "";
                            dr1["szrgtnjj"] = "";
                        }
                        #endregion


                        #region 本周认购数据
                        if (temp_ba_bz != null)
                        {
                            dr1["xkts"] = temp_ba_bz["xkts"];
                            dr1["xkxsts"] = temp_ba_bz["xkxsts"];
                            dr1["xktnjj"] = temp_ba_bz["xktnjj"];
                            dr1["bzrgts"] = temp_ba_bz["rgts"];
                            dr1["bzrgtnjj"] = temp_ba_bz["rgtnjj"];

                        }
                        else
                        {
                            dr1["xkts"] = "";
                            dr1["xkxsts"] = "";
                            dr1["xktnjj"] = "";
                            dr1["bzrgts"] = 0;
                            dr1["bzrgtnjj"] = 0;
                        }

                        #endregion
                        #region  上周成交数据
                        if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                        {

                            dr1["szbats"] = temp_cjba_sz.Sum(m => m["ts"].ints());
                            dr1["szbatnjj"] = temp_cjba_sz.Sum(n => n["tnmj"].ints()) != 0 ? (temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(n => n["tnmj"].doubls())).je_y() : 0;
                        }
                        else
                        {
                            dr1["szbats"] = 0;
                            dr1["szbatnjj"] = 0;
                        }
                        #endregion

                        #region 本周成交数据
                        if (temp_ba_bz != null)
                        {
                            dr1["bzbats"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                            dr1["bzbatnjj"] = temp_cjba_bz.Sum(n => n["tnmj"].ints()) != 0 ? (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].doubls())).je_y() : 0;

                        }
                        else
                        {
                            dr1["bzbats"] = 0;
                            dr1["bzbatnjj"] = 0;
                        }
                        #endregion

                        dt.Rows.Add(dr1);

                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    for (int i = 0; i < item.hxcs.Length; i++)
                    {
                        DataRow dr1 = dt.NewRow();
                        dr1[0] = item.qycs[0];//区域
                        dr1[1] = item.lpcs[0];//竞争楼盘名称
                        dr1[2] = item.hxcs[i];//竞争业态
                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);

                        //上周本案认购数据
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion


                        #region 上周认购数据
                        if (temp_ba_bz != null)
                        {
                            dr1["szrgts"] = temp_ba_sz["xkts"];
                            dr1["szrgtnjj"] = temp_ba_sz["xktnjj"];

                        }
                        else
                        {
                            dr1["szrgts"] = "";
                            dr1["szrgtnjj"] = "";
                        }
                        #endregion


                        #region 本周认购数据
                        if (temp_ba_bz != null)
                        {
                            dr1["xkts"] = temp_ba_bz["xkts"];
                            dr1["xkxsts"] = temp_ba_bz["xkxsts"];
                            dr1["xktnjj"] = temp_ba_bz["xktnjj"];
                            dr1["bzrgts"] = temp_ba_bz["rgts"];
                            dr1["bzrgtnjj"] = temp_ba_bz["rgtnjj"];

                        }
                        else
                        {
                            dr1["xkts"] = "";
                            dr1["xkxsts"] = "";
                            dr1["xktnjj"] = "";
                            dr1["bzrgts"] = 0;
                            dr1["bzrgtnjj"] = 0;
                        }

                        #endregion
                        #region  上周成交数据
                        if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                        {

                            dr1["szbats"] = temp_cjba_sz.Sum(m => m["ts"].ints());
                            dr1["szbatnjj"] = temp_cjba_sz.Sum(n => n["tnmj"].ints()) != 0 ? (temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(n => n["tnmj"].doubls())).je_y() : 0;
                        }
                        else
                        {
                            dr1["szbats"] = 0;
                            dr1["szbatnjj"] = 0;
                        }
                        #endregion

                        #region 本周成交数据
                        if (temp_ba_bz != null)
                        {
                            dr1["bzbats"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                            dr1["bzbatnjj"] = temp_cjba_bz.Sum(n => n["tnmj"].ints()) != 0 ? (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].doubls())).je_y() : 0;

                        }
                        else
                        {
                            dr1["bzbats"] = 0;
                            dr1["bzbatnjj"] = 0;
                        }
                        #endregion

                        dt.Rows.Add(dr1);
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();
                    dr1[0] = item.qycs[0];//区域
                    dr1[1] = item.lpcs[0];//竞争楼盘名称
                    dr1[2] = item.ytcs[0];//竞争业态
                    #region 数据准备
                    //竞品业态

                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                    //上周本案认购数据
                    var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    #endregion


                    #region 上周认购数据
                    if (temp_ba_bz != null)
                    {
                        dr1["szrgts"] = temp_ba_sz["xkts"];
                        dr1["szrgtnjj"] = temp_ba_sz["xktnjj"];

                    }
                    else
                    {
                        dr1["szrgts"] = "";
                        dr1["szrgtnjj"] = "";
                    }
                    #endregion


                    #region 本周认购数据
                    if (temp_ba_bz != null)
                    {
                        dr1["xkts"] = temp_ba_bz["xkts"];
                        dr1["xkxsts"] = temp_ba_bz["xkxsts"];
                        dr1["xktnjj"] = temp_ba_bz["xktnjj"];
                        dr1["bzrgts"] = temp_ba_bz["rgts"];
                        dr1["bzrgtnjj"] = temp_ba_bz["rgtnjj"];

                    }
                    else
                    {
                        dr1["xkts"] = "";
                        dr1["xkxsts"] = "";
                        dr1["xktnjj"] = "";
                        dr1["bzrgts"] = 0;
                        dr1["bzrgtnjj"] = 0;
                    }

                    #endregion

                    #region  上周成交数据
                    if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                    {

                        dr1["szbats"] = temp_cjba_sz.Sum(m => m["ts"].ints());
                        dr1["szbatnjj"] = temp_cjba_sz.Sum(n => n["tnmj"].ints())!=0?(temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(n => n["tnmj"].doubls())).je_y():0;
                    }
                    else
                    {
                        dr1["szbats"] = 0;
                        dr1["szbatnjj"] = 0;
                    }
                    #endregion

                    #region 本周成交数据
                    if (temp_ba_bz != null)
                    {
                        dr1["bzbats"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                        dr1["bzbatnjj"] = temp_cjba_bz.Sum(n => n["tnmj"].ints())!=0?(temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].doubls())).je_y():0;

                    }
                    else
                    {
                        dr1["bzbats"] = 0;
                        dr1["bzbatnjj"] = 0;
                    }
                    #endregion
                    dt.Rows.Add(dr1);
                }
            }


            return dt;
        }

        public DataTable GET_JPXM_JQDZ(DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {
                if (item.ytcs[0] == "别墅")
                {
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {

                        DataRow dr1 = dt.NewRow(); 
                        dr1[0] = item.lpcs[0]+"("+item.xfytcs[i]+")";//竞争楼盘名称

                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion

                        #region 上周认购数据
                        if (temp_ba_bz != null)
                        {
                            dr1["yh"] = temp_ba_bz["yh"];
                            dr1["yxdz"] = temp_ba_bz["yxdz"];
                            dr1["xzjtyj"] = temp_ba_bz["xzjtyj"];
                            dr1["bkfs"] = "-";

                        }
                        else
                        {
                            dr1["yh"] = "";
                            dr1["yxdz"] = "无";
                            dr1["xzjtyj"] = "--";
                            dr1["bkfs"] = "--";
                        }
                        #endregion
                        dt.Rows.Add(dr1);

                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    for (int i = 0; i < item.hxcs.Length; i++)
                    {
                        DataRow dr1 = dt.NewRow();
                        dr1[0] = item.lpcs[0] + "(" + item.hxcs[i] + ")";//竞争楼盘名称
                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion

                        #region 上周认购数据
                        if (temp_ba_bz != null)
                        {
                            dr1["yh"] = temp_ba_bz["yh"];
                            dr1["yxdz"] = temp_ba_bz["yxdz"];
                            dr1["xzjtyj"] = temp_ba_bz["xzjtyj"];
                            dr1["bkfs"] = "-";

                        }
                        else
                        {
                            dr1["yh"] = "";
                            dr1["yxdz"] = "无";
                            dr1["xzjtyj"] = "--";
                            dr1["bkfs"] = "--";
                        }
                        #endregion
                        dt.Rows.Add(dr1);

                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();
                    dr1[0] = item.lpcs[0] + "(" + item.ytcs[0] + ")";//竞争楼盘名称

                    #region 数据准备
                    //竞品业态

                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    #endregion

                    #region 上周认购数据
                    if (temp_ba_bz != null)
                    {
                        dr1["yh"] = temp_ba_bz["yh"];
                        dr1["yxdz"] = temp_ba_bz["yxdz"];
                        dr1["xzjtyj"] = temp_ba_bz["xzjtyj"];
                        dr1["bkfs"] = "-";

                    }
                    else
                    {
                        dr1["yh"] = "";
                        dr1["yxdz"] = "无";
                        dr1["xzjtyj"] = "--";
                        dr1["bkfs"] = "--";
                    }
                    #endregion
                    dt.Rows.Add(dr1);
                }
            }


            return dt;
        }
    }
}
