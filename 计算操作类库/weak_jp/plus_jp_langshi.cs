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
    public class plus_jp_langshi:plus_jp_base
    {

        /// <summary>
        /// 差别十分巨大，无法重用
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_langshi_1(string str, int cjbh)
        {
            return null;
        }

        public DataTable GET_JPXM_ROW(DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {
                if (item.ytcs[0] == "别墅")
                {
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {

                        DataRow dr1 = dt.NewRow();
                        dr1[1] = item.lpcs[0];//竞争楼盘名称
                        dr1[2] = item.xfytcs[i];//竞争业态
                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                        //上周本案认购数据
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion



                        #region 本周认购数据
                        if (temp_ba_bz != null)
                        {
                            dr1["zltnmjqj"] = temp_ba_bz["zltnmjqj"];
                            dr1["xkts"] = temp_ba_bz["xkts"];
                            dr1["rgts"] = temp_ba_bz["rgts"];
                            dr1["rgtnjj"] = temp_ba_bz["rgtnjj"];
                            dr1["hd"] = temp_ba_bz["hd"];
                        }
                        else
                        {
                            dr1["zltnmjqj"] = "";
                            dr1["xkts"] = "0";
                            dr1["rgts"] = "0";
                            dr1["rgtnjj"] = 0;
                            dr1["hd"] = "-";
                        }

                        #endregion



                        #region 本周成交数据
                        if (temp_ba_bz != null)
                        {
                            dr1["bzcjts"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                            dr1["bzcjjmjj"] = temp_cjba_bz.Sum(m => m["cjje"].ints()) / temp_cjba_bz.Sum(m => m["jzmj"].ints());


                        }
                        else
                        {
                            dr1["bzcjts"] = 0;
                            dr1["bzcjjmjj"] = 0;

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



                        #region 本周认购数据
                        if (temp_ba_bz != null)
                        {
                            dr1["zltnmjqj"] = temp_ba_bz["zltnmjqj"];
                            dr1["xkts"] = temp_ba_bz["xkts"];
                            dr1["rgts"] = temp_ba_bz["rgts"];
                            dr1["rgtnjj"] = temp_ba_bz["rgtnjj"];
                            dr1["hd"] = temp_ba_bz["hd"];
                        }
                        else
                        {
                            dr1["zltnmjqj"] = "";
                            dr1["xkts"] = "0";
                            dr1["rgts"] = "0";
                            dr1["rgtnjj"] = 0;
                            dr1["hd"] = "-";
                        }

                        #endregion



                        #region 本周成交数据
                        if (temp_ba_bz != null)
                        {
                            dr1["bzcjts"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                            dr1["bzcjjmjj"] = temp_cjba_bz.Sum(m => m["cjje"].ints()) / temp_cjba_bz.Sum(m => m["jzmj"].ints());


                        }
                        else
                        {
                            dr1["bzcjts"] = 0;
                            dr1["bzcjjmjj"] = 0;

                        }
                        #endregion

                        dt.Rows.Add(dr1);
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();
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



                    #region 本周认购数据
                    if (temp_ba_bz != null)
                    {
                        dr1["zltnmjqj"] = temp_ba_bz["zltnmjqj"];
                        dr1["xkts"] = temp_ba_bz["xkts"];
                        dr1["rgts"] = temp_ba_bz["rgts"];
                        dr1["rgtnjj"] = temp_ba_bz["rgtnjj"];
                        dr1["hd"] = temp_ba_bz["hd"];
                    }
                    else
                    {
                        dr1["zltnmjqj"] = "";
                        dr1["xkts"] = "0";
                        dr1["rgts"] = "0";
                        dr1["rgtnjj"] = 0;
                        dr1["hd"] = "-";
                    }

                    #endregion



                    #region 本周成交数据
                    if (temp_ba_bz != null)
                    {
                        dr1["bzcjts"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                        dr1["bzcjjmjj"] = temp_cjba_bz.Sum(m => m["cjje"].ints()) / temp_cjba_bz.Sum(m => m["jzmj"].ints());


                    }
                    else
                    {
                        dr1["bzcjts"] = 0;
                        dr1["bzcjjmjj"] = 0;

                    }
                    #endregion

                    dt.Rows.Add(dr1);
                }
            }


            return dt;
        }
    }
}
