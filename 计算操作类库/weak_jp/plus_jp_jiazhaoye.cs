﻿using Aspose.Slides;
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
    class plus_jp_jiazhaoye:plus_jp_base
    {
        /// <summary>
        ///  大业态循环
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns> 
        public ISlideCollection _plus_jp_huaqiaocheng_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);

              
                #region P2

                foreach (var item in param)
                {
                    var tp = new Presentation(str);
                    var temp = tp.Slides;
                    var page = temp[1];
                    IAutoShape text1 = (IAutoShape)page.Shapes[4];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add("jzgj");
                    dt.Columns.Add("lpmc");
                    dt.Columns.Add("yt");
                    dt.Columns.Add("xkts");
                    dt.Columns.Add("xkxsts");
                    dt.Columns.Add("xktnjj");

                    dt.Columns.Add("szcjts"); //上周成交数据
                    dt.Columns.Add("szcjtnjj"); //上周成交套内均价
                    dt.Columns.Add("szcjjmjj"); //上周成交建面均价

                    dt.Columns.Add("szrgts");   //上周认购套数
                    dt.Columns.Add("szrgtnjj"); //上周认购套内均价
                    dt.Columns.Add("szrgjmjj"); //上周认购建面均价


                    dt.Columns.Add("bzcjts");   //本周成交套数
                    dt.Columns.Add("bzcjtnjj"); //本周成交套内均价
                    dt.Columns.Add("szcjjmjj"); //本周成交建面均价

                    dt.Columns.Add("bzrgts"); //本周认购数据
                    dt.Columns.Add("bzrgtnjj"); //本周认购套内均价
                    dt.Columns.Add("bzrgjmjj"); //本周建面均价

                    dt.Columns.Add("tshb");  //认购环比
                    dt.Columns.Add("jghb");  //价格环比
                    dt.Columns.Add("bhyy");  //变化原因
                    dt.Columns.Add("bz");    //下周加推预计

                    if(item.jpxmlb!=null&&item.jpxmlb.Count>0)
                    {
                        dt = GET_JPXM_ROW(dt, item.jpxmlb);
                    }
                  
                    Office_Tables.SetJP_FD_Table(page, dt, 2, null, null);
                    t.AddClone(page);
                }




                #endregion
                #region P3

                foreach (var item in _plus_jp_dyt_tgtp(cjbh))
                {
                    if (item != null)
                        t.AddClone(item);
                }

                #endregion
                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log(e.Message);
                return null;
            }
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

                                var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                                var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                                //上周本案认购数据
                                var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                                //本周本案认购数据
                                var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                                #endregion



                                #region 本周认购数据
                                if (temp_ba_bz != null)
                                {
                                    dr1["xkts"] = temp_ba_bz["xkts"];
                                    dr1["xkxsts"] = temp_ba_bz["xkxsts"];
                                    dr1["xktnjj"] = temp_ba_bz["xktnjj"];
                                    dr1["hd"] = temp_ba_bz["hd"];

                                }
                                else
                                {
                                    dr1["xkts"] = "";
                                    dr1["xkxsts"] = "";
                                    dr1["xktnjj"] = "";
                                    dr1["hd"] = "-";
                                }

                                #endregion

                                #region  上周成交数据
                                if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                                {

                                    dr1["szcjts"] = temp_cjba_sz.Sum(m => m["ts"].ints());
                                    dr1["szcjmj"] = temp_cjba_sz.Sum(m => m["jzmj"].ints());
                                    dr1["szcjje"] = temp_cjba_sz.Sum(m => m["cjje"].longs());
                                    dr1["sztnjj"] = (temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(n => n["tnmj"].ints())).je_y();
                                }
                                else
                                {
                                    dr1["szcjts"] = 0;
                                    dr1["szcjmj"] = 0;
                                    dr1["szcjje"] = 0;
                                    dr1["sztnjj"] = 0;
                                }
                                #endregion

                                #region 本周成交数据
                                if (temp_ba_bz != null)
                                {
                                    dr1["bzcjts"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                                    dr1["bzcjmj"] = temp_cjba_bz.Sum(m => m["jzmj"].ints());
                                    dr1["bzcjje"] = temp_cjba_bz.Sum(m => m["cjje"].longs());
                                    dr1["bztnjj"] = (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].ints())).je_y();

                                }
                                else
                                {
                                    dr1["bzcjts"] = 0;
                                    dr1["bzcjmj"] = 0;
                                    dr1["bzcjje"] = 0;
                                    dr1["bztnjj"] = 0;
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
                            dr1["xkts"] = temp_ba_bz["xkts"];
                            dr1["xkxsts"] = temp_ba_bz["xkxsts"];
                            dr1["xktnjj"] = temp_ba_bz["xktnjj"];
                            dr1["hd"] = temp_ba_bz["hd"];

                        }
                        else
                        {
                            dr1["xkts"] = "";
                            dr1["xkxsts"] = "";
                            dr1["xktnjj"] = "";
                            dr1["hd"] = "-";
                        }

                        #endregion

                        #region  上周成交数据
                        if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                        {

                            dr1["szcjts"] = temp_cjba_sz.Sum(m => m["ts"].ints());
                            dr1["szcjmj"] = temp_cjba_sz.Sum(m => m["jzmj"].ints());
                            dr1["szcjje"] = temp_cjba_sz.Sum(m => m["cjje"].longs());
                            dr1["sztnjj"] = (temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(n => n["tnmj"].ints())).je_y();
                        }
                        else
                        {
                            dr1["szcjts"] = 0;
                            dr1["szcjmj"] = 0;
                            dr1["szcjje"] = 0;
                            dr1["sztnjj"] = 0;
                        }
                        #endregion

                        #region 本周成交数据
                        if (temp_ba_bz != null)
                        {
                            dr1["bzcjts"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                            dr1["bzcjmj"] = temp_cjba_bz.Sum(m => m["jzmj"].ints());
                            dr1["bzcjje"] = temp_cjba_bz.Sum(m => m["cjje"].longs());
                            dr1["bztnjj"] = (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].ints())).je_y();

                        }
                        else
                        {
                            dr1["bzcjts"] = 0;
                            dr1["bzcjmj"] = 0;
                            dr1["bzcjje"] = 0;
                            dr1["bztnjj"] = 0;
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
                        dr1["xkts"] = temp_ba_bz["xkts"];
                        dr1["xkxsts"] = temp_ba_bz["xkxsts"];
                        dr1["xktnjj"] = temp_ba_bz["xktnjj"];
                        dr1["hd"] = temp_ba_bz["hd"];

                    }
                    else
                    {
                        dr1["xkts"] = "";
                        dr1["xkxsts"] = "";
                        dr1["xktnjj"] = "";
                        dr1["hd"] = "-";
                    }

                    #endregion

                    #region  上周成交数据
                    if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                    {

                        dr1["szcjts"] = temp_cjba_sz.Sum(m => m["ts"].ints());
                        dr1["szcjmj"] = temp_cjba_sz.Sum(m => m["jzmj"].ints());
                        dr1["szcjje"] = temp_cjba_sz.Sum(m => m["cjje"].longs());
                        dr1["sztnjj"] = (temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(n => n["tnmj"].ints())).je_y();
                    }
                    else
                    {
                        dr1["szcjts"] = 0;
                        dr1["szcjmj"] = 0;
                        dr1["szcjje"] = 0;
                        dr1["sztnjj"] = 0;
                    }
                    #endregion

                    #region 本周成交数据
                    if (temp_ba_bz != null)
                    {
                        dr1["bzcjts"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                        dr1["bzcjmj"] = temp_cjba_bz.Sum(m => m["jzmj"].ints());
                        dr1["bzcjje"] = temp_cjba_bz.Sum(m => m["cjje"].longs());
                        dr1["bztnjj"] = (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].ints())).je_y();

                    }
                    else
                    {
                        dr1["bzcjts"] = 0;
                        dr1["bzcjmj"] = 0;
                        dr1["bzcjje"] = 0;
                        dr1["bztnjj"] = 0;
                    }
                    #endregion

                    dt.Rows.Add(dr1);
                }
            }


            return dt ;
        }
    }
}
