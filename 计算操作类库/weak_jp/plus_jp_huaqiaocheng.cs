using Aspose.Slides;
using Calculation.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS.weak_jp
{
    /// <summary>
    /// 华侨城
    /// </summary>
    public class plus_jp_huaqiaocheng :plus_jp_base
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

                #region P1 

                foreach (var item in _plus_jp_dyt_jzgj(cjbh))
                {
                    if (item != null)
                        t.AddClone(item);
                }
                #endregion
                #region P2

                foreach (var item in param)
                {
                    var tp = new Presentation(str);
                    var temp = tp.Slides;
                    var page = temp[1];
                    IAutoShape text1 = (IAutoShape)page.Shapes[4];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add("lpmc");
                    dt.Columns.Add("yt");
                    dt.Columns.Add("xkts");
                    dt.Columns.Add("xkxsts");
                    dt.Columns.Add("xktnjj");

                    dt.Columns.Add("szcjts");
                    dt.Columns.Add("szcjmj");
                    dt.Columns.Add("szcjje");
                    dt.Columns.Add("sztnjj");

                    dt.Columns.Add("bzcjts");
                    dt.Columns.Add("bzcjmj");
                    dt.Columns.Add("bzcjje");
                    dt.Columns.Add("bztnjj");
                    dt.Columns.Add("yxhd");
                    DataRow dr = dt.NewRow();
                    dr["lpmc"] = item.lpcs[0];
                    dr["yt"] = item.ytcs[0];
                    #region 数据准备
                    //本周当前业态认购数据
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周当前业态备案数据
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //上周当前野田认购数据
                    //var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //上周当前业态备案数据
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //上周本案认购数据
                    //var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    #endregion

                    #region 本周认购数据
                    if(temp_ba_bz != null)
                    {
                        dr["xkts"] =temp_ba_bz["xkts"];
                        dr["xkxsts"] = temp_ba_bz["xkxsts"];
                        dr["xktnjj"] = temp_ba_bz["xktnjj"];
                        dr["hd"] = temp_ba_bz["hd"];

                    }
                    else
                    {
                        dr["xkts"] = "";
                        dr["xkxsts"] = "";
                        dr["xktnjj"] = "";
                        dr["hd"] = "-";
                    }

                    #endregion

                    #region  上周成交数据
                    if (temp_cjba_sz != null&&temp_cjba_sz.Count()>0)
                    {

                        dr["szcjts"] = temp_cjba_sz.Sum(m=>m["ts"].ints()) ;
                        dr["szcjmj"] = temp_cjba_sz.Sum(m => m["jzmj"].ints());
                        dr["szcjje"] = temp_cjba_sz.Sum(m=>m["cjje"].longs());
                        dr["sztnjj"] = (temp_cjba_sz.Sum(m=>m["cjje"].longs())/temp_cjba_sz.Sum(n=>n["tnmj"].ints())).je_y();
                    }
                    else
                    {
                        dr["szcjts"] = 0;
                        dr["szcjmj"] = 0;
                        dr["szcjje"] = 0;
                        dr["sztnjj"] = 0;
                    }
                    #endregion

                    #region 本周成交数据
                    if (temp_ba_bz != null)
                    {
                        dr["bzcjts"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                        dr["bzcjmj"] = temp_cjba_bz.Sum(m => m["jzmj"].ints());
                        dr["bzcjje"] = temp_cjba_bz.Sum(m => m["cjje"].longs());
                        dr["bztnjj"] = (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].ints())).je_y();

                    }
                    else
                    {
                        dr["bzcjts"] = 0;
                        dr["bzcjmj"] = 0;
                        dr["bzcjje"] = 0;
                        dr["bztnjj"] = 0;
                    }
                    #endregion

                    
                    dt.Rows.Add(dr);
                    //竞争项目
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        foreach (var item_jp in item.jpxmlb)
                        {
                            DataRow dr1 = dt.NewRow();
                            dr1[1] = item_jp.lpcs[0];//竞争楼盘名称
                            dr1[2] = item.ytcs[0];//竞争业态
                            #region 数据准备
                            //竞品业态
                            string jpyt = item_jp.ytcs == null ? item.ytcs[0] : item_jp.ytcs[0];

                            var temp_rgsj_bz1 = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == jpyt);
                            var temp_cjba_bz1 = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == jpyt);

                            var temp_rgsj_sz1 = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == jpyt);
                            var temp_cjba_sz1 = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == jpyt);

                            //上周本案认购数据
                            var temp_ba_sz1 = temp_rgsj_sz1.FirstOrDefault();
                            //本周本案认购数据
                            var temp_ba_bz1 = temp_rgsj_bz1.FirstOrDefault();
                            #endregion



                            #region 本周认购数据
                            if (temp_ba_bz1 != null)
                            {
                                dr1["xkts"] = temp_ba_bz1["xkts"];
                                dr1["xkxsts"] = temp_ba_bz1["xkxsts"];
                                dr1["xktnjj"] = temp_ba_bz1["xktnjj"];
                                dr1["hd"] = temp_ba_bz1["hd"];

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

                                dr1["szcjts"] = temp_cjba_sz1.Sum(m => m["ts"].ints());
                                dr1["szcjmj"] = temp_cjba_sz1.Sum(m => m["jzmj"].ints());
                                dr1["szcjje"] = temp_cjba_sz1.Sum(m => m["cjje"].longs());
                                dr1["sztnjj"] = (temp_cjba_sz1.Sum(m => m["cjje"].longs()) / temp_cjba_sz1.Sum(n => n["tnmj"].ints())).je_y();
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
                            if (temp_ba_bz1 != null)
                            {
                                dr1["bzcjts"] = temp_cjba_bz1.Sum(m => m["ts"].ints());
                                dr1["bzcjmj"] = temp_cjba_bz1.Sum(m => m["jzmj"].ints());
                                dr1["bzcjje"] = temp_cjba_bz1.Sum(m => m["cjje"].longs());
                                dr1["bztnjj"] = (temp_cjba_bz1.Sum(m => m["cjje"].longs()) / temp_cjba_bz1.Sum(n => n["tnmj"].ints())).je_y();

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

        /// <summary>
        /// 细分业态循环
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_huaqiaocheng_2(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);

                #region P1 


                foreach (var item in _plus_jp_xfyt_jzgj(cjbh))
                {
                    if (item != null)
                        t.AddClone(item);
                }
                #endregion
                #region P2

                foreach (var item in param)
                {
                    if (item.ytcs[0] == "别墅" || item.ytcs[0] == "商务")
                    {
                        #region 本案细分业态有值
                        if (item.xfytcs != null && item.xfytcs.Length > 0)
                        {

                            //添加本案数据
                            for (int i = 0; i < item.xfytcs.Count(); i++)
                            {
                                var temp = new Presentation(str).Slides;
                                var page = temp[1];
                                IAutoShape text1 = (IAutoShape)page.Shapes[4];
                                text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.xfytcs[i]);
                                System.Data.DataTable dt = new System.Data.DataTable();
                                dt.Columns.Add("lpmc");
                                dt.Columns.Add("yt");
                                dt.Columns.Add("xkts");
                                dt.Columns.Add("xkxsts");
                                dt.Columns.Add("xktnjj");

                                dt.Columns.Add("szcjts");
                                dt.Columns.Add("szcjmj");
                                dt.Columns.Add("szcjje");
                                dt.Columns.Add("sztnjj");

                                dt.Columns.Add("bzcjts");
                                dt.Columns.Add("bzcjmj");
                                dt.Columns.Add("bzcjje");
                                dt.Columns.Add("bztnjj");
                                dt.Columns.Add("yxhd");
                                DataRow dr = dt.NewRow();
                                dr["lpmc"] = item.lpcs[0];
                                dr["yt"] = item.xfytcs[i];


                                #region 数据准备
                                //本周当前业态认购数据
                                var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                                //本周当前业态备案数据
                                var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);         
                                //上周当前业态备案数据
                                var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                                //本周本案认购数据
                                var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                                #endregion


                                #region 本周认购数据
                                if (temp_ba_bz != null)
                                {
                                    dr["xkts"] = temp_ba_bz["xkts"];
                                    dr["xkxsts"] = temp_ba_bz["xkxsts"];
                                    dr["xktnjj"] = temp_ba_bz["xktnjj"];
                                    dr["hd"] = temp_ba_bz["hd"];

                                }
                                else
                                {
                                    dr["xkts"] = "";
                                    dr["xkxsts"] = "";
                                    dr["xktnjj"] = "";
                                    dr["hd"] = "-";
                                }

                                #endregion

                                #region  上周成交数据
                                if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                                {

                                    dr["szcjts"] = temp_cjba_sz.Sum(m => m["ts"].ints());
                                    dr["szcjmj"] = temp_cjba_sz.Sum(m => m["jzmj"].ints());
                                    dr["szcjje"] = temp_cjba_sz.Sum(m => m["cjje"].longs());
                                    dr["sztnjj"] = (temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(n => n["tnmj"].ints())).je_y();
                                }
                                else
                                {
                                    dr["szcjts"] = 0;
                                    dr["szcjmj"] = 0;
                                    dr["szcjje"] = 0;
                                    dr["sztnjj"] = 0;
                                }
                                #endregion

                                #region 本周成交数据
                                if (temp_ba_bz != null)
                                {
                                    dr["bzcjts"] = temp_cjba_bz.Sum(m => m["ts"].ints());
                                    dr["bzcjmj"] = temp_cjba_bz.Sum(m => m["jzmj"].ints());
                                    dr["bzcjje"] = temp_cjba_bz.Sum(m => m["cjje"].longs());
                                    dr["bztnjj"] = (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(n => n["tnmj"].ints())).je_y();

                                }
                                else
                                {
                                    dr["bzcjts"] = 0;
                                    dr["bzcjmj"] = 0;
                                    dr["bzcjje"] = 0;
                                    dr["bztnjj"] = 0;
                                }
                                #endregion
                                dt.Rows.Add(dr);

                                #region 竞争项目

                               
                                foreach (var item_jp in item.jpxmlb)
                                {
                                    if (item_jp.xfytcs != null && item_jp.xfytcs.Length > 0)
                                    {
                                        for (int j = 0; j < item_jp.xfytcs.Length; j++)
                                        {
                                            if (item_jp.xfytcs[j] != item.xfytcs[i])
                                                continue;
                                            DataRow dr1 = dt.NewRow();
                                            dr1[0] = item_jp.jzgjmc;
                                            dr1[1] = item_jp.lpcs[0];
                                            dr1[2] = item_jp.xfytcs[j];
                                            #region 数据准备
                                            //竞品业态
                                            //string jpyt = item_jp.xfytcs == null ? item.xfytcs[0] : item_jp.xfytcs[i];

                                            //本周当前业态认购数据
                                            var temp_rgsj_bz1 = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                                            //本周当前业态备案数据
                                            var temp_cjba_bz1 = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);                         
                                            //上周当前业态备案数据
                                            var temp_cjba_sz1 = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                                            //本周本案认购数据
                                            var temp_ba_bz1 = temp_rgsj_bz1.FirstOrDefault();
                                            #endregion


                                            #region 本周认购数据
                                            if (temp_ba_bz1 != null)
                                            {
                                                dr1["xkts"] = temp_ba_bz1["xkts"];
                                                dr1["xkxsts"] = temp_ba_bz1["xkxsts"];
                                                dr1["xktnjj"] = temp_ba_bz1["xktnjj"];
                                                dr1["hd"] = temp_ba_bz1["hd"];

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

                                                dr1["szcjts"] = temp_cjba_sz1.Sum(m => m["ts"].ints());
                                                dr1["szcjmj"] = temp_cjba_sz1.Sum(m => m["jzmj"].ints());
                                                dr1["szcjje"] = temp_cjba_sz1.Sum(m => m["cjje"].longs());
                                                dr1["sztnjj"] = (temp_cjba_sz1.Sum(m => m["cjje"].longs()) / temp_cjba_sz1.Sum(n => n["tnmj"].ints())).je_y();
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
                                            if (temp_ba_bz1 != null)
                                            {
                                                dr1["bzcjts"] = temp_cjba_bz1.Sum(m => m["ts"].ints());
                                                dr1["bzcjmj"] = temp_cjba_bz1.Sum(m => m["jzmj"].ints());
                                                dr1["bzcjje"] = temp_cjba_bz1.Sum(m => m["cjje"].longs());
                                                dr1["bztnjj"] = (temp_cjba_bz1.Sum(m => m["cjje"].longs()) / temp_cjba_bz1.Sum(n => n["tnmj"].ints())).je_y();

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
                                        //这里后面来了
                                    }
                                }
                                #endregion

                                #region 本案细分业态无值
                                //还没弄
                                #endregion
                                Office_Tables.SetJP_FD_Table(page, dt, 2, null, null);
                                t.AddClone(page);
                            }

                            #endregion

                        }


                    }
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
    }
}
