using Aspose.Slides;
using Calculation.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    //
    public class plus_jp_gguomei : plus_jp_base
    {
        /// <summary>
        ///  大业态循环
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns> 
        public ISlideCollection _plus_jp_guomei_1(string str, int cjbh)
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
                    dt.Columns.Add("jzgjmc");
                    dt.Columns.Add("lpmc");
                    dt.Columns.Add("yt");
                    dt.Columns.Add("bzts");
                    dt.Columns.Add("dtxsts");
                    dt.Columns.Add("xkjmjj");

                    dt.Columns.Add("szbats");
                    dt.Columns.Add("szbajmjj");
                    dt.Columns.Add("szrgts");
                    dt.Columns.Add("szrgjmjj");

                    dt.Columns.Add("bzbats");
                    dt.Columns.Add("bzbajmjj");
                    dt.Columns.Add("bzrgts");
                    dt.Columns.Add("bzrgjmjj");

                    dt.Columns.Add("thb");
                    dt.Columns.Add("jghb");
                    dt.Columns.Add("bhyy");
                    DataRow dr = dt.NewRow();
                    dr[0] = "本案";
                    dr[1] = item.lpcs[0];
                    dr[2] = item.ytcs[0];
                    #region 数据准备
                    //本周当前业态认购数据
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周当前业态备案数据
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //上周当前野田认购数据
                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //上周当前业态备案数据
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //上周本案认购数据
                    var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    #endregion

                    #region  上周认购数据
                    if (temp_ba_sz != null)
                    {

                        dr[8] = temp_ba_sz["rgts"].ints();
                        dr[9] = temp_ba_sz["rgjmjj"].ints();
                    }
                    else
                    {
                        dr[8] = 0;
                        dr[9] = 0;
                    }
                    #endregion

                    #region 本周认购数据
                    if (temp_ba_bz != null)
                    {
                        dr[3] = temp_ba_bz["xkts"]; //新开套数
                        dr[4] = temp_ba_bz["xkxsts"]; //新开销售套数
                        dr[5] = temp_ba_bz["kpjmjj"];//新开建面均价
                        dr[12] = temp_ba_bz["rgts"].ints();
                        dr[13] = temp_ba_bz["rgjmjj"].ints();
                        dr[14] = temp_ba_bz["cjtshb"];
                        dr[15] = temp_ba_bz["tnjjhb"];
                        dr[16] = temp_ba_bz["bhyy"].ToString();
                    }
                    else
                    {
                        dr[3] = ""; //新开套数
                        dr[4] = ""; //新开销售套数
                        dr[5] = "";//新开建面均价       
                        dr[12] = 0;
                        dr[13] = 0;
                        dr[14] = "-";
                        dr[15] = "-";
                        dr[16] = "-";
                    }
                    #endregion

                    #region 上周成交备案
                    if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                    {
                        dr[6] = temp_cjba_sz.Sum(m => m["ts"].ints());
                        dr[7] = (temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(m => m["jzmj"].doubls())).je_y();
                    }
                    else
                    {
                        dr[6] = 0;
                        dr[7] = 0;
                    }
                    #endregion

                    #region 本周成交备案                       
                    if (temp_cjba_bz != null && temp_cjba_bz.Count() > 0)
                    {
                        dr[10] = temp_cjba_bz.Sum(m => m["ts"].ints());
                        dr[11] = (temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(m => m["jzmj"].doubls())).je_y();
                    }
                    else
                    {
                        dr[10] = 0;
                        dr[11] = 0;
                    }
                    #endregion
                    dt.Rows.Add(dr);
                    //竞争项目
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        foreach (var item_jp in item.jpxmlb)
                        {
                            DataRow dr1 = dt.NewRow();
                            dr1[0] = item_jp.jzgjmc;//竞争格局名称
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

                            #region  上周认购数据
                            if (temp_ba_sz1 != null)
                            {

                                dr1[8] = temp_ba_sz1["rgts"].ints();
                                dr1[9] = temp_ba_sz1["rgjmjj"].ints();
                            }
                            else
                            {
                                dr1[8] = 0;
                                dr1[9] = 0;
                            }
                            #endregion

                            #region 本周认购数据
                            if (temp_ba_bz1 != null)
                            {
                                dr1[3] = temp_ba_bz1["xkts"]; //新开套数
                                dr1[4] = temp_ba_bz1["xkxsts"]; //新开销售套数
                                dr1[5] = temp_ba_bz1["kpjmjj"];//新开建面均价
                                dr1[12] = temp_ba_bz1["rgts"].ints();
                                dr1[13] = temp_ba_bz1["rgjmjj"].ints();
                                dr1[14] = temp_ba_bz1["cjtshb"];
                                dr1[15] = temp_ba_bz1["tnjjhb"];
                                dr1[16] = temp_ba_bz1["bhyy"].ToString();
                            }
                            else
                            {
                                dr1[3] = ""; //新开套数
                                dr1[4] = ""; //新开销售套数
                                dr1[5] = "";//新开建面均价       
                                dr1[12] = 0;
                                dr1[13] = 0;
                                dr1[14] = "-";
                                dr1[15] = "-";
                                dr1[16] = "-";
                            }
                            #endregion

                            #region 上周成交备案
                            if (temp_cjba_sz1 != null && temp_cjba_sz1.Count() > 0)
                            {
                                dr1[6] = temp_cjba_sz1.Sum(m => m["ts"].ints());
                                dr1[7] = (temp_cjba_sz1.Sum(m => m["cjje"].longs()) / temp_cjba_sz1.Sum(m => m["jzmj"].doubls())).je_y();
                            }
                            else
                            {
                                dr1[6] = 0;
                                dr1[7] = 0;
                            }
                            #endregion

                            #region 本周成交备案                       
                            if (temp_cjba_bz1 != null && temp_cjba_bz1.Count() > 0)
                            {
                                dr1[10] = temp_cjba_bz1.Sum(m => m["ts"].ints());
                                dr1[11] = (temp_cjba_bz1.Sum(m => m["cjje"].longs()) / temp_cjba_bz1.Sum(m => m["jzmj"].doubls())).je_y();
                            }
                            else
                            {
                                dr1[10] = 0;
                                dr1[11] = 0;
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
        public ISlideCollection _plus_jp_guomei_2(string str, int cjbh)
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
                                System.Data.DataTable dt2 = new System.Data.DataTable();
                                dt2.Columns.Add("jzgjmc");
                                dt2.Columns.Add("lpmc");
                                dt2.Columns.Add("yt");
                                dt2.Columns.Add("bzts");
                                dt2.Columns.Add("dtxsts");
                                dt2.Columns.Add("xkjmjj");

                                dt2.Columns.Add("szbats");
                                dt2.Columns.Add("szbajmjj");
                                dt2.Columns.Add("szrgts");
                                dt2.Columns.Add("szrgjmjj");

                                dt2.Columns.Add("bzbats");
                                dt2.Columns.Add("bzbajmjj");
                                dt2.Columns.Add("bzrgts");
                                dt2.Columns.Add("bzrgjmjj");

                                dt2.Columns.Add("thb");
                                dt2.Columns.Add("jghb");
                                dt2.Columns.Add("bhyy");
                                DataRow dr2 = dt2.NewRow();
                                dr2[0] = "本案";
                                dr2[1] = item.lpcs[0];
                                dr2[2] = item.xfytcs[i];
                                #region 数据准备
                                //本周当前业态认购数据
                                var temp_rgsj_bz2 = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                                //本周当前业态备案数据
                                var temp_cjba_bz2 = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                                //上周当前野田认购数据
                                var temp_rgsj_sz2 = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                                //上周当前业态备案数据
                                var temp_cjba_sz2 = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                                //上周本案认购数据
                                var temp_ba_sz2 = temp_rgsj_sz2.FirstOrDefault();
                                //本周本案认购数据
                                var temp_ba_bz2 = temp_rgsj_bz2.FirstOrDefault();
                                #endregion

                                #region  上周认购数据
                                if (temp_ba_sz2 != null)
                                {

                                    dr2[8] = temp_ba_sz2["rgts"].ints();
                                    dr2[9] = temp_ba_sz2["rgjmjj"].ints();
                                }
                                else
                                {
                                    dr2[8] = 0;
                                    dr2[9] = 0;
                                }
                                #endregion

                                #region 本周认购数据
                                if (temp_ba_bz2 != null)
                                {
                                    dr2[3] = temp_ba_bz2["xkts"]; //新开套数
                                    dr2[4] = temp_ba_bz2["xkxsts"]; //新开销售套数
                                    dr2[5] = temp_ba_bz2["kpjmjj"];//新开建面均价
                                    dr2[12] = temp_ba_bz2["rgts"].ints();
                                    dr2[13] = temp_ba_bz2["rgjmjj"].ints();
                                    dr2[14] = temp_ba_bz2["cjtshb"];
                                    dr2[15] = temp_ba_bz2["tnjjhb"];
                                    dr2[16] = temp_ba_bz2["bhyy"].ToString();
                                }
                                else
                                {
                                    dr2[3] = ""; //新开套数
                                    dr2[4] = ""; //新开销售套数
                                    dr2[5] = "";//新开建面均价       
                                    dr2[12] = 0;
                                    dr2[13] = 0;
                                    dr2[14] = "-";
                                    dr2[15] = "-";
                                    dr2[16] = "-";
                                }
                                #endregion

                                #region 上周成交备案
                                if (temp_cjba_sz2 != null && temp_cjba_sz2.Count() > 0)
                                {
                                    dr2[6] = temp_cjba_sz2.Sum(m => m["ts"].ints());
                                    dr2[7] = (temp_cjba_sz2.Sum(m => m["cjje"].longs()) / temp_cjba_sz2.Sum(m => m["jzmj"].doubls())).je_y();
                                }
                                else
                                {
                                    dr2[6] = 0;
                                    dr2[7] = 0;
                                }
                                #endregion

                                #region 本周成交备案                       
                                if (temp_cjba_bz2 != null && temp_cjba_bz2.Count() > 0)
                                {
                                    dr2[10] = temp_cjba_bz2.Sum(m => m["ts"].ints());
                                    dr2[11] = (temp_cjba_bz2.Sum(m => m["cjje"].longs()) / temp_cjba_bz2.Sum(m => m["jzmj"].doubls())).je_y();
                                }
                                else
                                {
                                    dr2[10] = 0;
                                    dr2[11] = 0;
                                }
                                #endregion
                                dt2.Rows.Add(dr2);

                                //竞争项目
                                foreach (var item_jp in item.jpxmlb)
                                {
                                    if (item_jp.xfytcs != null && item_jp.xfytcs.Length > 0)
                                    {
                                        for (int j = 0; j < item_jp.xfytcs.Length; j++)
                                        {
                                            if (item_jp.xfytcs[j] != item.xfytcs[i])
                                                continue;
                                            DataRow dr3 = dt2.NewRow();
                                            dr3[0] = item_jp.jzgjmc;
                                            dr3[1] = item_jp.lpcs[0];
                                            dr3[2] = item_jp.xfytcs[j];
                                            #region 数据准备
                                            //竞品业态
                                            //string jpyt = item_jp.xfytcs == null ? item.xfytcs[0] : item_jp.xfytcs[i];

                                            //本周当前业态认购数据
                                            var temp_rgsj_bz3 = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                                            //本周当前业态备案数据
                                            var temp_cjba_bz3 = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                                            //上周当前野田认购数据
                                            var temp_rgsj_sz3 = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                                            //上周当前业态备案数据
                                            var temp_cjba_sz3 = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                                            //上周本案认购数据
                                            var temp_ba_sz3 = temp_rgsj_sz3.FirstOrDefault();
                                            //本周本案认购数据
                                            var temp_ba_bz3 = temp_rgsj_bz3.FirstOrDefault();
                                            #endregion

                                            #region  上周认购数据
                                            if (temp_ba_sz3 != null)
                                            {

                                                dr3[8] = temp_ba_sz3["rgts"].ints();
                                                dr3[9] = temp_ba_sz3["rgjmjj"].ints();
                                            }
                                            else
                                            {
                                                dr3[8] = 0;
                                                dr3[9] = 0;
                                            }
                                            #endregion

                                            #region 本周认购数据
                                            if (temp_ba_bz3 != null)
                                            {
                                                dr3[3] = temp_ba_bz3["xkts"]; //新开套数
                                                dr3[4] = temp_ba_bz3["xkxsts"]; //新开销售套数
                                                dr3[5] = temp_ba_bz3["kpjmjj"];//新开建面均价
                                                dr3[12] = temp_ba_bz3["rgts"].ints();
                                                dr3[13] = temp_ba_bz3["rgjmjj"].ints();
                                                dr3[14] = temp_ba_bz3["cjtshb"];
                                                dr3[15] = temp_ba_bz3["tnjjhb"];
                                                dr3[16] = temp_ba_bz3["bhyy"].ToString();
                                            }
                                            else
                                            {
                                                dr3[3] = ""; //新开套数
                                                dr3[4] = ""; //新开销售套数
                                                dr3[5] = "";//新开建面均价       
                                                dr3[12] = 0;
                                                dr3[13] = 0;
                                                dr3[14] = "-";
                                                dr3[15] = "-";
                                                dr3[16] = "-";
                                            }
                                            #endregion

                                            #region 上周成交备案
                                            if (temp_cjba_sz3 != null && temp_cjba_sz3.Count() > 0)
                                            {
                                                dr3[6] = temp_cjba_sz3.Sum(m => m["ts"].ints());
                                                dr3[7] = (temp_cjba_sz3.Sum(m => m["cjje"].longs()) / temp_cjba_sz3.Sum(m => m["jzmj"].doubls())).je_y();
                                            }
                                            else
                                            {
                                                dr3[6] = 0;
                                                dr3[7] = 0;
                                            }
                                            #endregion

                                            #region 本周成交备案                       
                                            if (temp_cjba_bz3 != null && temp_cjba_bz3.Count() > 0)
                                            {
                                                dr3[10] = temp_cjba_bz3.Sum(m => m["ts"].ints());
                                                dr3[11] = (temp_cjba_bz3.Sum(m => m["cjje"].longs()) / temp_cjba_bz3.Sum(m => m["jzmj"].doubls())).je_y();
                                            }
                                            else
                                            {
                                                dr3[10] = 0;
                                                dr3[11] = 0;
                                            }
                                            #endregion
                                            dt2.Rows.Add(dr3);
                                        }
                                    }
                                    else
                                    {
                                        //这里后面来了
                                    }
                                }

                                #region 本案细分业态无值
                                //还没弄
                                #endregion
                                Office_Tables.SetJP_FD_Table(page, dt2, 2, null, null);
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
