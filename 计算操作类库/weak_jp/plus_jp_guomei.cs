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
    public class plus_jp_guomei : plus_jp_base
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
                    var page = temp[0];
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
                    dt.Columns.Add("bhyy");

                    #region 商务

                    
                    if (item.ytcs[0] == "商务")
                    {
                        if (item.hxcs != null && item.hxcs.Length > 0)
                        {
                            for (int i = 0; i < item.hxcs.Length; i++)
                            {

                                DataRow dr = dt.NewRow();
                                dr[0] = "本案";
                                dr[1] = item.lpcs[0];
                                dr[2] = item.ytcs[0];
                                #region 数据准备
                                //本周当前业态认购数据
                                var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                                //本周当前业态备案数据
                                var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
                                //上周当前野田认购数据
                                var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                                //上周当前业态备案数据
                                var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["hx"].ToString() == item.hxcs[i]);
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
                                    dr[5] = temp_ba_bz["xkjmjj"];//新开建面均价
                                    dr[12] = temp_ba_bz["rgts"].ints();
                                    dr[13] = temp_ba_bz["rgjmjj"].ints();
                                    dr[14] = temp_ba_bz["bhyy"].ToString();
                                }
                                else
                                {
                                    dr[3] = ""; //新开套数
                                    dr[4] = ""; //新开销售套数
                                    dr[5] = "";//新开建面均价       
                                    dr[12] = 0;
                                    dr[13] = 0;
                                    dr[14] = "-";
                                }
                                #endregion

                                #region 上周成交备案
                                if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                                {
                                    dr[6] = temp_cjba_sz.Sum(m => m["ts"].ints());
                                    dr[7] = temp_cjba_sz.Sum(m => m["tnmj"].doubls())!=0?(temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(m => m["tnmj"].doubls())).je_y():0;
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
                                    dr[11] = temp_cjba_bz.Sum(m => m["tnmj"].doubls())!=0?(temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(m => m["tnmj"].doubls())).je_y():0;
                                }
                                else
                                {
                                    dr[10] = 0;
                                    dr[11] = 0;
                                }
                                #endregion
                                dt.Rows.Add(dr);
                            }
                        }

                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            foreach (var item_jp in item.jpxmlb)
                            {
                                if (item_jp.hxcs != null && item_jp.hxcs.Length > 0)
                                {
                                    for (int i = 0; i < item_jp.hxcs.Length; i++)
                                    {

                                        DataRow dr1 = dt.NewRow();
                                        dr1[0] = item_jp.jzgjmc;//竞争格局名称
                                        dr1[1] = item_jp.lpcs[0];//竞争楼盘名称
                                        dr1[2] = item_jp.hxcs[i];//竞争业态
                                        #region 数据准备
                                        //竞品业态

                                        var temp_rgsj_bz1 = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item_jp.hxcs[i]);
                                        var temp_cjba_bz1 = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["hx"].ToString() == item_jp.hxcs[i]);

                                        var temp_rgsj_sz1 = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item_jp.hxcs[i]);
                                        var temp_cjba_sz1 = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["hx"].ToString() == item_jp.hxcs[i]);

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
                                            dr1[5] = temp_ba_bz1["xkjmjj"];//新开建面均价
                                            dr1[12] = temp_ba_bz1["rgts"].ints();
                                            dr1[13] = temp_ba_bz1["rgjmjj"].ints();
                                            dr1[14] = temp_ba_bz1["bhyy"].ToString();
                                        }
                                        else
                                        {
                                            dr1[3] = ""; //新开套数
                                            dr1[4] = ""; //新开销售套数
                                            dr1[5] = "";//新开建面均价       
                                            dr1[12] = 0;
                                            dr1[13] = 0;
                                            dr1[14] = "-";
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
                                else
                                {

                                }

                            }
                        }
                        Office_Tables.SetJP_FD_Table(page, dt, 2, null, null);
                        t.AddClone(page);

                    }
                    #endregion

                    #region 其他
                    else
                    {
                        //若有细分业态
                        if (item.xfytcs != null && item.xfytcs.Length > 0)
                        {
                            for (int i = 0; i < item.xfytcs.Length; i++)
                            {
                                DataRow dr = dt.NewRow();
                                dr[0] = "本案";
                                dr[1] = item.lpcs[0];
                                dr[2] = item.ytcs[0];
                                #region 数据准备
                                //本周当前业态认购数据
                                var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[0]);
                                //本周当前业态备案数据
                                var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[0]);
                                //上周当前野田认购数据
                                var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[0]);
                                //上周当前业态备案数据
                                var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[0]);
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
                                    dr[5] = temp_ba_bz["xkjmjj"];//新开建面均价
                                    dr[12] = temp_ba_bz["rgts"].ints();
                                    dr[13] = temp_ba_bz["rgjmjj"].ints();

                                    dr[14] = temp_ba_bz["bhyy"].ToString();
                                }
                                else
                                {
                                    dr[3] = ""; //新开套数
                                    dr[4] = ""; //新开销售套数
                                    dr[5] = "";//新开建面均价       
                                    dr[12] = 0;
                                    dr[13] = 0;

                                    dr[14] = "-";
                                }
                                #endregion

                                #region 上周成交备案
                                if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                                {
                                    dr[6] = temp_cjba_sz.Sum(m => m["ts"].ints());
                                    dr[7] = temp_cjba_sz.Sum(m => m["tnmj"].doubls())!=0? (temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(m => m["tnmj"].doubls())).je_y():0;
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
                                    dr[11] = temp_cjba_bz.Sum(m => m["tnmj"].doubls())!=0?(temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(m => m["tnmj"].doubls())).je_y():0;
                                }
                                else
                                {
                                    dr[10] = 0;
                                    dr[11] = 0;
                                }
                                #endregion
                                dt.Rows.Add(dr);
                            }
                        }
                        //无细分业态
                        else
                        {
                            #region MyRegion

                            DataRow dr = dt.NewRow();
                            dr[0] = "本案";
                            dr[1] = item.lpcs[0];
                            dr[2] = item.ytcs[0];
                            #region 数据准备
                            //本周当前业态认购数据
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                            //本周当前业态备案数据
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                            //上周当前野田认购数据
                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
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
                                dr[5] = temp_ba_bz["xkjmjj"];//新开建面均价
                                dr[12] = temp_ba_bz["rgts"].ints();
                                dr[13] = temp_ba_bz["rgjmjj"].ints();

                                dr[14] = temp_ba_bz["bhyy"].ToString();
                            }
                            else
                            {
                                dr[3] = ""; //新开套数
                                dr[4] = ""; //新开销售套数
                                dr[5] = "";//新开建面均价       
                                dr[12] = 0;
                                dr[13] = 0;

                                dr[14] = "-";
                            }
                            #endregion

                            #region 上周成交备案
                            if (temp_cjba_sz != null && temp_cjba_sz.Count() > 0)
                            {
                                dr[6] = temp_cjba_sz.Sum(m => m["ts"].ints());
                                dr[7] = temp_cjba_sz.Sum(m => m["tnmj"].doubls())!=0?(temp_cjba_sz.Sum(m => m["cjje"].longs()) / temp_cjba_sz.Sum(m => m["tnmj"].doubls())).je_y():0;
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
                                dr[11] = temp_cjba_bz.Sum(m => m["tnmj"].doubls())!=0?(temp_cjba_bz.Sum(m => m["cjje"].longs()) / temp_cjba_bz.Sum(m => m["tnmj"].doubls())).je_y():0;
                            }
                            else
                            {
                                dr[10] = 0;
                                dr[11] = 0;
                            }
                            #endregion
                            dt.Rows.Add(dr);
                            #endregion
                        }
                        //竞争项目
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            foreach (var item_jp in item.jpxmlb)
                            {
                                if (item_jp.xfytcs != null && item_jp.xfytcs.Length > 0)
                                {
                                    for (int i = 0; i < item_jp.xfytcs.Length; i++)
                                    {
                                        #region MyRegion
                                        DataRow dr1 = dt.NewRow();
                                        dr1[0] = item_jp.jzgjmc;//竞争格局名称
                                        dr1[1] = item_jp.lpcs[0];//竞争楼盘名称
                                        dr1[2] = item_jp.xfytcs[i];//竞争业态
                                        #region 数据准备

                                        var temp_rgsj_bz1 = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item_jp.xfytcs[i]);
                                        var temp_cjba_bz1 = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["xfyt"].ToString() == item_jp.xfytcs[i]);

                                        var temp_rgsj_sz1 = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item_jp.xfytcs[i]);
                                        var temp_cjba_sz1 = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["xfyt"].ToString() == item_jp.xfytcs[i]);

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
                                            dr1[5] = temp_ba_bz1["xkjmjj"];//新开建面均价
                                            dr1[12] = temp_ba_bz1["rgts"].ints();
                                            dr1[13] = temp_ba_bz1["rgjmjj"].ints();

                                            dr1[14] = temp_ba_bz1["bhyy"].ToString();
                                        }
                                        else
                                        {
                                            dr1[3] = ""; //新开套数
                                            dr1[4] = ""; //新开销售套数
                                            dr1[5] = "";//新开建面均价       
                                            dr1[12] = 0;
                                            dr1[13] = 0;

                                            dr1[14] = "-";
                                        }
                                        #endregion

                                        #region 上周成交备案
                                        if (temp_cjba_sz1 != null && temp_cjba_sz1.Count() > 0)
                                        {
                                            dr1[6] = temp_cjba_sz1.Sum(m => m["ts"].ints());
                                            dr1[7] = temp_cjba_sz1.Sum(m => m["tnmj"].doubls())!=0?(temp_cjba_sz1.Sum(m => m["cjje"].longs()) / temp_cjba_sz1.Sum(m => m["tnmj"].doubls())).je_y():0;
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
                                            dr1[11] = temp_cjba_bz1.Sum(m => m["tnmj"].doubls())!=0?(temp_cjba_bz1.Sum(m => m["cjje"].longs()) / temp_cjba_bz1.Sum(m => m["tnmj"].doubls())).je_y():0;
                                        }
                                        else
                                        {
                                            dr1[10] = 0;
                                            dr1[11] = 0;
                                        }
                                        #endregion

                                        dt.Rows.Add(dr1);
                                        #endregion
                                    }
                                }
                                else
                                {
                                    if (item_jp.ytcs != null && item_jp.ytcs.Length > 0)
                                    {
                                        #region MyRegion
                                        DataRow dr1 = dt.NewRow();
                                        dr1[0] = item_jp.jzgjmc;//竞争格局名称
                                        dr1[1] = item_jp.lpcs[0];//竞争楼盘名称
                                        dr1[2] = item_jp.ytcs[0];//竞争业态
                                        #region 数据准备
                                        string jpyt = item_jp.ytcs == null ? item.ytcs[0] : item_jp.ytcs[0];

                                        var temp_rgsj_bz1 = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == jpyt);
                                        var temp_cjba_bz1 = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == jpyt);

                                        var temp_rgsj_sz1 = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == jpyt);
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
                                            dr1[5] = temp_ba_bz1["xkjmjj"];//新开建面均价
                                            dr1[12] = temp_ba_bz1["rgts"].ints();
                                            dr1[13] = temp_ba_bz1["rgjmjj"].ints();

                                            dr1[14] = temp_ba_bz1["bhyy"].ToString();
                                        }
                                        else
                                        {
                                            dr1[3] = ""; //新开套数
                                            dr1[4] = ""; //新开销售套数
                                            dr1[5] = "";//新开建面均价       
                                            dr1[12] = 0;
                                            dr1[13] = 0;

                                            dr1[14] = "-";
                                        }
                                        #endregion

                                        #region 上周成交备案
                                        if (temp_cjba_sz1 != null && temp_cjba_sz1.Count() > 0)
                                        {
                                            dr1[6] = temp_cjba_sz1.Sum(m => m["ts"].ints());
                                            dr1[7] = temp_cjba_sz1.Sum(m => m["tnmj"].doubls())!=0?(temp_cjba_sz1.Sum(m => m["cjje"].longs()) / temp_cjba_sz1.Sum(m => m["tnmj"].doubls())).je_y():0;
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
                                            dr1[11] = temp_cjba_bz1.Sum(m => m["tnmj"].doubls())!=0?(temp_cjba_bz1.Sum(m => m["cjje"].longs()) / temp_cjba_bz1.Sum(m => m["tnmj"].doubls())).je_y():0;
                                        }
                                        else
                                        {
                                            dr1[10] = 0;
                                            dr1[11] = 0;
                                        }
                                        #endregion

                                        dt.Rows.Add(dr1);
                                        #endregion
                                    }
                                    else
                                    {
                                        Base_Log.Log("没有输入业态参数");
                                        continue;
                                    }
                                }

                            }
                        }
                        Office_Tables.SetJP_FD_Table(page, dt, 2, null, null);
                        t.AddClone(page);
                    }
                }
                #endregion



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


    }
}
