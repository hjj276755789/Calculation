using Aspose.Slides;
using Aspose.Slides.Charts;
using Calculation.Base;
using Calculation.Models.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    public class plus_jp_fd :weak
    {
        /// <summary>
        /// 复地-竞品-竞争格局-图1
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_fudi_1(string str, int cjbh)
        {
            try
            {
        
            var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
            var t = new Presentation(str).Slides;
            t.RemoveAt(0);
            foreach (var item in param)
            {
                var temp = new Presentation(str).Slides;
                var page = temp[0];
                IAutoShape text1 = (IAutoShape)page.Shapes[2];
                text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, string.Join(",", item.ytcs));
                List<JP_BA_INFO> p = new List<JP_BA_INFO>();
                p.Add(item);
                IChart chart = (IChart)page.Shapes[3];
                var data1 = (from a in Cache_data_cjjl.bz.AsEnumerable()
                             where item.lpcs.Contains(a["lpmc"]) && item.ytcs.Contains(a["yt"])
                             group a by new { lpmc = a["lpmc"], xfyt = a["xfyt"] } into m
                             select new
                             {
                                 lpmc = m.Key.lpmc + "(" + m.Key.xfyt + ")",
                                 cjts = m.Sum(n => n["ts"].ints()),
                                 jmjj = m.Sum(n => n["cjje"].longs()) / m.Sum(n => n["jzmj"].doubls())
                             }
                           ).ToList();
                System.Data.DataTable jzgjt = new System.Data.DataTable();
                jzgjt.Columns.Add("");
                jzgjt.Columns.Add("成交套数", typeof(int));
                jzgjt.Columns.Add("建面均价", typeof(double));
                
                if (data1.Count > 0)
                {
                    DataRow dr1 = jzgjt.NewRow();
                    dr1[0] = data1[0].lpmc;
                    dr1[1] = data1[0].cjts;
                    dr1[2] = data1[0].jmjj.je_y();
                    jzgjt.Rows.Add(dr1);
                }
                else
                {
                    for (int i = 0; i < item.lpcs.Count(); i++)
                    {
                        DataRow dr1 = jzgjt.NewRow();
                        dr1[0] = item.lpcs[i];
                        dr1[1] = 0;
                        dr1[2] = 0;
                        jzgjt.Rows.Add(dr1);
                    }

                }
               

                foreach (var jpxm in item.jpxmlb)
                {
                    var data2 = (from a in Cache_data_cjjl.bz.AsEnumerable()
                                 where jpxm.lpcs.Contains(a["lpmc"]) && jpxm.ytcs.Contains(a["yt"])
                                 group a by new { lpmc = a["lpmc"], xfyt = a["xfyt"] } into m
                                 select new
                                 {
                                     lpmc = m.Key.lpmc + "(" + m.Key.xfyt + ")",
                                     cjts = m.Sum(n => n["ts"].ints()),
                                     jmjj = m.Sum(n => n["cjje"].longs()) / m.Sum(n => n["jzmj"].doubls())
                                 }
                           ).ToList();

                    if (data2.Count > 0)
                    {
                        DataRow dr1 = jzgjt.NewRow();
                        dr1[0] = data2[0].lpmc;
                        dr1[1] = data2[0].cjts;
                        dr1[2] = data2[0].jmjj.je_y();
                        jzgjt.Rows.Add(dr1);
                    }
                    else
                    {
                        for (int i = 0; i < jpxm.lpcs.Count(); i++)
                        {
                            DataRow dr1 = jzgjt.NewRow();
                            dr1[0] = jpxm.lpcs[i];
                            dr1[1] = 0;
                            dr1[2] = 0;


                            jzgjt.Rows.Add(dr1);

                        }

                    }


                   
                }
               
                Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);

                t.AddClone(temp[0]);
            }
            return t;
            }
            catch (Exception)
            {

                return null;
            }
        }
        /// <summary>
        /// 复地-竞品-竞争格局-图2
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_fudi_2(string str, int cjbh)
        {

            var t = new Presentation(str).Slides;
            t.RemoveAt(0);
            var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
            foreach (var item in param)
            {
                if (item.ytcs[0] != "别墅")
                {
                    var temp = new Presentation(str).Slides;
                    var page = temp[0];
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
                    for (int i = 0; i < item.ytcs.Count(); i++)
                    {
                        DataRow dr = dt.NewRow();
                        dr[0] = "本案";
                        dr[1] = item.lpcs[i];
                        dr[2] = item.ytcs[i];
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
                        if (temp_ba_sz!=null)
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
                        if (temp_ba_bz != null) { 
                            dr[3] = temp_ba_bz["xkts"]; //新开套数
                            dr[4] = temp_ba_bz["xkxsts"]; //新开销售套数
                            dr[5] = temp_ba_bz["kpjmjj"];//新开建面均价
                            dr[12] = temp_ba_bz["rgts"].ints();
                            dr[13] = temp_ba_bz["rgjmjj"].ints();
                            dr[14] = temp_ba_bz["cjtshb"];
                            dr[15] = temp_ba_bz["tnjjhb"];
                            dr[16] = temp_ba_bz["bhyy"].ToString();
                        }
                        else{
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
                        if (temp_cjba_sz != null&&temp_cjba_sz.Count()>0) { 
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
                        if (temp_cjba_bz != null&& temp_cjba_bz.Count() > 0)
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
                                var temp_rgsj_bz1 = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item_jp.ytcs[0]);
                                var temp_cjba_bz1 = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item_jp.ytcs[0]);

                                var temp_rgsj_sz1 = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item_jp.ytcs[0]);
                                var temp_cjba_sz1 = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item_jp.lpcs[0] && m["yt"].ToString() == item_jp.ytcs[0]);

                                //上周本案认购数据
                                var temp_ba_sz1 = temp_rgsj_sz1.FirstOrDefault();
                                //本周本案认购数据
                                var temp_ba_bz1 = temp_rgsj_bz1.FirstOrDefault();

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
                    }


                    Office_Tables.SetJP_FD_Table(page, dt,2, null,null);
                    t.AddClone(page);
                }

            }
            return t;
        }
        /// <summary>
        /// 复地-竞品-竞品近期动作
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_fudi_3(string str, int cjbh)
        {
            return new Presentation(str).Slides;
        }
    }
}
