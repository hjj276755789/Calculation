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
            //return new Presentation(str).Slides;
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
               

                //var data2 = (from a in Cache_data_cjjl.bz.AsEnumerable()
                //            where item.lpcs.Contains(a["lpmc"]) && item.ytcs.Contains(a["yt"])
                //            group a by new { lpmc = a["lpmc"], xfyt = a["xfyt"] } into m
                //            select new
                //            {
                //                lpmc = m.Key.lpmc + "(" + m.Key.xfyt + ")",
                //                cjts = m.Sum(n => n["ts"].ints()),
                //                jmjj = m.Sum(n => n["cjje"].longs()) / m.Sum(n => n["jzmj"].doubls())
                //            }
                //           ).ToList();
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
                //var data = (from a in p
                //            join b in Cache_data_cjjl.bz.AsEnumerable() on new {lpmc= a.lpcs,xfyt= a.xfytcs }  equals new { lpmc= b["lpmc"].ToString(),xfyt =b["xfyt"].ToString()}
                //            group a by new { lpmc = a["lpmc"], xfyt = a["xfyt"] } into m
                //            select new
                //            {
                //                lpmc = m.Key.lpmc + "(" + m.Key.xfyt + ")",
                //                cjts = m.Sum(n => n["cjts"].ints()),
                //                jmjj = m.Sum(n => n["cjje"].longs()) / m.Sum(n => n["jzmj"].doubls())
                //            }
                //           ).ToList();

                //var temp6 = (from a in swsc_cj
                //             join b in swsc_gy on a.zc equals b.zc into temp
                //             from tt in temp.DefaultIfEmpty()
                //             select new
                //             {
                //                 zcmc = a.zcmc,
                //                 xzgyl = tt == null ? 0 : tt.xzgyl,//这里主要第二个集合有可能为空。需要判断
                //                 cjmj = a.cjmj,
                //                 jmjj = a.cjje / a.cjmj
                //             }).ToList();


                //System.Data.DataTable gxfx_dt = new System.Data.DataTable();
                //gxfx_dt.Columns.Add("");
                //gxfx_dt.Columns.Add("成交套数", typeof(int));
                //gxfx_dt.Columns.Add("建面均价", typeof(double));

                //var gxfx_gy = data1.OrderBy(m => m.cjts).ToList();
                //for (int i1 = 0; i1 < gxfx_gy.Count(); i1++)
                //{
                //    DataRow dr = gxfx_dt.NewRow();
                //    dr[0] = gxfx_gy[i1].lpmc;
                //    dr[1] = gxfx_gy[i1].cjts;
                //    dr[2] = gxfx_gy[i1].jmjj.je_y();
                    
                //    gxfx_dt.Rows.Add(dr);
                //}
                Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);

                t.AddClone(temp[0]);
            }
            return t;
        }
        /// <summary>
        /// 复地-竞品-竞争格局-图2
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_fudi_2(string str, int cjbh)
        {
            return new Presentation(str).Slides;
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
