using Aspose.Slides;
using Aspose.Slides.Charts;
using Calculation.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculation.JS
{
    public class weak
    {
        
        public weak()
        {
            Cache_data_cjjl.ini_zb();
            Cache_data_tdjyjl.ini_zb();
            Cache_data_xzys.ini_zb();
            Cache_Result_zb.ini();
           // Cache_param_zb.ini_zb();
        }

        #region 铭腾版插件

        
        public ISlideCollection plus1(string str, int cjbh)
        {
            return new Presentation(str).Slides;
        }
        #region 插件2
        public ISlideCollection plus2(string str, int cjbh)
        {
            var t = new Presentation(str).Slides;
            string str1 =
                string.Format("本周供应{0}万方，成交面积{1}万方，成交价格{2}元/㎡，[总结]",
                Cache_Result_zb.bz_cj_jzmj_xzys.mj_wf(),
                Cache_Result_zb.bz_cj_jzmj.mj_wf(),
                (Cache_Result_zb.bz_cj_cjje / Cache_Result_zb.bz_cj_jzmj).je_y()
                );
            //住宅
            var bz_gy_czf = Cache_data_xzys.bz.AsEnumerable().Where(m => m["tyyt"].ToString() == "别墅" || m["tyyt"].ToString() == "高层" || m["tyyt"].ToString() == "小高层" || m["tyyt"].ToString() == "洋房" || m["tyyt"].ToString() == "洋楼");
            var bz_cj_czf = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["yt"].ToString() == "别墅" || m["yt"].ToString() == "高层" || m["yt"].ToString() == "小高层" || m["yt"].ToString() == "洋房" || m["yt"].ToString() == "洋楼");

            string str2 = string.Format("本周供应{0}万方，成交面积{1}万方，均价{2}元/㎡，[总结]",
                bz_gy_czf.Sum(m =>m["jzmj"].doubls()).mj_wf(),
                bz_cj_czf.Sum(m =>m["jzmj"].doubls()).mj_wf(),
                (bz_cj_czf.Sum(m => m["cjje"].longs()) / bz_cj_czf.Sum(m =>m["jzmj"].doubls())).je_y()
                );
            //商务
            var bz_gy_sw = Cache_data_xzys.bz.AsEnumerable().Where(m => m["tyyt"].ToString() == "商务");
            var bz_cj_sw = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["yt"].ToString() == "商务");

            string str3 = string.Format("本周新增供应{0}万方，成交面积{1}万方，均价{2}元/㎡，[总结]",
               bz_gy_sw.Sum(m =>m["jzmj"].doubls()).mj_wf(),
                bz_cj_sw.Sum(m =>m["jzmj"].doubls()).mj_wf(),
                (bz_cj_sw.Sum(m => m["cjje"].longs()) / bz_cj_sw.Sum(m =>m["jzmj"].doubls())).je_y()
                );
            //商业
            var bz_gy_sy = Cache_data_xzys.bz.AsEnumerable().Where(m => m["tyyt"].ToString() == "商业");
            var bz_cj_sy = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["yt"].ToString() == "商铺");

            string str4 = string.Format("本周供应{0}万方，成交面积{1}万方，均价{2}元/㎡，[总结]",
               bz_gy_sy.Sum(m =>m["jzmj"].doubls()).mj_wf(),
                bz_cj_sy.Sum(m =>m["jzmj"].doubls()).mj_wf(),
                (bz_cj_sy.Sum(m => m["cjje"].longs()) / bz_cj_sy.Sum(m =>m["jzmj"].doubls())).je_y()
                );
            //商业
            string str5 = "【自填】，[总结]。热销项目为：[调整]";

            IAutoShape text1 = (IAutoShape)t[0].Shapes[4];
            IAutoShape text2 = (IAutoShape)t[0].Shapes[7];
            IAutoShape text3 = (IAutoShape)t[0].Shapes[10];
            IAutoShape text4 = (IAutoShape)t[0].Shapes[13];
            IAutoShape text5 = (IAutoShape)t[0].Shapes[16];
            text1.TextFrame.Text = str1;
            text2.TextFrame.Text = str2;
            text3.TextFrame.Text = str3;
            text4.TextFrame.Text = str4;
            text5.TextFrame.Text = str5;

            Office_StyleHelper.setdefaultstyle(text1);
            Office_StyleHelper.setdefaultstyle(text2);
            Office_StyleHelper.setdefaultstyle(text3);
            Office_StyleHelper.setdefaultstyle(text4);
            Office_StyleHelper.setdefaultstyle(text5);


            return t;
        }
        #endregion
        #region 插件3
        public ISlideCollection plus3(string str, int cjbh)
        {
            var t = new Presentation(str).Slides;
            string str1 = string.Format(@"本周（{0}）土地成交{1}宗 [总结]",Base_date.bzwz,Cache_Result_zb.td_bz_cjsl);
            string str2 = string.Format(@"本周（{0}），成交{1}宗，土地面积{2}亩，可建体量{3}万方，成交总金额{4}亿元。 [总结]",
                Base_date.bzwz,
                Cache_Result_zb.td_bz_cjsl,
                Cache_Result_zb.td_bz_zyd.mj_m(),
                Cache_Result_zb.td_bz_kjtl.mj(),
                Cache_Result_zb.td_bz_cjje.je_yy());
            IAutoShape text1 = (IAutoShape)t[1].Shapes[4];
            IAutoShape text2 = (IAutoShape)t[1].Shapes[5];
            text1.TextFrame.Text = str1;
            text2.TextFrame.Text = str2;

            Office_ChartStyle style = new Office_ChartStyle();
            style.坐标方向 = Base_Config.坐标方向.纵向;
            style.文字位置 = Aspose.Slides.Charts.LegendDataLabelPosition.OutsideEnd;
            style.文字旋转方向 = TextVerticalType.Horizontal;
            style.是否显示文字 = true;
            Office_Charts.DoubleAxexchart(t[2], Cache_Result_zb.jsjg_xkpzs, 5,1,2);

            var cjts = from a in Cache_data_cjjl.bz.AsEnumerable()
                       group a by new { xmmc = a.Field<string>("lpmc"), zt = a.Field<string>("zt") } into m
                       select new
                       {
                           xmmc = m.Key.xmmc,
                           zt = m.Key.zt,
                           cjts = m.Count()
                       };
            var cjmj = from a in Cache_data_cjjl.bz.AsEnumerable()
                       group a by new { xmmc = a.Field<string>("lpmc"), zt = a.Field<string>("zt") } into m
                       select new
                       {
                           xmmc = m.Key.xmmc,
                           zt = m.Key.zt,
                           cjmj = m.Sum(n=> double.Parse(n["jzmj"].ToString()))
                       };
            var cjje = from a in Cache_data_cjjl.bz.AsEnumerable()
                       group a by new { xmmc = a.Field<string>("lpmc"), zt = a.Field<string>("zt") } into m
                       select new
                       {
                           xmmc = m.Key.xmmc,
                           zt = m.Key.zt,
                           cjje = m.Sum(n => double.Parse(n["cjje"].ToString()))
                       };
            var cjts1 = cjts.OrderByDescending(m => m.cjts).Take(10).ToList();
            var cjmj1 = cjmj.OrderByDescending(m => m.cjmj).Take(10).ToList();
            var cjje1 = cjje.OrderByDescending(m => m.cjje).Take(10).ToList();
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("xh");
            dt.Columns.Add("xmmc1");
            dt.Columns.Add("zt1");
            dt.Columns.Add("cjts");
            dt.Columns.Add("xmmc2");
            dt.Columns.Add("zt2");
            dt.Columns.Add("cjmj");
            dt.Columns.Add("xmmc3");
            dt.Columns.Add("zt3");
            dt.Columns.Add("cjje");
            for (int i = 0; i < 10; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = i+1;
                dr[1] = cjts1[i].xmmc;
                dr[2] = cjts1[i].zt;
                dr[3] = cjts1[i].cjts;

                dr[4] = cjmj1[i].xmmc;
                dr[5] = cjmj1[i].zt;
                dr[6] = cjmj1[i].cjmj.mj();

                dr[7] = cjje1[i].xmmc;
                dr[8] = cjje1[i].zt;
                dr[9] = cjje1[i].cjje.je_wy();
                dt.Rows.Add(dr);
            }

            IAutoShape test = (IAutoShape)t[3].Shapes[2];
            test.TextFrame.Text = string.Format("本周成交排名（{0}", Base_date.bzwz);
            Office_Tables.SetChart(t[3], dt, 4, null,10);
            return t;
        }
        #endregion
        #region 插件4
        public ISlideCollection plus4(string str, int cjbh)
        {
            ISlideCollection t = new Presentation(str).Slides;

            #region 商品房成交情况
            #region 第一页 供需分析
            var s1 = t[1];
            //图表等下做
            var dt1_1 = from a in Cache_data_xzys.jbz.AsEnumerable()
                      group a by new { zc = a["zc"],zcmc=a["zcmc"] } into s
                      select new
                      {
                          zc=s.Key.zc,
                          zcmc= s.Key.zcmc,
                          xzgyl = s.Sum(m => m["jzmj"].doubls()+  m["fzzmj"].doubls()).mj_wf() 
                      };
            var dt1_2 = from a in Cache_data_cjjl.jbz.AsEnumerable()
                        group a by new { zc = a["zc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            cjmj = s.Sum(m =>m["jzmj"].doubls()).mj_wf()
                      };
            System.Data.DataTable gxfx_dt = new System.Data.DataTable();
            gxfx_dt.Columns.Add("时间");
            gxfx_dt.Columns.Add("预售新增供应量（万㎡）", typeof(double));
            gxfx_dt.Columns.Add("成交建面（万㎡）", typeof(double));
            gxfx_dt.Columns.Add("批售比", typeof(double));

            var gxfx_gy = dt1_1.OrderBy(m => int.Parse(m.zc.ToString())).ToList();
            var gxfx_cj = dt1_2.OrderBy(m => int.Parse(m.zc.ToString())).ToList();
            for (int i1 = 0; i1 < gxfx_gy.Count(); i1++)
            {
                DataRow dr = gxfx_dt.NewRow();
                dr[0] = gxfx_gy[i1].zcmc;
                dr[1] = gxfx_gy[i1].xzgyl;
                dr[2] = gxfx_cj[i1].cjmj;
                dr[3] = (gxfx_gy[i1].xzgyl / gxfx_cj[i1].cjmj).ss_bfb_ys();
                gxfx_dt.Rows.Add(dr);
            }
           Office_Charts.Chart_gxfx(s1, gxfx_dt, 5);
            //文字
            IAutoShape test = (IAutoShape)s1.Shapes[6];
            var ckjjsyf = Cache_data_cjjl.bz.AsEnumerable().Where(m => (m["yt"].ToString() == "车库" || m["yt"].ToString() == "经济适用房")).ToList();
            test.TextFrame.Paragraphs[0].Text = string.Format("本周新增供应{0}，环比{1}；[总结]", Cache_Result_zb.bz_cj_jzmj_xzys.mj_wf(),
                ((Cache_Result_zb.bz_cj_jzmj_xzys - Cache_Result_zb.sz_cj_jzmj_xzys) / Cache_Result_zb.sz_cj_jzmj_xzys).ss_bfb());
            test.TextFrame.Paragraphs[2].Text = string.Format("本周成交面积{0}，环比{1}；[总结]", Cache_Result_zb.bz_cj_jzmj.mj_wf(),
                ((Cache_Result_zb.bz_cj_jzmj - Cache_Result_zb.sz_cj_jzmj) / Cache_Result_zb.sz_cj_jzmj).ss_bfb());

            test.TextFrame.Paragraphs[4].Text = string.Format("本周成交{0}套，环比{1}，其中车库及经适房{2}套，占比{3}，体量占比{4}；",
                Cache_data_cjjl.bz.Rows.Count,
                ((Cache_data_cjjl.bz.Rows.Count - Cache_data_cjjl.sz.Rows.Count) / (double)Cache_data_cjjl.sz.Rows.Count).ss_bfb(),
                ckjjsyf.Count,
                (ckjjsyf.Count / double.Parse(Cache_data_cjjl.bz.Rows.Count.ToString())).ss_bfb_jdz(),
                (ckjjsyf.Sum(m =>m["jzmj"].doubls()) / Cache_Result_zb.bz_cj_jzmj).ss_bfb_jdz()
                );

            test.TextFrame.Paragraphs[6].Text = "[总结]";
            #endregion

            #region 第二页 量价关系
            var s2 = t[2];
            IChart c2 = (IChart)s2.Shapes[6];
            var dt2_1 = from a in Cache_data_cjjl.jbz.AsEnumerable()
                    group a by new { zc = a["zc"]} into s
                    select new
                    {
                        zc = s.Key.zc,
                        cjje = s.Sum(a=>a["cjje"].longs()),
                        jzmj = s.Sum(a => double.Parse(a["jzmj"].ToString())),
                    };
            System.Data.DataTable ljgx_dt = new System.Data.DataTable();
            ljgx_dt.Columns.Add("时间");
            ljgx_dt.Columns.Add("成交金额（亿元）", typeof(long));
            ljgx_dt.Columns.Add("成交建面（万㎡）", typeof(double));
            ljgx_dt.Columns.Add("建面均价", typeof(string));
            ///量价关系
            var ljgx_gx = dt2_1.OrderBy(m => int.Parse(m.zc.ToString())).ToList();
            for (int i1 = 0; i1 < ljgx_gx.Count(); i1++)
            {
                DataRow dr = ljgx_dt.NewRow();
                dr[0] = ljgx_gx[i1].zc;
                dr[1] = ljgx_gx[i1].cjje.je_yy();
                dr[2] = ljgx_gx[i1].jzmj.mj_wf();
                dr[3] = (ljgx_gx[i1].cjje / ljgx_gx[i1].jzmj).je_y();
                ljgx_dt.Rows.Add(dr);
            }
            Office_Charts.Chart_gxfx(s2, ljgx_dt, 6);

            IAutoShape text2 = (IAutoShape)s2.Shapes[3];
            text2.TextFrame.Paragraphs[0].Text = string.Format("本周成交面积{0}万方，环比{1}，比去年同期{2}，同比{3}",
                Cache_Result_zb.bz_cj_jzmj.mj_wf(),
                ((Cache_Result_zb.bz_cj_jzmj - Cache_Result_zb.sz_cj_jzmj) / Cache_Result_zb.sz_cj_jzmj).ss_bfb(),
                (Cache_Result_zb.bz_cj_jzmj - Cache_Result_zb.tz_cj_jzmj).mj_wf_ms(),
                ((Cache_Result_zb.bz_cj_jzmj - Cache_Result_zb.tz_cj_jzmj) / Cache_Result_zb.tz_cj_jzmj).ss_bfb());

            text2.TextFrame.Paragraphs[2].Text = string.Format("本周成交均价{0}元/㎡，环比{1}，比去年同期{2}元/㎡，同比{3}", 
                (Cache_Result_zb.bz_cj_cjje / Cache_Result_zb.bz_cj_jzmj).je_y(),
                (((Cache_Result_zb.bz_cj_cjje / Cache_Result_zb.bz_cj_jzmj) - (Cache_Result_zb.sz_cj_cjje / Cache_Result_zb.sz_cj_jzmj)) / (Cache_Result_zb.sz_cj_cjje / Cache_Result_zb.sz_cj_jzmj)).ss_bfb(),
                (Cache_Result_zb.bz_cj_cjje / Cache_Result_zb.bz_cj_jzmj) - (Cache_Result_zb.tz_cj_cjje / Cache_Result_zb.tz_cj_jzmj),
                (((Cache_Result_zb.bz_cj_cjje / Cache_Result_zb.bz_cj_jzmj) - (Cache_Result_zb.tz_cj_cjje / Cache_Result_zb.tz_cj_jzmj)) / (Cache_Result_zb.tz_cj_cjje / Cache_Result_zb.tz_cj_jzmj)).ss_bfb()
                );

            text2.TextFrame.Paragraphs[4].Text = string.Format("本周成交金额{0}亿元，环比{1}％，【总结】，成交金额同比{2}",
                Cache_Result_zb.bz_cj_cjje.je_yy(),
                ((Cache_Result_zb.bz_cj_cjje - Cache_Result_zb.sz_cj_cjje) / Cache_Result_zb.sz_cj_cjje).ss_bfb(),
                ((Cache_Result_zb.bz_cj_cjje - Cache_Result_zb.tz_cj_cjje) / Cache_Result_zb.tz_cj_cjje).ss_bfb()
                );



            #endregion

            #region 第三页 成交区域对比
             var s3 = t[3];
            IChart c3 = (IChart)s3.Shapes[5];
            System.Data.DataTable dt1 = Cache_data_cjjl.bz;
            var cjqy = from a in Cache_data_cjjl.bz.AsEnumerable()
                       group a by new { qymc = a.Field<string>("qy") } into m
                       select new
                       {
                           qymc = m.Key.qymc,
                           jzmj = m.Sum(a=> double.Parse(a.Field<double>("jzmj").ToString())).mj_wf(),
                           cjje = m.Sum(a=>a.Field<object>("cjje").longs()),
                           jmjj = (m.Sum(a => a.Field<object>("cjje").longs()) / m.Sum(a => double.Parse(a.Field<double>("jzmj").ToString()))).je_y() 
                       };
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("qymc");
            dt.Columns.Add("成交面积",typeof(double));
            dt.Columns.Add("成交金额",typeof(long));
            dt.Columns.Add("建面均价",typeof(double));
            dt.Columns.Add("市场均价",typeof(double));
            var cjqy1 = cjqy.OrderByDescending(m=>m.jzmj).ToList();
            for (int i1 = 0; i1 < cjqy1.Count(); i1++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = cjqy1[i1].qymc;
                dr[1] = cjqy1[i1].jzmj;
                dr[2] = cjqy1[i1].cjje.je_yy();
                dr[3] = cjqy1[i1].jmjj;
                dr[4] = (Cache_Result_zb.bz_cj_cjje / Cache_Result_zb.bz_cj_jzmj).je_y();
                dt.Rows.Add(dr);
            }
            Office_Charts.ThreeWchart(s3, dt,5);
            IAutoShape text3 = (IAutoShape)s3.Shapes[6];
            double bz_cjmj_zl = cjqy1.Sum(m => m.jzmj);
            text3.TextFrame.Text =string.Format( "本周成交主力为{0}，其成交总量占全市成交{1}。其次为{2}、{3}、{4}、{5}，成交量占全市成交的{6}、{7}、{8}、{9}。",
                cjqy1[0].qymc, (cjqy1[0].jzmj/ bz_cjmj_zl).ss_bfb_jdz(), cjqy1[1].qymc, cjqy1[2].qymc, cjqy1[3].qymc, cjqy1[4].qymc, 
                (cjqy1[1].jzmj / bz_cjmj_zl).ss_bfb_jdz(), (cjqy1[2].jzmj / bz_cjmj_zl).ss_bfb_jdz(), (cjqy1[3].jzmj / bz_cjmj_zl).ss_bfb_jdz(), (cjqy1[4].jzmj / bz_cjmj_zl).ss_bfb_jdz()
                );

            #endregion
            #endregion

            #region 住宅成交情况
            #region 第四页 住宅市场
            var s4 = t[4];

           
            //图表等下做
            var dt4_1 = from a in Cache_data_xzys.jbz.AsEnumerable()
                        where a["tyyt"].ToString() == "别墅" || a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "小高层" || a["tyyt"].ToString() == "洋房" || a["tyyt"].ToString() == "洋楼"
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            zcmc = s.Key.zcmc,
                            xzgyl = s.Sum(m =>m["jzmj"].doubls()).mj_wf()
                        };
            var dt4_2 = from a in Cache_data_cjjl.jbz.AsEnumerable()
                        where a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼"
                        group a by new { zc = a["zc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            cjje =s.Sum(m => m["cjje"].longs()),
                            cjmj = s.Sum(m =>m["jzmj"].doubls())
                        };

            System.Data.DataTable zzsc_dt = new System.Data.DataTable();
            zzsc_dt.Columns.Add("时间");
            zzsc_dt.Columns.Add("预售新增供应量（万㎡）", typeof(double));
            zzsc_dt.Columns.Add("成交建面（万㎡）", typeof(double));
            zzsc_dt.Columns.Add("建面均价", typeof(double));

            var zzsc_gy = dt4_1.OrderBy(m => int.Parse(m.zc.ToString())).ToList();
            var zzsc_cj = dt4_2.OrderBy(m => int.Parse(m.zc.ToString())).ToList();
            for (int i1 = 0; i1 < zzsc_gy.Count(); i1++)
            {
                DataRow dr = zzsc_dt.NewRow();
                dr[0] = zzsc_gy[i1].zcmc;
                dr[1] = zzsc_gy[i1].xzgyl;
                dr[2] = zzsc_cj[i1].cjmj.mj_wf();
                dr[3] = (zzsc_cj[i1].cjje / zzsc_cj[i1].cjmj).je_y();
                zzsc_dt.Rows.Add(dr);
            }
            Office_Charts.Chart_gxfx(s4, zzsc_dt, 5);
            IAutoShape text4 = (IAutoShape)s4.Shapes[6];

            double bz_zz_xzgy = Cache_Result_zb.bz_cj_czz_xzys.Sum(m =>m["jzmj"].doubls());
            double sz_zz_xzgy = Cache_Result_zb.sz_cj_czz_xzys.Sum(m =>m["jzmj"].doubls());
            double bz_zz_cjmj = Cache_Result_zb.bz_cj_czz.Sum(m =>m["jzmj"].doubls());
            double sz_zz_cjmj = Cache_Result_zb.sz_cj_czz.Sum(m =>m["jzmj"].doubls());
            double bz_zz_cjje = Cache_Result_zb.bz_cj_czz.Sum(m => m["cjje"].longs());
            double sz_zz_cjje = Cache_Result_zb.sz_cj_czz.Sum(m => m["cjje"].longs());

            text4.TextFrame.Paragraphs[0].Text = string.Format("本周新增供应{0}万方，环比{1}，[总结]",
                bz_zz_xzgy.mj_wf(), ((bz_zz_xzgy - sz_zz_xzgy) / sz_zz_xzgy).ss_bfb());
            text4.TextFrame.Paragraphs[2].Text = string.Format("本周成交面积{0}万方，环比{1}，[总结]",
                bz_zz_cjmj.mj_wf(), ((bz_zz_cjmj - sz_zz_cjmj) / sz_zz_cjmj).ss_bfb());
            text4.TextFrame.Paragraphs[4].Text = string.Format("本周成交均价{0}元/㎡，环比{1}，[总结]",
                (bz_zz_cjje / bz_zz_cjmj).je_y(), (((bz_zz_cjje / bz_zz_cjmj) - (sz_zz_cjje / sz_zz_cjmj))  / (sz_zz_cjje / sz_zz_cjmj)).ss_bfb());
            #endregion

            #region 第五页 住宅市场
            var s5 = t[5];
            //住宅成交区域
            var zz_cjqy = from a in Cache_Result_zb.bz_cj_czz.AsEnumerable()
                          where a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼"
                          group a by new { qymc = a["qy"] } into m
                       select new
                       {
                           qymc = m.Key.qymc,
                           jzmj = m.Sum(a => a["jzmj"].doubls()),
                           cjje = m.Sum(a => a["cjje"].longs()),
                           jmjj = m.Sum(a => a["cjje"].longs()) / m.Sum(a => a["jzmj"].doubls())
                       };
            System.Data.DataTable zz_dt = new System.Data.DataTable();
            zz_dt.Columns.Add("qymc");
            zz_dt.Columns.Add("成交面积", typeof(double));
            zz_dt.Columns.Add("成交金额", typeof(double));
            zz_dt.Columns.Add("建面均价", typeof(double));
            zz_dt.Columns.Add("市场均价", typeof(double));
            var zz_cjqy1 = zz_cjqy.OrderByDescending(m => m.jzmj).ToList();
            for (int i1 = 0; i1 < zz_cjqy1.Count(); i1++)
            {
                DataRow dr = zz_dt.NewRow();
                dr[0] = zz_cjqy1[i1].qymc;
                dr[1] = zz_cjqy1[i1].jzmj.mj_wf();
                dr[2] = zz_cjqy1[i1].cjje.je_yy();
                dr[3] = zz_cjqy1[i1].jmjj.je_y();
                dr[4] = (Cache_Result_zb.bz_cj_cjje / Cache_Result_zb.bz_cj_jzmj).je_y();
                zz_dt.Rows.Add(dr);
            }
            Office_Charts.ThreeWchart(s5, zz_dt, 0);
            IAutoShape text5 = (IAutoShape)s5.Shapes[3];
            double bz_zz_cjmj_zl = zz_cjqy1.Sum(m => m.jzmj);
            text5.TextFrame.Text = string.Format("本周成交主力为{0}，其成交总量占全市成交{1}。其次为{2}、{3}、{4}、{5}，成交量占全市成交的{6}、{7}、{8}、{9}。",
                zz_cjqy1[0].qymc, (zz_cjqy1[0].jzmj / bz_zz_cjmj_zl).ss_bfb_jdz(), zz_cjqy1[1].qymc, zz_cjqy1[2].qymc, zz_cjqy1[3].qymc, zz_cjqy1[4].qymc,
                (zz_cjqy1[1].jzmj / bz_zz_cjmj_zl).ss_bfb_jdz(), (zz_cjqy1[2].jzmj / bz_zz_cjmj_zl).ss_bfb_jdz(), (zz_cjqy1[3].jzmj / bz_zz_cjmj_zl).ss_bfb_jdz(), (zz_cjqy1[4].jzmj / bz_zz_cjmj_zl).ss_bfb_jdz());
            #endregion
            #endregion

            #region 商务成交情况

            #region 第六页 
            var s6 = t[6];
            IChart c6 = (IChart)s6.Shapes[5];
            var dt6_1 = from a in Cache_data_xzys.jbz.AsEnumerable()
                        where a["tyyt"].ToString() == "商务" 
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            zcmc = s.Key.zcmc,
                            xzgyl = s.Sum(m => m["fzzmj"].doubls())
                        };
            var dt6_2 = from a in Cache_data_cjjl.jbz.AsEnumerable()
                        where a["yt"].ToString() == "商务"
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            zcmc=s.Key.zcmc,
                            cjje = s.Sum(m => m["cjje"].longs()),
                            cjmj = s.Sum(m =>m["jzmj"].doubls())
                        };

            System.Data.DataTable swsc_dt = new System.Data.DataTable();
            swsc_dt.Columns.Add("时间");
            swsc_dt.Columns.Add("预售新增供应量（万㎡）", typeof(double));
            swsc_dt.Columns.Add("成交建面（万㎡）", typeof(double));
            swsc_dt.Columns.Add("建面均价", typeof(double));

            var swsc_gy = dt6_1.OrderBy(m =>m.zc.ints()).ToList();
            var swsc_cj = dt6_2.OrderBy(m => m.zc.ints()).ToList();

            var temp6 = (from a in swsc_cj
                    join b in swsc_gy on a.zc equals b.zc into temp
                              from tt in temp.DefaultIfEmpty()
                              select new
                              {
                                  zcmc = a.zcmc,
                                  xzgyl = tt == null ? 0 : tt.xzgyl,//这里主要第二个集合有可能为空。需要判断
                                  cjmj =a.cjmj,
                                  jmjj = a.cjje /a.cjmj
                              }).ToList();


            for (int i1 = 0; i1 < temp6.Count(); i1++)
            {
                DataRow dr = swsc_dt.NewRow();
                dr[0] = temp6[i1].zcmc;
                dr[1] = temp6[i1].xzgyl.mj_wf();
                dr[2] = temp6[i1].cjmj.mj_wf();
                dr[3] = temp6[i1].jmjj.je_y();
                swsc_dt.Rows.Add(dr);
            }
            Office_Charts.Chart_gxfx(s6, swsc_dt, 5);




            IAutoShape text6 = (IAutoShape)s6.Shapes[6];
            var sw_bz_xzys = Cache_data_xzys.bz.AsEnumerable().Where(m => m["tyyt"].ToString() == "商务");
            var sw_bz_cjjl = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["yt"].ToString() == "商务");
            var sw_sz_cjjl = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["yt"].ToString() == "商务");

            double sw_bz_jzmj = sw_bz_cjjl.Sum(m =>m["jzmj"].doubls());
            double sw_sz_jzmj = sw_sz_cjjl.Sum(m =>m["jzmj"].doubls());

            double sw_bz_cjjj = sw_bz_cjjl.Sum(m => m["cjje"].longs()) / sw_bz_cjjl.Sum(m =>m["jzmj"].doubls());
            double sw_sz_cjjj = sw_sz_cjjl.Sum(m => m["cjje"].longs()) / sw_sz_cjjl.Sum(m =>m["jzmj"].doubls());

            text6.TextFrame.Paragraphs[0].Text = string.Format("本周新增供应{0}万方，[总结]", sw_bz_xzys.Sum(m =>m["fzzmj"].doubls().mj_wf()));
            text6.TextFrame.Paragraphs[2].Text = string.Format("本周成交面积{0}万方，环比{1}，[总结]",
                sw_bz_cjjl.Sum(m =>m["jzmj"].doubls()).mj_wf(),
                ((sw_bz_jzmj - sw_sz_jzmj) / sw_sz_jzmj).ss_bfb());
            text6.TextFrame.Paragraphs[4].Text = string.Format("本周成交均价{0}元/㎡，环比{1}，[总结]",
                sw_bz_cjjj.je_y(),
                ((sw_bz_cjjj - sw_sz_cjjj) / sw_sz_cjjj).ss_bfb()
                );
            #endregion

            #region 第七页 
            var s7 = t[7];

            IChart c7 = (IChart)s7.Shapes[5];
            var dt7_1_1 = from a in Cache_data_xzys.jbz.AsEnumerable()
                          where a["tyyt"].ToString() == "别墅" || a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "小高层" || a["tyyt"].ToString() == "洋房" || a["tyyt"].ToString() == "洋楼"
                          group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            zcmc = s.Key.zcmc,
                            xzgyl = s.Sum(m => m["fzzmj"].doubls())
                        };
            var dt7_1_2 = from a in Cache_data_xzys.jbz.AsEnumerable()
                        where a["tyyt"].ToString() == "商业"
                        group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            zcmc = s.Key.zcmc,
                            xzgyl = s.Sum(m => m["fzzmj"].doubls())
                        };
            var dt7_2 = from a in Cache_data_cjjl.jbz.AsEnumerable()
                        where a["yt"].ToString() == "商铺"
                        group a by new { zc = a["zc"] } into s
                        select new
                        {
                            zc = s.Key.zc,
                            cjje = s.Sum(m => m["cjje"].longs()),
                            cjmj = s.Sum(m => m["jzmj"].doubls())
                        };

            System.Data.DataTable sysc_dt = new System.Data.DataTable();
            sysc_dt.Columns.Add("时间");
            sysc_dt.Columns.Add("预售新增供应量（万㎡）", typeof(double));
            sysc_dt.Columns.Add("成交建面（万㎡）", typeof(double));
            sysc_dt.Columns.Add("建面均价", typeof(double));

            var sysc_gy_1 = dt7_1_1.OrderBy(m => m.zc.ints()).ToList();
            var sysc_gy_2 = dt7_1_2.OrderBy(m => m.zc.ints()).ToList();
            var sysc_cj = dt7_2.OrderBy(m => m.zc.ints()).ToList();
            for (int i1 = 0; i1 < sysc_gy_1.Count(); i1++)
            {
                DataRow dr = sysc_dt.NewRow();
                dr[0] = sysc_gy_1[i1].zcmc;
                dr[1] = (sysc_gy_1[i1].xzgyl + sysc_gy_2[i1].xzgyl).mj_wf();
                dr[2] = sysc_cj[i1].cjmj.mj_wf();
                dr[3] = (sysc_cj[i1].cjje / sysc_cj[i1].cjmj).je_y();
                sysc_dt.Rows.Add(dr);
            }
            Office_Charts.Chart_gxfx(s7, sysc_dt, 5);




            IAutoShape text7 = (IAutoShape)s7.Shapes[6];


            var sy_bz_xzys1 = Cache_data_xzys.bz.AsEnumerable().Where(m => m["tyyt"].ToString() == "商业");
            var sy_bz_xzys2 = Cache_data_xzys.bz.AsEnumerable().Where(a => a["tyyt"].ToString() == "别墅" || a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "小高层" || a["tyyt"].ToString() == "洋房" || a["tyyt"].ToString() == "洋楼");
            var sy_sz_xzys1 = Cache_data_xzys.sz.AsEnumerable().Where(m => m["tyyt"].ToString() == "商业");
            var sy_sz_xzys2 = Cache_data_xzys.sz.AsEnumerable().Where(a => a["tyyt"].ToString() == "别墅" || a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "小高层" || a["tyyt"].ToString() == "洋房" || a["tyyt"].ToString() == "洋楼");

            var sy_bz_cjjl = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["yt"].ToString() == "商铺");
            var sy_sz_cjjl = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["yt"].ToString() == "商铺");

            double sy_bz_jzmj = sy_bz_cjjl.Sum(m =>m["jzmj"].doubls());
            double sy_sz_jzmj = sy_sz_cjjl.Sum(m =>m["jzmj"].doubls());

            double sy_bz_cjjj = sy_bz_cjjl.Sum(m => m["cjje"].longs()) / sy_bz_cjjl.Sum(m =>m["jzmj"].doubls());
            double sy_sz_cjjj = sy_sz_cjjl.Sum(m => m["cjje"].longs()) / sy_sz_cjjl.Sum(m =>m["jzmj"].doubls());

            text7.TextFrame.Paragraphs[0].Text = string.Format("本周新增供应{0}万方,环比{1}，[总结]",
                (sy_bz_xzys1.Sum(m =>m["fzzmj"].doubls()) + sy_bz_xzys2.Sum(m => m["fzzmj"].doubls())).mj_wf(),
                
                ((sy_bz_xzys1.Sum(m => m["fzzmj"].doubls()) + sy_bz_xzys2.Sum(m => m["fzzmj"].doubls())- sy_sz_xzys1.Sum(m => m["fzzmj"].doubls()) - sy_sz_xzys2.Sum(m => m["fzzmj"].doubls())) / (sy_sz_xzys1.Sum(m => m["fzzmj"].doubls()) + sy_sz_xzys2.Sum(m => m["fzzmj"].doubls()))).ss_bfb());
            text7.TextFrame.Paragraphs[2].Text = string.Format("本周成交面积{0}万方，环比{1}，[总结]",
                sy_bz_cjjl.Sum(m =>m["jzmj"].doubls()).mj_wf(),
                ((sy_bz_jzmj - sy_sz_jzmj) / sy_sz_jzmj).ss_bfb());
            text7.TextFrame.Paragraphs[4].Text = string.Format("本周成交均价{0}元/㎡，环比{1}，[总结]",
                sy_bz_cjjj.je_y(),
                ((sy_bz_cjjj - sy_sz_cjjj) / sy_sz_cjjj).ss_bfb()
                );
            #endregion
            #endregion


            return t;
           

        }
        #endregion

        public ISlideCollection plus5(string str, int cjbh)
        {
            var p = from a in Cache_param_zb.value
                    where a.cjid == cjbh
                    select new { a.csnr };
            var path = p.FirstOrDefault();
            if (path != null && !string.IsNullOrEmpty(path.csnr))
                return new Presentation(path.csnr).Slides;
            else
                return null;
        }
        public ISlideCollection plus6(string str, int cjbh)
        {
            var p = from a in Cache_param_zb.value
                    where a.cjid == cjbh
                    select new { a.csnr };
            var path = p.FirstOrDefault();
            if (path != null && !string.IsNullOrEmpty(path.csnr))
                return new Presentation(path.csnr).Slides;
            else
                return null;
        }
        public ISlideCollection plus7(string str, int cjbh)
        {
            return new Presentation(str).Slides;
        }
        /// <summary>
        /// 进8周板块走势
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection plus8(string str,int cjbh)
        {

            #region 准备阶段          
            var s = new Presentation(str).Slides;

            var p = from a in Cache_param_zb.value
                    where a.cjid == cjbh
                    select new{ a.csnr};
            string bkmc = "";
            foreach (var item in p)
            {
                bkmc += item.csnr + "、";
            }
            if (bkmc.Length > 0)
                bkmc = bkmc.Substring(0, bkmc.Length - 1);

            var zc = from a in Cache_data_xzys.jbz.AsEnumerable()
                     group a by new { zcmc = a["zcmc"], zc = a["zc"] } into d
                     select new{zc = d.Key.zc,zcmc = d.Key.zcmc};
            #endregion

            #region 1P


            ///记号 替换板块参数
            var data = from a in Cache_data_cjjl.bz.AsEnumerable()
                       join b in p on a["zt"].ToString() equals b.csnr /*&& (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")*/
                       group a by new { xmmc = a["lpmc"] } into d
                       select new
                       {
                           d.Key.xmmc,
                           cjts =d.Count(),
                           cjje = d.Sum(m=>m["cjje"].longs()),
                           jzmj = d.Sum(m => m["jzmj"].doubls()),
                           tnmj = d.Sum(m =>m["tnmj"].doubls()),
                       };

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("排名");
            dt.Columns.Add("项目名称");
            dt.Columns.Add("成交套数");
            dt.Columns.Add("成交金额（万元）");
            dt.Columns.Add("建面体量（平方米）");
            dt.Columns.Add("套内体量（平方米）");
            dt.Columns.Add("建面均价（元/㎡）");
            dt.Columns.Add("套内均价（元/㎡）");
            var list = data.OrderByDescending(m => m.cjje).Take(10).ToList();
            for (int i = 0; i < list.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = i + 1;
                dr[1] = list[i].xmmc;
                dr[2] = list[i].cjts;
                dr[3] = list[i].cjje.je_wy();

                dr[4] = list[i].jzmj.mj();
                dr[5] = list[i].tnmj.mj();

                dr[6] = (list[i].cjje / list[i].jzmj).je_y();

                dr[7] = (list[i].cjje / list[i].tnmj).je_y();

                dt.Rows.Add(dr);
            }

            Office_Tables.SetChart(s[0], dt, 0, null, 10);
            IAutoShape text1 = (IAutoShape)s[0].Shapes[3];
           
            text1.TextFrame.Text =string.Format("{0}板块住宅项目本周成交排名",bkmc);
            #endregion

            #region 2P
           
            var temp_zz_gy = from a in Cache_data_xzys.jbz.AsEnumerable()
                          join b in p on a["zt"].ToString() equals b.csnr
                          where (a["tyyt"].ToString() == "别墅" || a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "小高层" || a["tyyt"].ToString() == "洋房" || a["tyyt"].ToString() == "洋楼")
                          group a by new { zc = a["zc"], zt = a["zt"] } into d
                          select new
                          {
                              zc = d.Key.zc,
                              gytl = d.Sum(m => m["jzmj"].doubls() + m["fzzmj"].doubls()),
                          };
            var temp_zz_cj = from a in Cache_data_cjjl.jbz.AsEnumerable()
                          join b in p on a["zt"].ToString() equals b.csnr
                          where (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                          group a by new { zc = a["zc"] } into d
                          select new
                          {
                              zc = d.Key.zc,
                              cjtl = d.Sum(m => m["jzmj"].doubls()),
                              cjje = d.Sum(m => m["cjje"].doubls())
                          };
            var data_zz_gy = (from a in zc
                           join b in temp_zz_gy
                           on a.zc equals b.zc
                           into temp
                           from tt in temp.DefaultIfEmpty()
                           select new
                           {
                               zc = a.zc,
                               zcmc = a.zcmc,
                               gytl = tt == null ? 0 : tt.gytl
                           }
                           ).GroupBy(m=> new{ m.zc,m.zcmc })
                           .Select(m=>new { zc = m.Key.zc,zcmc = m.Key.zcmc,gytl=m.Sum(a=>a.gytl)})
                           .OrderBy(m => m.zc).ToList();
             

            var data_zz_cj = (from a in zc
                           join b in temp_zz_cj
                           on a.zc equals b.zc
                           into temp
                           from tt in temp.DefaultIfEmpty()
                           select new
                           {
                               zc = a.zc,
                               zcmc = a.zcmc,
                               cjtl = tt == null ? 0 : tt.cjtl,
                               cjje = tt == null ? 0 : tt.cjje
                           }).OrderBy(m => m.zc).ToList();

            System.Data.DataTable dt_2 = new System.Data.DataTable();
            dt_2.Columns.Add("");
            for (int i = 1; i <= data_zz_gy.Count; i++)
            {
                dt_2.Columns.Add(data_zz_gy[i - 1].zcmc.ToString());
            }

            DataRow dr_2_1 = dt_2.NewRow();
            dr_2_1[0] = "供应体量";
            for (int i = 1; i <= data_zz_gy.Count; i++)
            {
                dr_2_1[i] = data_zz_gy[i - 1].gytl.mj_wf();
            }
            dt_2.Rows.Add(dr_2_1);

            DataRow dr_2_2 = dt_2.NewRow();
            DataRow dr_2_3 = dt_2.NewRow();

            dr_2_2[0] = "成交体量";
            dr_2_3[0] = "建面均价";
            for (int i = 1; i <= data_zz_cj.Count; i++)
            {
                dr_2_2[i] = data_zz_cj[i - 1].cjtl.mj_wf();
                dr_2_3[i] = (data_zz_cj[i - 1].cjje / data_zz_cj[i - 1].cjtl).je_y();
            }
            dt_2.Rows.Add(dr_2_2);
            dt_2.Rows.Add(dr_2_3);
            Office_Charts.Chart_gxzs(s[1], dt_2, 0);

            IAutoShape text2_1 = (IAutoShape)s[1].Shapes[3];
            text2_1.TextFrame.Text = string.Format(text2_1.TextFrame.Text, bkmc);

            IAutoShape text2_2 = (IAutoShape)s[1].Shapes[4];
            double cjtl_2 = data_zz_cj.FirstOrDefault(m => m.zc.ints() == Base_date.bz).cjtl;
            double cjjj_2 = data_zz_cj.FirstOrDefault(m => m.zc.ints() == Base_date.bz).cjje / cjtl_2;
            double jbzgytl_2 = data_zz_gy.Sum(m => m.gytl);
            double jbzcjtl_2 = data_zz_cj.Sum(m => m.cjtl);
            double gxb_2 = (jbzgytl_2 / jbzcjtl_2).ss_bfb_ys();
            text2_2.TextFrame.Text = string.Format(text2_2.TextFrame.Text, cjtl_2.mj_wf(), cjjj_2.je_y(), jbzgytl_2.mj_wf(), jbzcjtl_2.mj_wf(), gxb_2);

            IAutoShape text2_3 = (IAutoShape)s[1].Shapes[5];
            text2_3.TextFrame.Text = string.Format(text2_3.TextFrame.Text, bkmc);
            #endregion

            #region 3P
            ///015
            ///
          
            var temp_gc_cj = from a in Cache_data_cjjl.jbz.AsEnumerable()
                             join b in p on a["zt"].ToString() equals b.csnr
                             where (a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" )
                             group a by new { zc = a["zc"] } into d
                             select new
                             {
                                 zc = d.Key.zc,
                                 cjmj = d.Sum(m => m["jzmj"].doubls()),
                                 cjje = d.Sum(m => m["cjje"].doubls()),
                                 cjts = d.Count()
                             };
            var data_gc_cj = (from a in zc
                              join b in temp_gc_cj
                              on a.zc equals b.zc
                              into temp
                              from tt in temp.DefaultIfEmpty()
                              select new
                              {
                                  zc = a.zc,
                                  zcmc = a.zcmc,
                                  cjts = tt == null ? 0 : tt.cjts,
                                  cjjj = tt == null ? 0 : tt.cjje/tt.cjmj
                              }).OrderBy(m => m.zc).ToList();

            System.Data.DataTable dt_3 = new System.Data.DataTable();
            dt_3.Columns.Add("");
            for (int i = 1; i <= data_zz_gy.Count; i++)
            {
                dt_3.Columns.Add(data_gc_cj[i - 1].zcmc.ToString());
            }
            DataRow dr_3_2 = dt_3.NewRow();
            DataRow dr_3_3 = dt_3.NewRow();

            dr_3_2[0] = "成交套数";
            dr_3_3[0] = "建面均价";
            for (int i = 1; i <= data_zz_cj.Count; i++)
            {
                dr_3_2[i] = data_gc_cj[i - 1].cjts;
                dr_3_3[i] = data_gc_cj[i - 1].cjjj.je_y();
            }
            dt_3.Rows.Add(dr_3_2);
            dt_3.Rows.Add(dr_3_3);
            Office_Charts.Chart_cjqs(s[2], dt_3, 2);

            IAutoShape text3_1 = (IAutoShape)s[2].Shapes[0];
            text3_1.TextFrame.Text = string.Format(text3_1.TextFrame.Text, bkmc);
            var temp_data = temp_gc_cj.FirstOrDefault(m => m.zc.ints() == Base_date.bz);
            IAutoShape text3_2 = (IAutoShape)s[2].Shapes[1];
            text3_2.TextFrame.Text = string.Format(text3_2.TextFrame.Text, temp_data.cjts,(temp_data.cjje/temp_data.cjmj).ss_bfb_ys());

            IAutoShape text3_3 = (IAutoShape)s[2].Shapes[5];
            text3_3.TextFrame.Text = string.Format(text3_3.TextFrame.Text, bkmc);
            #endregion

            #region 4P
            ///234
            var temp_yf_cj = from a in Cache_data_cjjl.jbz.AsEnumerable()
                             join b in p on a["zt"].ToString() equals b.csnr
                             where (a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                             group a by new { zc = a["zc"] } into d
                             select new
                             {
                                 zc = d.Key.zc,
                                 cjmj = d.Sum(m => m["jzmj"].doubls()),
                                 cjje = d.Sum(m => m["cjje"].doubls()),
                                 cjts = d.Count()
                             };
            var data_yf_cj = (from a in zc
                              join b in temp_yf_cj
                              on a.zc equals b.zc
                              into temp
                              from tt in temp.DefaultIfEmpty()
                              select new
                              {
                                  zc = a.zc,
                                  zcmc = a.zcmc,
                                  cjts = tt == null ? 0 : tt.cjts,
                                  cjjj = tt == null ? 0 : tt.cjje / tt.cjmj
                              }).OrderBy(m => m.zc).ToList();

            System.Data.DataTable dt_4 = new System.Data.DataTable();
            dt_4.Columns.Add("");
            for (int i = 1; i <= data_yf_cj.Count; i++)
            {
                dt_4.Columns.Add(data_yf_cj[i - 1].zcmc.ToString());
            }
            DataRow dr_4_2 = dt_4.NewRow();
            DataRow dr_4_3 = dt_4.NewRow();

            dr_4_2[0] = "成交套数";
            dr_4_3[0] = "建面均价";
            for (int i = 1; i <= data_yf_cj.Count; i++)
            {
                dr_4_2[i] = data_yf_cj[i - 1].cjts;
                dr_4_3[i] = data_yf_cj[i - 1].cjjj.je_y();
            }
            dt_4.Rows.Add(dr_4_2);
            dt_4.Rows.Add(dr_4_3);
            Office_Charts.Chart_cjqs(s[3], dt_4, 0);

            IAutoShape text4_1 = (IAutoShape)s[3].Shapes[2];
            text4_1.TextFrame.Text = string.Format(text4_1.TextFrame.Text, bkmc);
            var temp_data_4 = data_yf_cj.FirstOrDefault(m => m.zc.ints() == Base_date.bz);
            IAutoShape text4_2 = (IAutoShape)s[3].Shapes[3];
            text4_2.TextFrame.Text = string.Format(text4_2.TextFrame.Text, temp_data_4.cjts, temp_data_4.cjjj.je_y());

            IAutoShape text4_3 = (IAutoShape)s[3].Shapes[4];
            text4_3.TextFrame.Text = string.Format(text4_3.TextFrame.Text, bkmc);
            #endregion

            #region 5P
            //234
            var temp_bs_cj = from a in Cache_data_cjjl.jbz.AsEnumerable()
                             join b in p on a["zt"].ToString() equals b.csnr
                             where (a["yt"].ToString() == "别墅" )
                             group a by new { zc = a["zc"] } into d
                             select new
                             {
                                 zc = d.Key.zc,
                                 cjmj = d.Sum(m => m["jzmj"].doubls()),
                                 cjje = d.Sum(m => m["cjje"].doubls()),
                                 cjts = d.Count()
                             };
            var data_bs_cj = (from a in zc
                              join b in temp_bs_cj
                              on a.zc equals b.zc
                              into temp
                              from tt in temp.DefaultIfEmpty()
                              select new
                              {
                                  zc = a.zc,
                                  zcmc = a.zcmc,
                                  cjts = tt == null ? 0 : tt.cjts,
                                  cjjj = tt == null ? 0 : tt.cjje / tt.cjmj
                              }).OrderBy(m => m.zc).ToList();

            System.Data.DataTable dt_5 = new System.Data.DataTable();
            dt_5.Columns.Add("");
            for (int i = 1; i <= data_bs_cj.Count; i++)
            {
                dt_5.Columns.Add(data_bs_cj[i - 1].zcmc.ToString());
            }
            DataRow dr_5_2 = dt_5.NewRow();
            DataRow dr_5_3 = dt_5.NewRow();

            dr_5_2[0] = "成交套数";
            dr_5_3[0] = "建面均价";
            for (int i = 1; i <= data_bs_cj.Count; i++)
            {
                dr_5_2[i] = data_bs_cj[i - 1].cjts;
                dr_5_3[i] = data_bs_cj[i - 1].cjjj.je_y();
            }
            dt_5.Rows.Add(dr_5_2);
            dt_5.Rows.Add(dr_5_3);
            Office_Charts.Chart_cjqs(s[4], dt_5, 0);

            IAutoShape text5_1 = (IAutoShape)s[4].Shapes[2];
            text5_1.TextFrame.Text = string.Format(text5_1.TextFrame.Text, bkmc);
            var temp_data_5 = data_bs_cj.FirstOrDefault(m => m.zc.ints() == Base_date.bz);
            IAutoShape text5_2 = (IAutoShape)s[4].Shapes[3];
            text5_2.TextFrame.Text = string.Format(text5_2.TextFrame.Text, temp_data_5.cjts, temp_data_5.cjjj.je_y());

            IAutoShape text5_3 = (IAutoShape)s[4].Shapes[4];
            text5_3.TextFrame.Text = string.Format(text5_3.TextFrame.Text, bkmc);
            #endregion

            return s;
        }
       
        public ISlideCollection plus9(string str, int cjbh)
        {
            var s =  new Presentation(str).Slides;

            //var zc = from a in Cache_data_xzys.jbz.AsEnumerable()
            //         group a by new { zcmc = a["zcmc"], zc = a["zc"] } into d
            //         select new
            //         {
            //             zc = d.Key.zc,
            //             zcmc = d.Key.zcmc
            //         };
            //var temp_gy = from a in Cache_data_xzys.jbz.AsEnumerable()
            //    where a["zt"].ToString() == "茶园"  && (a["tyyt"].ToString() == "别墅" || a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "小高层" || a["tyyt"].ToString() == "洋房" || a["tyyt"].ToString() == "洋楼")
            //    group a by new { zc = a["zc"], zt = a["zt"] } into d
            //    select new
            //    {
            //        zc = d.Key.zc,
            //        gytl = d.Sum(m =>m["jzmj"].doubls()+ double.Parse(m["fzzmj"].ToString())),
            //    };
            //var temp_cj = from a in Cache_data_cjjl.jbz.AsEnumerable()
            //               where a["zt"].ToString() == "茶园" && (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
            //              group a by new { zc = a["zc"] } into d
            //               select new
            //               {
            //                   zc = d.Key.zc,
            //                   cjtl = d.Sum(m =>m["jzmj"].doubls()),
            //                   cjje = d.Sum(m => m["cjje"].longs())
            //               };
            //var data_gy = (from a in zc
            //               join b in temp_gy
            //               on a.zc equals b.zc
            //               into temp
            //               from tt in temp.DefaultIfEmpty()
            //               select new
            //               {
            //                   zc = a.zc,
            //                   zcmc = a.zcmc,
            //                   gytl = tt == null ? 0 : tt.gytl
            //              }).OrderBy(m=>m.zc).ToList();


            //var data_cj = (from a in zc
            //               join b in temp_cj
            //               on a.zc equals b.zc
            //               into temp
            //               from tt in temp.DefaultIfEmpty()
            //               select new
            //               {
            //                   zc = a.zc,
            //                   zcmc = a.zcmc,
            //                   cjtl = tt == null ? 0 : tt.cjtl,
            //                   cjje = tt == null ? 0 : tt.cjje
            //               }).OrderBy(m => m.zc).ToList();

            //System.Data.DataTable dt = new System.Data.DataTable();
            //dt.Columns.Add("");
            //for (int i = 1; i <= data_gy.Count; i++)
            //{
            //    dt.Columns.Add(data_gy[i - 1].zcmc.ToString());
            //}

            //DataRow dr1 = dt.NewRow();
            //dr1[0]= "供应体量";
            //for (int i = 1; i <= data_gy.Count; i++)
            //{
            //    dr1[i] = data_gy[i-1].gytl.mj_wf();
            //}
            //dt.Rows.Add(dr1);

            //DataRow dr2 = dt.NewRow();
            //DataRow dr3 = dt.NewRow();

            //dr2[0] = "成交体量";
            //dr3[0] = "建面均价";
            //for (int i = 1; i <= data_cj.Count; i++)
            //{
            //    dr2[i] = data_cj[i-1].cjtl.mj_wf();
            //    dr3[i] = (data_cj[i-1].cjje / data_cj[i-1].cjtl).je_y();
            //}
            //dt.Rows.Add(dr2);
            //dt.Rows.Add(dr3);
            //Office_Charts.Chart_gxzs(s[0], dt, 2);
            
            return s;
        }
        public ISlideCollection plus10(string str, int cjbh)
        {
            return new Presentation(str).Slides;
        }
        public ISlideCollection plus11(string str, int cjbh)
        {
            return new Presentation(str).Slides;
        }
        public ISlideCollection plus12(string str, int cjbh)
        {
            return new Presentation(str).Slides;
        }
        public ISlideCollection plus13(string str, int cjbh)
        {
            return new Presentation(str).Slides;
        }
        public ISlideCollection plus14(string str, int cjbh)
        {
            return new Presentation(str).Slides;
        }
        public ISlideCollection plus15(string str, int cjbh)
        {
            return new Presentation(str).Slides;
        }
        public ISlideCollection plus16(string str, int cjbh)
        {
            return new Presentation(str).Slides;
        }
        #endregion

        #region 竞品插件
        #region 复地
        /// <summary>
        /// 复地竞品插件
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_fudi(string str, int cjbh)
        {
            Cache_data_cjjl.bz.Select("");
            return new Presentation(str).Slides;
        }
        #endregion
        #endregion

    }
}
