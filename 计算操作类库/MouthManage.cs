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
    class MouthManage
    {
        public MouthManage()
        {
            // Aspose_Crack.SlideCrack();
            //MessageBox.Show("之前"+Cache_Result_yb.td_by_zyd.ToString());
            Cache_data_cjjl.ini_yb(2018, 4);
            Cache_data_xzys.ini_yb();
            Cache_data_tdjyjl.ini_yb();
            Cache_Result_yb.ini();
        }
        public static int T = 0;
        public void t()
        {
            Presentation ppt = SlideFactory.GetInstance().ppt;
            slide7(ppt.Slides[6]);
            slide14(ppt.Slides[13]);
            slide16(ppt.Slides[15]);
            slide17(ppt.Slides[16]);
            slide18(ppt.Slides[17]);
            slide19(ppt.Slides[18]);
            slide20(ppt.Slides[19]);


            ppt.Save("D:\\123123.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            T = 1;

        }

        public void slide7(ISlide slide)
        {
            string str1 = thread1();
            string str2 = thread2();
            IAutoShape itf = (IAutoShape)slide.Shapes[2];
            itf.TextFrame.Paragraphs[0].Text = str1;
            itf.TextFrame.Paragraphs[1].Text = str2;
            ITextFrame tf = itf.TextFrame;
            foreach (var item in tf.Paragraphs)
            {
                IPortion port = item.Portions[0];
                port.PortionFormat.LatinFont = new FontData("微软雅黑");
                port.PortionFormat.FontBold = NullableBool.NotDefined;
                port.PortionFormat.FontHeight = 12;
            }

        }
        public void slide14(ISlide slide)
        {

        }
        public void slide16(ISlide slide)
        {
            string str1 = thread4();
            IAutoShape itf = (IAutoShape)slide.Shapes[6];
            itf.TextFrame.Text = str1;
            ITextFrame tf = itf.TextFrame;
            foreach (var item in tf.Paragraphs)
            {
                IPortion port = item.Portions[0];
                port.PortionFormat.LatinFont = new FontData("微软雅黑");
                port.PortionFormat.FontBold = NullableBool.NotDefined;
                port.PortionFormat.FontHeight = 12;
            }

        }
        public void slide17(ISlide slide)
        {
            Office_Charts.DoubleAxexchart(slide, Cache_Result_yb.jsjg_scgxfx, 3, 0, 1);
            Office_Charts.DoubleAxexchart(slide, Cache_Result_yb.jsjg_scgxfx_psb, 2, Aspose.Slides.Charts.ChartType.StackedBar);
        }
        public void slide18(ISlide slide)
        {
            Office_ChartStyle style = new Office_ChartStyle();
            style.坐标方向 = Base_Config.坐标方向.横向;
            style.文字位置 = Aspose.Slides.Charts.LegendDataLabelPosition.Center;
            style.文字旋转方向 = TextVerticalType.Vertical270;
            style.是否显示文字 = true;
            Office_Charts.SingleAxexchart(slide, Cache_Result_yb.jsjg_scjgfx, 2, style);

            IAutoShape itf = (IAutoShape)slide.Shapes[5];
            itf.TextFrame.Paragraphs[0].Text = thread5();
            itf.TextFrame.Paragraphs[1].Text = thread6();
            ITextFrame tf = itf.TextFrame;
            foreach (var item in tf.Paragraphs)
            {
                IPortion port = item.Portions[0];
                port.PortionFormat.LatinFont = new FontData("微软雅黑");
                port.PortionFormat.FontBold = NullableBool.NotDefined;
                port.PortionFormat.FontHeight = 12;
            }
        }
        public void slide19(ISlide slide)
        {
            // var a = from a in Cache_data_cjjl.@by.AsEnumerable()
            //        group a by 
            //Charts.SingleAxexchart(slide, Cache_data_cjjl.by., 2);

            var query = from t in Cache_data_cjjl.@by.AsEnumerable()
                        group t by new { t1 = t.Field<string>("kfsmc") } into m
                        select new
                        {
                            kfsmc = m.Key.t1,
                            cjje = m.Sum(n => n.Field<long>("cjje"))
                        };
            DataTable dt = new DataTable();
            dt.Columns.Add("kfsmc");
            dt.Columns.Add("cjje");
            foreach (var item in query.OrderByDescending(m => m.cjje).Take(10).ToList())
            {
                DataRow dr = dt.NewRow();
                dr[0] = item.kfsmc;
                dr[1] = item.cjje.je_yy();
                dt.Rows.Add(dr);
            }
            Office_ChartStyle style = new Office_ChartStyle();
            style.坐标方向 = Base_Config.坐标方向.纵向;
            style.文字位置 = Aspose.Slides.Charts.LegendDataLabelPosition.OutsideEnd;
            style.文字旋转方向 = TextVerticalType.Horizontal;
            style.是否显示文字 = true;
            Office_Charts.SingleAxexchart(slide, dt, 3, style);

        }
        public void slide20(ISlide slide)
        {
            // var a = from a in Cache_data_cjjl.@by.AsEnumerable()
            //        group a by 
            //Charts.SingleAxexchart(slide, Cache_data_cjjl.by., 2);

            var query = from t in Cache_data_cjjl.@by.AsEnumerable()
                        group t by new { t1 = t.Field<string>("lpmc") } into m
                        select new
                        {
                            lpmc = m.Key.t1,
                            cjje = m.Sum(n => n.Field<long>("cjje"))
                        };
            DataTable dt = new DataTable();
            dt.Columns.Add("lpmc");
            dt.Columns.Add("cjje");
            foreach (var item in query.OrderByDescending(m => m.cjje).Take(10).ToList())
            {
                DataRow dr = dt.NewRow();
                dr[0] = item.lpmc;
                dr[1] = item.cjje.je_yy();
                dt.Rows.Add(dr);
            }
            Office_ChartStyle style = new Office_ChartStyle();
            style.坐标方向 = Base_Config.坐标方向.纵向;
            style.文字位置 = Aspose.Slides.Charts.LegendDataLabelPosition.OutsideEnd;
            style.文字旋转方向 = TextVerticalType.Horizontal;
            style.是否显示文字 = true;
            Office_Charts.SingleAxexchart(slide, dt, 1, style);

        }
        public void slide22(ISlide slide)
        {


        }
        private string thread1()
        {

            //**************成交记录*********************//
            string str = string.Format(@"重庆{0}月市场新增供应（{1}）万方，环比{2}，同比({3})万方，成交面积({4})万方,环比{5}，同比{6}；成交金额（{7}）亿元，环比{8}，同比{9}；成交建面单价（{10}）元 /㎡ 套内（{11}）元 /㎡），环比{12}，同比{13}；其中纯住宅（不含经适房）套内单价（{14}）元/㎡，环比{15}。",
                Base_date.by_First.Month,
                (Cache_Result_yb.by_cj_jzmj_xzys + Cache_Result_yb.by_cj_jzmj_fzz_xzys).mj_wf(),
                (((Cache_Result_yb.by_cj_jzmj_xzys + Cache_Result_yb.by_cj_jzmj_fzz_xzys) - (Cache_Result_yb.sy_cj_jzmj_xzys + Cache_Result_yb.sy_cj_jzmj_fzz_xzys)) * 100 / (Cache_Result_yb.sy_cj_jzmj_xzys + Cache_Result_yb.sy_cj_jzmj_fzz_xzys)).ss_bfb(), //环比
                "0",//同比
                Cache_Result_yb.by_cj_jzmj.mj_wf(),//建筑面积
                ((Cache_Result_yb.by_cj_jzmj - Cache_Result_yb.sy_cj_jzmj) * 100 / Cache_Result_yb.sy_cj_jzmj).ss_bfb(), //环比
                "0",//同比
                 Cache_Result_yb.by_cj_cjje.je_yy(),//成交金额
                ((Cache_Result_yb.by_cj_cjje - Cache_Result_yb.sy_cj_cjje) * 100 / Cache_Result_yb.sy_cj_cjje).doubls().ss_bfb(),//环比
                "0",//同比
                (Cache_Result_yb.by_cj_cjje / Cache_Result_yb.by_cj_jzmj).je_y(), //建面单价
                (Cache_Result_yb.by_cj_cjje / Cache_Result_yb.by_cj_tnmj).je_y(), //套内单价
                (((Cache_Result_yb.by_cj_cjje / Cache_Result_yb.by_cj_jzmj) - (Cache_Result_yb.sy_cj_cjje / Cache_Result_yb.sy_cj_jzmj)) / (Cache_Result_yb.sy_cj_cjje / Cache_Result_yb.sy_cj_jzmj)).ss_bfb(),
                "0",//同比
                (Cache_Result_yb.by_cj_czz.Sum(m => double.Parse(m["cjje"].ToString())) / Cache_Result_yb.by_cj_czz.Sum(m => double.Parse(m["tnmj"].ToString()))).je_y(),
                ((Cache_Result_yb.by_cj_czz.Sum(m => double.Parse(m["cjje"].ToString())) / Cache_Result_yb.by_cj_czz.Sum(m => double.Parse(m["tnmj"].ToString())) - (Cache_Result_yb.sy_cj_czz.Sum(m => double.Parse(m["cjje"].ToString())) / Cache_Result_yb.sy_cj_czz.Sum(m => double.Parse(m["tnmj"].ToString())))) * 100 / (Cache_Result_yb.sy_cj_czz.Sum(m => double.Parse(m["cjje"].ToString())) / Cache_Result_yb.sy_cj_czz.Sum(m => double.Parse(m["tnmj"].ToString())))).ss_bfb()
                );
            return str;
        }

        private string thread2()
        {
            //----------------------本月-------------------------//
            //**************成交记录*********************//
            //成交性质
            var by_cjxz = Cache_data_tdjyjl.by.AsEnumerable().GroupBy(m => m[2]);
            //成交方式
            var by_cjfs = Cache_data_tdjyjl.by.AsEnumerable().GroupBy(m => m[6]);
            //----------------------上月-------------------------//
            //成交性质
            var sy_cjxz = Cache_data_tdjyjl.sy.AsEnumerable().GroupBy(m => m[2]);
            //成交方式
            var sy_cjfs = Cache_data_tdjyjl.sy.AsEnumerable().GroupBy(m => m[6]);


            string str = string.Format(@"{0}月商住类地块成交{1}宗（纯住{2}宗，住兼商{3}宗，纯商业{4}宗，商兼住{5}宗），较上月{6}宗， 成交土地面积{7}亩，环比{8}%，同比{9}%；可供开发体量{10}万方，环比{11}%，同比{12}%；土地综合出让金{13}亿元，环比{14}%，同比{15}%，{16}宗拍卖成交，{17}宗挂牌成交，整体溢价率{18}% ",
                 Base_date.by_First.Month,
                 by_cjxz.Sum(m => m.Count()),
                 by_cjxz.FirstOrDefault(m => m.Key.ToString() == "居住") == null ? "0" : by_cjxz.FirstOrDefault(m => m.Key.ToString() == "居住").Count().ToString(),
                 by_cjxz.FirstOrDefault(m => m.Key.ToString() == "居住兼容商业") == null ? "0" : by_cjxz.FirstOrDefault(m => m.Key.ToString() == "居住兼容商业").Count().ToString(),
                 by_cjxz.FirstOrDefault(m => m.Key.ToString() == "商业") == null ? "0" : by_cjxz.FirstOrDefault(m => m.Key.ToString() == "商业").Count().ToString(),
                 by_cjxz.FirstOrDefault(m => m.Key.ToString() == "商业兼容居住") == null ? "0" : by_cjxz.FirstOrDefault(m => m.Key.ToString() == "商业兼容居住").Count().ToString(),
                 by_cjxz.Sum(m => m.Count()) - sy_cjxz.Sum(m => m.Count()),
                 Cache_Result_yb.td_by_zyd.mj_m(),
                 ((Cache_Result_yb.td_by_zyd - Cache_Result_yb.td_sy_zyd) * 100 / Cache_Result_yb.td_sy_zyd).ss_bfb(), //环比
                 "0",//同比
                 Cache_Result_yb.td_by_kjtl.mj(),//可建体量
                 ((Cache_Result_yb.td_by_kjtl - Cache_Result_yb.td_sy_kjtl) * 100 / Cache_Result_yb.td_sy_kjtl).ss_bfb(), //环比
                 "0",//同比
                 Cache_Result_yb.td_by_cjje.je_w_to_yy(),//成交金额
                 ((Cache_Result_yb.td_by_cjje - Cache_Result_yb.td_sy_cjje) * 100 / Cache_Result_yb.td_sy_cjje).ss_bfb(),//环比
                 "0",//同比
                 (from a in by_cjfs where a.Key.ToString() == "拍卖" select a.Count()).ElementAt(0).ToString(),
                 (from a in by_cjfs where a.Key.ToString() == "挂牌" select a.Count()).ElementAt(0).ToString(),
                 "0"
                 );
            return str;
        }

        private string thread3()
        {
            int by_jyzs = Cache_data_cjjl.by.Rows.Count;
            var by_ytfz = Cache_data_cjjl.by.AsEnumerable().GroupBy(m => m[11]);

            var sy_ytfz = Cache_data_cjjl.by.AsEnumerable().GroupBy(m => m[11]);
            string.Format(@"{0}月各业态成交量全线下降，其中高层（占{1}%，{2}%），别墅（占{3} %，{4}%），另外商业（占{5}%）、商务（{6}%）、车库（占{7}%）占比小幅下降，经适房（占{8}%）、洋房（占{9}%）占比持平。"
                    , Base_date.by_First.Month
                    , by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层") == null ? "0" : ((double)(by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count() * 100) / by_jyzs).ss_bfb()
                    , by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层") == null ? "0" :
                                    ((double)(by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count() - (sy_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count())) / by_jyzs).ss_bfb()
                    , by_ytfz.FirstOrDefault(m => m.Key.ToString() == "别墅") == null ? "0" : ((double)(by_ytfz.FirstOrDefault(m => m.Key.ToString() == "别墅").Count() * 100) / by_jyzs).ss_bfb()
                    , by_ytfz.FirstOrDefault(m => m.Key.ToString() == "别墅") == null ? "0" :
                                    ((double)(by_ytfz.FirstOrDefault(m => m.Key.ToString() == "别墅").Count() - (sy_ytfz.FirstOrDefault(m => m.Key.ToString() == "别墅").Count())) / by_jyzs).ss_bfb()
                    , by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层") == null ? "0" : ((double)(by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count() * 100) / by_jyzs).ss_bfb()
                    , by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层") == null ? "0" :
                                    ((double)(by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count() - (sy_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count())) / by_jyzs).ss_bfb()
                    , by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层") == null ? "0" : ((double)(by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count() * 100) / by_jyzs).ss_bfb()
                    , by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层") == null ? "0" :
                                    ((double)(by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count() - (sy_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count())) / by_jyzs).ss_bfb()
                    , by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层") == null ? "0" : ((double)(by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count() * 100) / by_jyzs).ss_bfb()
                    , by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层") == null ? "0" :
                                    ((double)(by_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count() - (sy_ytfz.FirstOrDefault(m => m.Key.ToString() == "高层").Count())) / by_jyzs).ss_bfb()
                                    );

            return "";
        }

        private string thread4()
        {

            //----------------------本月-------------------------//
            //**************新增预售*********************//

            //**************成交记录*********************//
            //纯住房
            var by_czf = Cache_data_cjjl.by.AsEnumerable().Where(m => m[11].ToString() == "别墅" || m[11].ToString() == "高层" || m[11].ToString() == "小高层" || m[11].ToString() == "洋房" || m[11].ToString() == "洋楼");

            //----------------------上月-------------------------//

            //**************成交记录*********************//
            //纯住房
            var sy_czf = Cache_data_cjjl.sy.AsEnumerable().Where(m => m[11].ToString() == "别墅" || m[11].ToString() == "高层" || m[11].ToString() == "小高层" || m[11].ToString() == "洋房" || m[11].ToString() == "洋楼");

            string str = string.Format(@"重庆{0}月市场整体成交面积{1}万方，环比{2}，同比{3}；成交金额{4}亿元，环比{5}，同比{6}；成交均价{7}元/㎡（套内{8}元/㎡），环比{9}，同比{10}；其中纯住宅（不含经适房）套内单价{11}元/㎡，环比{12}。本月整体市场成交环比“量跌价涨” ，同比“量跌价涨”",
                                        Base_date.by_First.Month,
                Cache_Result_yb.by_cj_jzmj.mj_wf(),//建筑面积
                ((Cache_Result_yb.by_cj_jzmj - Cache_Result_yb.sy_cj_jzmj) * 100 / Cache_Result_yb.sy_cj_jzmj).ss_bfb(), //环比
                "0",//同比
                Cache_Result_yb.by_cj_cjje.je_yy(),//成交金额
                ((Cache_Result_yb.by_cj_cjje - Cache_Result_yb.sy_cj_cjje) * 100 / Cache_Result_yb.sy_cj_cjje).doubls().ss_bfb(),//环比
                "0",//同比
                (Cache_Result_yb.by_cj_cjje / Cache_Result_yb.by_cj_jzmj).je_y(), //建面单价
                (Cache_Result_yb.by_cj_cjje / Cache_Result_yb.by_cj_tnmj).je_y(), //套内单价
                (((Cache_Result_yb.by_cj_cjje / Cache_Result_yb.by_cj_jzmj) - (Cache_Result_yb.by_cj_cjje / Cache_Result_yb.by_cj_jzmj)) / (Cache_Result_yb.by_cj_cjje / Cache_Result_yb.by_cj_jzmj)).ss_bfb(),
                "0",//同比
                (by_czf.Sum(m => double.Parse(m[18].ToString())) / by_czf.Sum(m => double.Parse(m[17].ToString()))).mj(),
                ((by_czf.Sum(m => double.Parse(m[18].ToString())) / by_czf.Sum(m => double.Parse(m[17].ToString())) - sy_czf.Sum(m => double.Parse(m[18].ToString())) / sy_czf.Sum(m => double.Parse(m[17].ToString()))) / sy_czf.Sum(m => double.Parse(m[18].ToString())) / sy_czf.Sum(m => double.Parse(m[17].ToString()))).ss_bfb()
                                        );
            return str;
        }

        private string thread5()
        {
            var by_jzmj = Cache_data_cjjl.by.AsEnumerable().Sum(m => double.Parse(m[16].ToString()));
            var sy_jzmj = Cache_data_cjjl.sy.AsEnumerable().Sum(m => double.Parse(m[16].ToString()));
            //var ty_jzmj = Cache_data_cjjl.ty.AsEnumerable().Sum(m => double.Parse(m[16].ToString()));
            string str = string.Format(@"{0}月成交量为{1}万方，环比{2}%，同比{3}% ，本月呈“量跌价涨”",
                Base_date.by_First.ToString("yyyy年MM"),
                 Cache_Result_yb.by_cj_jzmj.mj_wf(),
                 ((Cache_Result_yb.by_cj_jzmj - Cache_Result_yb.sy_cj_jzmj) * 100 / Cache_Result_yb.sy_cj_jzmj).ss_bfb(),
                 0
                );
            return str;

        }
        private string thread6()
        {

            //var ty_jzmj = Cache_data_cjjl.ty.AsEnumerable().Sum(m => double.Parse(m[16].ToString()));
            string str = string.Format(@"{0}重庆主城均价为{1}元/㎡，环比上涨{2}%，同比上涨{3}% 。",
                Base_date.by_First.ToString("yyyy年MM"),
                 (Cache_Result_yb.by_cj_cjje / Cache_Result_yb.by_cj_jzmj).je_y(),
                 (((Cache_Result_yb.by_cj_cjje / Cache_Result_yb.by_cj_jzmj) - (Cache_Result_yb.sy_cj_cjje / Cache_Result_yb.sy_cj_jzmj)) * 100 / (Cache_Result_yb.sy_cj_cjje / Cache_Result_yb.sy_cj_jzmj)).ss_bfb(),
                 0
                );
            return str;

        }

        private string thread7()
        {
            return "";
        }
    }
}
