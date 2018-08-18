﻿using Aspose.Slides;
using Aspose.Slides.Charts;
using Calculation.Base;
using Calculation.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    public class plus_jp_base :weak
    {
        /// <summary>
        /// 窄度界限
        /// </summary>
        public static double zd = 0.85;
        /// <summary>
        /// 宽度界限
        /// </summary>
        public static double kd = 1.2;
        /// <summary>
        /// 竞争格局
        /// 别墅按细分业态参数分页
        /// 商务按户型查参数业态
        /// 其他按业态查参数
        /// 数据源是认购数据
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_dyt_jzgj(int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);

                #region P1 
                foreach (var item in param)
                {

                    #region 商务

                    
                    if (item.ytcs[0] == "商务")
                    {
                        var page = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_JZGJ"]).Slides[0];
                        IAutoShape text1 = (IAutoShape)page.Shapes[2];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.hxcs[0]);
                        //数据
                        System.Data.DataTable jzgjt = new System.Data.DataTable();
                        jzgjt.Columns.Add("");
                        jzgjt.Columns.Add("成交套数", typeof(int));
                        jzgjt.Columns.Add("建面均价", typeof(double));
                        //图表
                        IChart chart = (IChart)page.Shapes[3];
                        foreach (var item_jp in item.jpxmlb)
                        {


                            string jpyt = item_jp.hxcs == null ? item.hxcs[0] : item_jp.hxcs[0];
                            var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == jpyt).FirstOrDefault();

                            DataRow dr1 = jzgjt.NewRow();
                            dr1[0] = item_jp.lpcs[0] + "(" + item.ytcs[0] + ")";
                            if (jpcjxx != null)
                            {

                                dr1[1] = jpcjxx["xkts"].ints();
                                dr1[2] = jpcjxx["xkjmjj"].ints();
                            }
                            else
                            {
                                dr1[1] = 0;
                                dr1[2] = 0;
                            }
                            jzgjt.Rows.Add(dr1);

                        }
                        Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                        t.AddClone(page);

                    }
                    #endregion

                    #region 别墅
                   

                    else if (item.ytcs[0] == "别墅")
                    {
                        if(item.xfytcs!=null&&item.xfytcs.Length>0)
                        {
                            for (int i = 0; i < item.xfytcs.Length; i++)
                            {
                                var page = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_JZGJ"]).Slides[0];
                                IAutoShape text1 = (IAutoShape)page.Shapes[2];
                                text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.xfytcs[i]);
                                System.Data.DataTable jzgjt = new System.Data.DataTable();
                                jzgjt.Columns.Add("");
                                jzgjt.Columns.Add("成交套数", typeof(int));
                                jzgjt.Columns.Add("建面均价", typeof(double));

                                IChart chart = (IChart)page.Shapes[3];
                                foreach (var item_jp in item.jpxmlb)
                                {

                                    
                                    var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item.xfytcs[i]).FirstOrDefault();

                                    DataRow dr1 = jzgjt.NewRow();
                                    dr1[0] = item_jp.lpcs[0] + "(" + item.ytcs[0] + ")";
                                    if (jpcjxx != null)
                                    {

                                        dr1[1] = jpcjxx["xkts"].ints();
                                        dr1[2] = jpcjxx["xkjmjj"].ints();
                                        jzgjt.Rows.Add(dr1);
                                    }
                                    else
                                    {
                                        if (item_jp.xfytcs.Contains(item.xfytcs[i]))
                                        {
                                            dr1[1] = 0;
                                            dr1[2] = 0;
                                            jzgjt.Rows.Add(dr1);
                                        }
                                        else
                                            continue;
                                    }
                                }
                                Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                                t.AddClone(page);
                            }
                           
                        }
                        else
                        {
                            continue;
                        }
                    }

                    #endregion

                    #region 大业态

                  
                    else {
                        var page = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_JZGJ"]).Slides[0];
                        IAutoShape text1 = (IAutoShape)page.Shapes[2];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                        //数据
                        System.Data.DataTable jzgjt = new System.Data.DataTable();
                        jzgjt.Columns.Add("");
                        jzgjt.Columns.Add("成交套数", typeof(int));
                        jzgjt.Columns.Add("建面均价", typeof(double));
                        foreach (var item_jp in item.jpxmlb)
                        {
                            string jpyt = item_jp.ytcs == null ? item.ytcs[0] : item_jp.ytcs[0];
                            var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == jpyt).FirstOrDefault();

                            DataRow dr1 = jzgjt.NewRow();
                            dr1[0] = item_jp.lpcs[0] + "(" + item.ytcs[0] + ")";
                            if (jpcjxx != null)
                            {

                                dr1[1] = jpcjxx["xkts"].ints();
                                dr1[2] = jpcjxx["xkjmjj"].ints();
                            }
                            else
                            {
                                dr1[1] = 0;
                                dr1[2] = 0;
                            }
                            jzgjt.Rows.Add(dr1);

                          
                        }
                        Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                        t.AddClone(page);
                    }

                    #endregion

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
        /// 细分也太竞争格局
        /// </summary>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_xfyt_jzgj(int cjbh)
        {
            var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
            var p = new Presentation();
            var t = p.Slides;
            t.RemoveAt(0);

            #region P1 


            foreach (var item in param)
            {
                if (item.ytcs[0] == "别墅" || item.ytcs[0] == "商务")
                {
                    if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {
                            var tp = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_JZGJ"]);
                            var temp = tp.Slides;
                            var page = temp[0];
                            IAutoShape text1 = (IAutoShape)page.Shapes[2];
                            text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.xfytcs[i]);
                            //数据
                            System.Data.DataTable jzgjt = new System.Data.DataTable();
                            jzgjt.Columns.Add("");
                            jzgjt.Columns.Add("成交套数", typeof(int));
                            jzgjt.Columns.Add("建面均价", typeof(double));
                            //图表
                            IChart chart = (IChart)page.Shapes[3];
                            #region 本案
                            var bacjxx = Cache_data_cjjl.bz.AsEnumerable().Where(a => a["lpmc"].ToString() == item.lpcs[0] && a["xfyt"].ToString() == item.xfytcs[i]);

                            DataRow dr = jzgjt.NewRow();
                            dr[0] = item.lpcs[0] + item.xfytcs[i];
                            dr[1] = bacjxx.Sum(m => m["ts"].ints());
                            dr[2] = bacjxx.Sum(m => m["cjje"].ints()) / bacjxx.Sum(m => m["jzmj"].doubls());
                            jzgjt.Rows.Add(dr);
                            #endregion
                            #region 竞争项目
                            foreach (var item_jp in item.jpxmlb)
                            {
                                string jpyt = item.xfytcs[i];
                                var jpcjxx = Cache_data_cjjl.bz.AsEnumerable().Where(a => a["lpmc"].ToString() == item_jp.lpcs[0] && a["xfyt"].ToString() == jpyt);

                                DataRow dr1 = jzgjt.NewRow();
                                dr1[0] = item_jp.lpcs[0] + "(" + item.xfytcs[i] + ")";
                                if (jpcjxx != null)
                                {

                                    dr1[1] = jpcjxx.Sum(m => m["ts"].ints());
                                    dr1[2] = jpcjxx.Sum(m => m["cjje"].ints()) / jpcjxx.Sum(m => m["jzmj"].doubls());
                                }
                                else
                                {
                                    dr1[1] = 0;
                                    dr1[2] = 0;
                                }
                                jzgjt.Rows.Add(dr1);

                            }
                            #endregion
                            Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                            t.AddClone(page);
                        }
                    }

                }
            }
            return t;
            #endregion
        }
        /// <summary>
        /// 大业态推广图片
        /// </summary>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_dyt_tgtp(int cjbh)
        {
            string path = ConfigurationManager.AppSettings["DgPath"] + Base_date.bn + "\\" + Base_date.bz;
            
            var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
            var t = new Presentation().Slides;
            t.RemoveAt(0);
            foreach (var item in param)
            {
                List<Zb_Jp_Tgtp_Model> tgtplb = new List<Zb_Jp_Tgtp_Model>();
                try
                {
                    Image img = (Image)new Bitmap(Path.Combine(path, item.lpcs[0] + ".jpg"));
                    if ((img.Width / 1.0) / img.Height < zd)
                    {
                        Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                        tgtp.img = img;
                        tgtp.xmmc = item.lpcs[0];
                        tgtp.tplx = Models.Enums.TP_LX.窄图;
                        tgtplb.Add(tgtp);
                    }
                    else if ((img.Width / 1.0) / img.Height > zd && (img.Width / 1.0) / img.Height < kd)
                    {
                        Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                        tgtp.img = img;
                        tgtp.xmmc = item.lpcs[0];
                        tgtp.tplx = Models.Enums.TP_LX.方图;
                        tgtplb.Add(tgtp);
                    }
                    else
                    {
                        Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                        tgtp.img = img;
                        tgtp.xmmc = item.lpcs[0];
                        tgtp.tplx = Models.Enums.TP_LX.宽图;
                        tgtplb.Add(tgtp);
                    }
                }
                catch
                {
                    Base_Log.Log(Path.Combine(path, item.bamc + ".jpg") + "文件不存在");
                }
                foreach (var item_jp in item.jpxmlb)
                {
                    try
                    {
                        Image img = (Image)new Bitmap(Path.Combine(path, item_jp.lpcs[0] + ".jpg"));
                        if ((img.Width / 1.0) / img.Height < zd)
                        {
                            Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                            tgtp.img = img;
                            tgtp.xmmc = item_jp.lpcs[0];
                            tgtp.tplx = Models.Enums.TP_LX.窄图;
                            tgtplb.Add(tgtp);
                        }
                        else if ((img.Width / 1.0) / img.Height > zd && (img.Width / 1.0) / img.Height < kd)
                        {
                            Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                            tgtp.img = img;
                            tgtp.xmmc = item_jp.lpcs[0];
                            tgtp.tplx = Models.Enums.TP_LX.方图;
                            tgtplb.Add(tgtp);
                        }
                        else
                        {
                            Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                            tgtp.img = img;
                            tgtp.xmmc = item_jp.lpcs[0];
                            tgtp.tplx = Models.Enums.TP_LX.宽图;
                            tgtplb.Add(tgtp);
                        }
                    }
                    catch
                    {
                        Base_Log.Log(Path.Combine(path, item_jp.lpcs[0] + ".jpg") + "文件不存在");
                    }
                }
                if (tgtplb.Count > 0)
                {
                    List<Zb_Jp_Tgtp_Model> zt_pic = new List<Zb_Jp_Tgtp_Model>();
                    List<Zb_Jp_Tgtp_Model> ft_pic = new List<Zb_Jp_Tgtp_Model>();
                    var zt = tgtplb.Where(m => m.tplx == Models.Enums.TP_LX.窄图);
                    var ft = tgtplb.Where(m => m.tplx == Models.Enums.TP_LX.方图);
                    var kt = tgtplb.Where(m => m.tplx == Models.Enums.TP_LX.宽图);
                    if (zt != null && zt.Count() > 0)
                    {
                        var ztlist = zt.ToList();
                        for (int i = 0; i < ztlist.Count; i++)
                        {
                            zt_pic.Add(ztlist[i]);
                            if ((i + 1) % 2 == 0 || i + 1 >= ztlist.Count)
                            {
                                var tp1 = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_TGTP"]);
                                var temp1 = tp1.Slides;
                                for (int j = 0; j < zt_pic.Count; j++)
                                {
                                    IAutoShape text = temp1[2].Shapes.AddAutoShape(ShapeType.Rectangle, 20 + (220 * j), 130, 210, 40);
                                    text.TextFrame.Text = zt_pic[j].xmmc;
                                    text.ShapeStyle.FontColor.Color = Color.Black;
                                    text.FillFormat.FillType = FillType.NoFill;
                                    text.ShapeStyle.LineColor.Color = Color.White;
                                    IPPImage img1 = tp1.Images.AddImage(zt_pic[j].img);
                                    int height = (img1.Height * 210 / img1.Width);
                                    temp1[2].Shapes.AddPictureFrame(ShapeType.Rectangle, 20 + (220 * j), 170, 210, height, img1);
                                }
                                t.AddClone(temp1[2]);
                                zt_pic.Clear();
                            }
                        }
                    }
                    if (ft != null && ft.Count() > 0)
                    {
                        var ftlist = ft.ToList();
                        for (int i = 0; i < ftlist.Count; i++)
                        {

                            ft_pic.Add(ftlist[i]);
                            if ((i + 1) % 2 == 0)
                            {
                                var tp1 = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_TGTP"]);
                                var temp1 = tp1.Slides;
                                for (int j = 0; j < ft_pic.Count; j++)
                                {
                                    IAutoShape text = temp1[2].Shapes.AddAutoShape(ShapeType.Rectangle, 20 + (280 * j), 130, 210, 40);
                                    text.TextFrame.Text = ft_pic[j].xmmc;
                                    text.ShapeStyle.FontColor.Color = Color.Black;
                                    text.FillFormat.FillType = FillType.NoFill;
                                    text.ShapeStyle.LineColor.Color = Color.White;
                                    IPPImage img1 = tp1.Images.AddImage(ft_pic[j].img);
                                    int height = (img1.Height * 270 / img1.Width);
                                    temp1[2].Shapes.AddPictureFrame(ShapeType.Rectangle, 20 + (280 * j), 170, 270, height, img1);
                                }
                                t.AddClone(temp1[2]);
                                ft_pic.Clear();
                            }
                            else if (i + 1 >= ftlist.Count)
                            {
                                var tp1 = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_TGTP"]);
                                var temp1 = tp1.Slides;
                                for (int j = 0; j < ft_pic.Count; j++)
                                {
                                    IAutoShape text = temp1[2].Shapes.AddAutoShape(ShapeType.Rectangle, 20 + (670 - 280) / 2, 130, 210, 40);
                                    text.TextFrame.Text = ft_pic[j].xmmc;
                                    text.ShapeStyle.FontColor.Color = Color.Black;
                                    text.FillFormat.FillType = FillType.NoFill;
                                    text.ShapeStyle.LineColor.Color = Color.White;
                                    IPPImage img1 = tp1.Images.AddImage(ft_pic[j].img);
                                    int height = (img1.Height * 270 / img1.Width);
                                    temp1[2].Shapes.AddPictureFrame(ShapeType.Rectangle, 20 + (670 - 280) / 2, 170, 270, height, img1);
                                }
                                t.AddClone(temp1[2]);
                                ft_pic.Clear();
                            }
                        }
                    }
                    if (kt != null && kt.Count() > 0)
                    {
                        var ktlist = kt.ToList();
                        for (int i = 0; i < ktlist.Count; i++)
                        {
                            var tp1 = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_TGTP"]);
                            var temp1 = tp1.Slides;
                            IAutoShape text = temp1[2].Shapes.AddAutoShape(ShapeType.Rectangle, 20 + (670 - 440) / 2, 130, 440, 40);
                            text.TextFrame.Text = ktlist[i].xmmc;
                            text.ShapeStyle.FontColor.Color = Color.Black;
                            text.FillFormat.FillType = FillType.NoFill;
                            text.ShapeStyle.LineColor.Color = Color.White;
                            IPPImage img1 = tp1.Images.AddImage(ktlist[i].img);
                            int height = (img1.Height * 430 / img1.Width);
                            temp1[2].Shapes.AddPictureFrame(ShapeType.Rectangle, 20 + (670 - 440) / 2, 170, 440, height, img1);
                            t.AddClone(temp1[2]);
                        }
                    }
                }
            }
            return t;
        }

        /// <summary>
        /// 周度业态排名（全局数据、通用）
        /// </summary>
        /// <returns></returns>
        public virtual ISlideCollection _plus_jp_zdpm(string bamc,string [] yt)
        {
            #region 准备数据
            
            var data_zd = (from a in Cache_data_cjjl.bz.AsEnumerable()
                          where yt.Contains(a["yt"])
                          group a by new
                          {
                              lpmc = a["lpmc"], zt = a["zt"]
                          } into g
                          select new
                          {
                              lpmc = g.Key.lpmc,
                              zt = g.Key.zt,
                              cjts = g.Sum(m => m["ts"].ints()),
                              cjje = g.Sum(m => m["cjje"].longs()).je_y(),
                              jzmj = g.Sum(m => m["jzmj"].doubls()).mj(),
                              tnmj = g.Sum(m => m["tnmj"].doubls()).mj(),
                          }
                          into b orderby b.cjts descending select b ).Take(5).ToList();


            #endregion

            #region 生成页面

            if(data_zd!=null&data_zd.Count>0)
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Columns.Add("pm");
                dt.Columns.Add("lpmc");
                dt.Columns.Add("zt");
                dt.Columns.Add("cjts");
                dt.Columns.Add("cjmj");
                dt.Columns.Add("cjje");
                dt.Columns.Add("jmjj");
                dt.Columns.Add("tnjj");
                for (int i = 0; i < data_zd.Count; i++)
                {
                    DataRow dr = dt.NewRow();
                    dr["pm"] = i;
                    dr["lpmc"] = data_zd[i].lpmc;
                    dr["zt"] = data_zd[i].zt;
                    dr["cjts"] = data_zd[i].cjts;
                    dr["cjmj"] = data_zd[i].jzmj;
                    dr["cjje"] = data_zd[i].cjje;
                    dr["jmjj"] = data_zd[i].cjje/ data_zd[i].jzmj;
                    dr["tnjj"] = data_zd[i].cjje/ data_zd[i].tnmj;
                    dt.Rows.Add(dr);
                }
                var tp = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_ZDPM"]);
                var temp = tp.Slides;
                var page = temp[0];
                IAutoShape text1 = (IAutoShape)page.Shapes[1];
                text1.TextFrame.Text = string.Format(text1.TextFrame.Text, bamc, string.Join(",",yt));
                Office_Tables.SetJP_BASE_ZDYTPM_Table(page, dt, 2, null, null);

                IAutoShape text2 = (IAutoShape)page.Shapes[3];
                text2.TextFrame.Text = string.Format(text2.TextFrame.Text,  string.Join(",", yt),Base_date.GET_ZCMC(Base_date.bn,Base_date.bz));

                return temp;
            }
            #endregion
            return null;
        }

    }
}
