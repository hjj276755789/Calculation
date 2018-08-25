using Aspose.Slides;
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
                            if (item_jp.hxcs != null)
                            {
                                for (int i = 0; i < item_jp.hxcs.Length; i++)
                                {
                                    var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item_jp.hxcs[i]).FirstOrDefault();

                                    DataRow dr1 = jzgjt.NewRow();
                                    dr1[0] = item_jp.lpcs[0] + "(" + item.hxcs[i] + ")";
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
                              
                            }
                        }
                        Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                        t.AddClone(page);

                    }
                    #endregion

                    #region 别墅


                    else if (item.ytcs[0] == "别墅")
                    {
                        var page = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_JZGJ"]).Slides[0];
                        IAutoShape text1 = (IAutoShape)page.Shapes[2];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                        System.Data.DataTable jzgjt = new System.Data.DataTable();
                        jzgjt.Columns.Add("");
                        jzgjt.Columns.Add("成交套数", typeof(int));
                        jzgjt.Columns.Add("建面均价", typeof(double));
                        foreach (var item_jp in item.jpxmlb)
                        {
                            if (item_jp.xfytcs != null) { 
                                for (int i = 0; i < item_jp.xfytcs.Length; i++)
                                {

                                    var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item_jp.xfytcs[i]).FirstOrDefault();

                                    DataRow dr1 = jzgjt.NewRow();
                                    dr1[0] = item_jp.lpcs[0] + "(" + item.xfytcs[i] + ")";
                                    if (jpcjxx != null)
                                    {

                                        dr1[1] = jpcjxx["rgts"].ints();
                                        dr1[2] = jpcjxx["rgjmjj"].ints();
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
                                
                            }
                        }
                            Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                            t.AddClone(page);
                        

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

        public ISlideCollection _plus_jp_dyt_jzgj(JP_BA_INFO item)
        {
            try
            {
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);

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
                            if (item_jp.hxcs != null)
                            {
                                for (int i = 0; i < item_jp.hxcs.Length; i++)
                                {
                                    var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item_jp.hxcs[i]).FirstOrDefault();

                                    DataRow dr1 = jzgjt.NewRow();
                                    dr1[0] = item_jp.lpcs[0] + "(" + item.hxcs[i] + ")";
                                    if (jpcjxx != null)
                                    {

                                        dr1[1] = jpcjxx["xkts"].ints();
                                        dr1[2] = jpcjxx["xktnjj"].ints();
                                    }
                                    else
                                    {
                                        dr1[1] = 0;
                                        dr1[2] = 0;
                                    }
                                    jzgjt.Rows.Add(dr1);
                                }

                            }
                        }
                        Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                        t.AddClone(page);

                    }
                    #endregion

                #region 别墅


                    else if (item.ytcs[0] == "别墅")
                    {
                        var page = new Presentation(ConfigurationManager.AppSettings["PLUS_JP_JZGJ"]).Slides[0];
                        IAutoShape text1 = (IAutoShape)page.Shapes[2];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                        System.Data.DataTable jzgjt = new System.Data.DataTable();
                        jzgjt.Columns.Add("");
                        jzgjt.Columns.Add("成交套数", typeof(int));
                        jzgjt.Columns.Add("建面均价", typeof(double));
                        foreach (var item_jp in item.jpxmlb)
                        {
                            if (item_jp.xfytcs != null)
                            {
                                for (int i = 0; i < item_jp.xfytcs.Length; i++)
                                {

                                    var jpcjxx = Cache_data_rgsj.bz.AsEnumerable().Where(a => a["xm"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == item_jp.xfytcs[i]).FirstOrDefault();

                                    DataRow dr1 = jzgjt.NewRow();
                                    dr1[0] = item_jp.lpcs[0] + "(" + item.xfytcs[i] + ")";
                                    if (jpcjxx != null)
                                    {

                                        dr1[1] = jpcjxx["xkts"].ints();
                                        dr1[2] = jpcjxx["xktnjj"].ints();
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

                            }
                        }
                        Office_Charts.Chart_jp_fudi_chart1(page, jzgjt, 3);
                        t.AddClone(page);


                    }


                    #endregion

                #region 大业态


                    else
                    {
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
                                dr1[2] = jpcjxx["xktnjj"].ints();
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
                
                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log(e.Message);
                return null;
            }
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
        /// 大业态推广图片
        /// </summary>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public ISlideCollection _plus_jp_dyt_tgtp(JP_BA_INFO item)
        {
            string path = ConfigurationManager.AppSettings["DgPath"] + Base_date.bn + "\\" + Base_date.bz;

            var t = new Presentation().Slides;
            t.RemoveAt(0);

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


       

        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                                DataRow temp_ba_bz,
                                DataRow temp_ba_sz,
                                EnumerableRowCollection<DataRow> temp_cjba_bz, 
                                EnumerableRowCollection<DataRow> temp_cjba_sz, 
                                JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {
               

                if (Base_Config_Rgsj._认购数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Rgsj.上周_新开套数:
                        case Base_Config_Rgsj.上周_新开销售套数:
                        case Base_Config_Rgsj.上周_新开建面均价:
                        case Base_Config_Rgsj.上周_新开套内均价:
                        case Base_Config_Rgsj.上周_认购套数:
                        case Base_Config_Rgsj.上周_认购套内体量:
                        case Base_Config_Rgsj.上周_认购套内均价:
                        case Base_Config_Rgsj.上周_认购建面体量:
                        case Base_Config_Rgsj.上周_认购建面均价:
                        case Base_Config_Rgsj.上周_认购金额:
                            {
                                if (temp_ba_sz != null)
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_sz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                }
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;

                            
                        case Base_Config_Rgsj.本周_新开套数:
                        case Base_Config_Rgsj.本周_新开销售套数:
                        case Base_Config_Rgsj.本周_新开建面均价:
                        case Base_Config_Rgsj.本周_新开套内均价:
                        case Base_Config_Rgsj.本周_认购套数:
                        case Base_Config_Rgsj.本周_认购套内体量:
                        case Base_Config_Rgsj.本周_认购套内均价:
                        case Base_Config_Rgsj.本周_认购建面体量:
                        case Base_Config_Rgsj.本周_认购建面均价:
                        case Base_Config_Rgsj.本周_认购金额:
                            {
                                if (temp_ba_bz != null)
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                }
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                        default: {
                                if (temp_ba_bz != null)
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName];
                                }
                                else
                                    dr1[dt.Columns[j].ColumnName] = "";
                            }; break;

                    }
                }
                else if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case Base_Config_Cjba.本周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz !=null? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()):0; }; break;
                            case Base_Config_Cjba.本周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()):0; }; break;
                            case Base_Config_Cjba.本周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()):0; }; break;
                            case Base_Config_Cjba.本周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].ints()):0; }; break;
                            case Base_Config_Cjba.本周_建面均价: {
                            
                                    if ((temp_cjba_bz != null && temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建面均价._ConfigCjbaMc()].doubls()) != 0))
                                        dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                    else
                                    {
                                        dr1[dt.Columns[j].ColumnName] = "";
                                    }
                            }; break;
                            case Base_Config_Cjba.本周_套内均价:
                                {
                                if((temp_cjba_bz != null && temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                            case Base_Config_Cjba.上周_备案套数: { dr1[dt.Columns[j].ColumnName] =  temp_cjba_sz != null ?  temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.上周_成交金额: { dr1[dt.Columns[j].ColumnName] =  temp_cjba_sz != null ?  temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) : 0; }; break;
                            case Base_Config_Cjba.上周_建筑面积: { dr1[dt.Columns[j].ColumnName] =  temp_cjba_sz != null ?  temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()) : 0; }; break;
                            case Base_Config_Cjba.上周_套内面积: { dr1[dt.Columns[j].ColumnName] =  temp_cjba_sz != null ?  temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].ints()) : 0; }; break;
                            case Base_Config_Cjba.上周_建面均价: {
                                if ((temp_cjba_sz != null && temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                            case Base_Config_Cjba.上周_套内均价:
                            {
                                if ((temp_cjba_sz != null && temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                            default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                        }
                   
                    
                }
                else if (Base_Config_Jzgj._竞争格局参数名称.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Jzgj.项目名称:
                            {
                                dr1[dt.Columns[j].ColumnName] = item.lpcs[0];
                            }; break;
                        case Base_Config_Jzgj.业态:
                            {
                                dr1[dt.Columns[j].ColumnName] = yt;
                            }; break;
                        case Base_Config_Jzgj.组团:
                            {
                                dr1[dt.Columns[j].ColumnName] = string.Join(",", item.ztcs);
                            }; break;
                        case Base_Config_Jzgj.竞争格局_主力面积区间:
                            {
                                dr1[dt.Columns[j].ColumnName] = item.zlmjqj;
                            }; break;
                        case Base_Config_Jzgj.竞争格局名称:
                            {
                                dr1[dt.Columns[j].ColumnName] = item.jzgjmc;
                            }; break;
                    }

                }
            }
            
                return dr1;
        }
        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                                DataRow temp_ba_bz,
                                EnumerableRowCollection<DataRow> temp_cjba_bz,
                                EnumerableRowCollection<DataRow> temp_cjba_sz, 
                                JP_JPXM_INFO item)
        {
            return GET_ROW(yt, dr1, dt, temp_ba_bz, null, temp_cjba_bz, temp_cjba_sz, item);
        }
        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                                DataRow temp_ba_bz,
                                EnumerableRowCollection<DataRow> temp_cjba_bz,
                                JP_JPXM_INFO item)
        {
            return GET_ROW(yt, dr1, dt, temp_ba_bz, null, temp_cjba_bz, null, item);
        }
        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                                DataRow temp_ba_bz,
                                DataRow temp_ba_sz,
                                EnumerableRowCollection<DataRow> temp_cjba_bz,
                                JP_JPXM_INFO item)
        {
            return GET_ROW(yt, dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, null, item);
        }




        public DataRow GET_ROW_BA(string yt, DataRow dr1, System.Data.DataTable dt,
                               DataRow temp_ba_bz,
                               DataRow temp_ba_sz,
                               EnumerableRowCollection<DataRow> temp_cjba_bz,
                               EnumerableRowCollection<DataRow> temp_cjba_sz,
                               JP_BA_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {


                if (Base_Config_Rgsj._认购数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Rgsj.上周_新开套数:
                        case Base_Config_Rgsj.上周_新开销售套数:
                        case Base_Config_Rgsj.上周_新开建面均价:
                        case Base_Config_Rgsj.上周_新开套内均价:
                        case Base_Config_Rgsj.上周_认购套数:
                        case Base_Config_Rgsj.上周_认购套内体量:
                        case Base_Config_Rgsj.上周_认购套内均价:
                        case Base_Config_Rgsj.上周_认购建面体量:
                        case Base_Config_Rgsj.上周_认购建面均价:
                        case Base_Config_Rgsj.上周_认购金额:
                            {
                                if (temp_ba_sz != null)
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_sz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                }
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;


                        case Base_Config_Rgsj.本周_新开套数:
                        case Base_Config_Rgsj.本周_新开销售套数:
                        case Base_Config_Rgsj.本周_新开建面均价:
                        case Base_Config_Rgsj.本周_新开套内均价:
                        case Base_Config_Rgsj.本周_认购套数:
                        case Base_Config_Rgsj.本周_认购套内体量:
                        case Base_Config_Rgsj.本周_认购套内均价:
                        case Base_Config_Rgsj.本周_认购建面体量:
                        case Base_Config_Rgsj.本周_认购建面均价:
                        case Base_Config_Rgsj.本周_认购金额:
                            {
                                if (temp_ba_bz != null)
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()];
                                }
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                        default:
                            {
                                if (temp_ba_bz != null)
                                {
                                    dr1[dt.Columns[j].ColumnName] = temp_ba_bz[dt.Columns[j].ColumnName];
                                }
                                else
                                    dr1[dt.Columns[j].ColumnName] = "";
                            }; break;

                    }
                }
                else if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Cjba.本周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba.本周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) : 0; }; break;
                        case Base_Config_Cjba.本周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建筑面积._ConfigCjbaMc()].doubls()) : 0; }; break;
                        case Base_Config_Cjba.本周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_bz != null ? temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba.本周_建面均价:
                            {

                                if ((temp_cjba_bz != null && temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_建面均价._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                        case Base_Config_Cjba.本周_套内均价:
                            {
                                if ((temp_cjba_bz != null && temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_bz.Sum(m => m[Base_Config_Cjba.本周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                        case Base_Config_Cjba.上周_备案套数: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_备案套数._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba.上周_成交金额: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) : 0; }; break;
                        case Base_Config_Cjba.上周_建筑面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()) : 0; }; break;
                        case Base_Config_Cjba.上周_套内面积: { dr1[dt.Columns[j].ColumnName] = temp_cjba_sz != null ? temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].ints()) : 0; }; break;
                        case Base_Config_Cjba.上周_建面均价:
                            {
                                if ((temp_cjba_sz != null && temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_建筑面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                        case Base_Config_Cjba.上周_套内均价:
                            {
                                if ((temp_cjba_sz != null && temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls()) != 0))
                                    dr1[dt.Columns[j].ColumnName] = (temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_成交金额._ConfigCjbaMc()].longs()) / temp_cjba_sz.Sum(m => m[Base_Config_Cjba.上周_套内面积._ConfigCjbaMc()].doubls())).je_y();
                                else
                                {
                                    dr1[dt.Columns[j].ColumnName] = "";
                                }
                            }; break;
                        default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                    }


                }
                else if (Base_Config_Jzgj._竞争格局参数名称.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Jzgj.项目名称:
                            {
                                dr1[dt.Columns[j].ColumnName] = item.lpcs[0];
                            }; break;
                        case Base_Config_Jzgj.业态:
                            {
                                dr1[dt.Columns[j].ColumnName] = yt;
                            }; break;
                        case Base_Config_Jzgj.组团:
                            {
                                dr1[dt.Columns[j].ColumnName] = string.Join(",", item.ztcs);
                            }; break;
                        case Base_Config_Jzgj.竞争格局_主力面积区间:
                            {
                                dr1[dt.Columns[j].ColumnName] = item.zlmjqj;
                            }; break;
                        case Base_Config_Jzgj.竞争格局名称:
                            {
                                dr1[dt.Columns[j].ColumnName] = "本案";
                            }; break;
                    }

                }
            }

            return dr1;
        }
        public DataRow GET_ROW_BA(string yt, DataRow dr1, System.Data.DataTable dt,
                                DataRow temp_ba_bz,
                                EnumerableRowCollection<DataRow> temp_cjba_bz,
                                EnumerableRowCollection<DataRow> temp_cjba_sz,
                                JP_BA_INFO item)
        {
            return GET_ROW_BA(yt, dr1, dt, temp_ba_bz, null, temp_cjba_bz, temp_cjba_sz, item);
        }
        public DataRow GET_ROW_BA(string yt, DataRow dr1, System.Data.DataTable dt,
                                DataRow temp_ba_bz,
                                EnumerableRowCollection<DataRow> temp_cjba_bz,
                                JP_BA_INFO item)
        {
            return GET_ROW_BA(yt, dr1, dt, temp_ba_bz, null, temp_cjba_bz, null, item);
        }
        public DataRow GET_ROW_BA(string yt, DataRow dr1, System.Data.DataTable dt,
                                DataRow temp_ba_bz,
                                DataRow temp_ba_sz,
                                EnumerableRowCollection<DataRow> temp_cjba_bz,
                                JP_BA_INFO item)
        {
            return GET_ROW_BA(yt, dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, null, item);
        }




        public ISlideCollection ztzdpm(string str,int[] index1,int[] index2,int [] index3, int[] index4, string qy )
        {
            var p = new Presentation();
            var t = p.Slides;
            t.RemoveAt(0);
            var pages = new Presentation(str).Slides;
            var jbz = pages[index1[0]];
            
            #region 近8周江北区住宅市场环境
            System.Data.DataTable zzsc = new System.Data.DataTable();
            zzsc.Columns.Add("时间");
            zzsc.Columns.Add("预售新增供应量（单位: 万㎡）");
            zzsc.Columns.Add("成交量（单位: 万㎡）");
            zzsc.Columns.Add("建面均价（元 /㎡）");
            var jbz_cjba = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                            where a["qy"].ToString() == qy && (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                            group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                            select new
                            {
                                zc = s.Key.zc,
                                zcmc = s.Key.zcmc,
                                cjje = s.Sum(a => a["cjje"].longs()),
                                jzmj = s.Sum(a => a["jzmj"].doubls()),
                            }).OrderBy(m => m.zc).ToList();
            var jbz_xzys = (from a in Cache_data_xzys.jbz.AsEnumerable()
                            where a["qx1"].ToString() == qy && (a["tyyt"].ToString() == "别墅" || a["tyyt"].ToString() == "高层" || a["tyyt"].ToString() == "小高层" || a["tyyt"].ToString() == "洋房" || a["tyyt"].ToString() == "洋楼")
                            group a by new { zc = a["zc"] } into s
                            select new
                            {
                                zc = s.Key.zc,
                                xzgy = s.Sum(a => a["jzmj"].doubls()),
                            }).OrderBy(m => m.zc).ToList();
            var temp6 = (from a in jbz_cjba
                         join b in jbz_xzys on a.zc equals b.zc into temp
                         from tt in temp.DefaultIfEmpty()
                         select new
                         {
                             zc = a.zc,
                             zcmc = a.zcmc,
                             xzgyl = tt == null ? 0 : tt.xzgy,//这里主要第二个集合有可能为空。需要判断
                             cjmj = a.jzmj,
                             jmjj = a.cjje / a.jzmj
                         }).ToList();
            for (int i = 0; i < temp6.Count(); i++)
            {
                DataRow dr = zzsc.NewRow();
                dr[0] = temp6[i].zcmc;
                dr[1] = temp6[i].xzgyl.mj_wf();
                dr[2] = temp6[i].cjmj.mj_wf();
                dr[3] = temp6[i].jmjj.je_y();
                zzsc.Rows.Add(dr);
            }
            Office_Charts.Chart_gxfx(jbz, zzsc,index1[1]);
            if (index3 != null)
            {
                var data_bz = temp6.FirstOrDefault(m => m.zc.ints() == Base_date.bz);
                var data_sz = temp6.FirstOrDefault(m => m.zc.ints() == Base_date.bz-1);
                IAutoShape qyhj_txt_1 = (IAutoShape)jbz.Shapes[index3[0]];
                qyhj_txt_1.TextFrame.Text = string.Format(qyhj_txt_1.TextFrame.Text, qy);


                IAutoShape qyhj_txt_2 = (IAutoShape)jbz.Shapes[index3[1]];
                qyhj_txt_2.TextFrame.Text = string.Format(qyhj_txt_2.TextFrame.Text, data_bz.xzgyl.mj_wf(), data_bz.cjmj.mj_wf(), data_bz.jmjj.mj_wf());
                IAutoShape qyhj_txt_3 = (IAutoShape)jbz.Shapes[index3[2]];
                qyhj_txt_3.TextFrame.Text = string.Format(qyhj_txt_3.TextFrame.Text, 
                    data_bz.xzgyl.mj_wf(), ((data_bz.xzgyl - data_sz.xzgyl)/ data_sz.xzgyl).ss_bfb(),
                    data_bz.cjmj.mj_wf(), ((data_bz.cjmj - data_sz.cjmj) / data_sz.cjmj).ss_bfb(),
                     data_bz.jmjj.je_y(), ((data_bz.jmjj - data_sz.jmjj) / data_sz.jmjj).ss_bfb()
                    );
            }
            t.AddClone(jbz);
            #endregion

            #region 江北区周度住宅排名
            var temp_data_cj = from a in Cache_data_cjjl.bz.AsEnumerable()
                               where a["qy"].ToString() == qy && (a["yt"].ToString() == "别墅" || a["yt"].ToString() == "高层" || a["yt"].ToString() == "小高层" || a["yt"].ToString() == "洋房" || a["yt"].ToString() == "洋楼")
                               group a by new { lpmc = a["lpmc"] } into d
                               select new
                               {
                                   lpmc = d.Key.lpmc,
                                   cjts = d.Sum(m => m["ts"].ints()),
                                   cjtl = d.Sum(m => m["jzmj"].doubls()),
                                   cjje = d.Sum(m => m["cjje"].doubls())
                               };
            var cjpm_ts = temp_data_cj.OrderByDescending(m => m.cjts).Take(10).ToList();
            var cjpm_mj = temp_data_cj.OrderByDescending(m => m.cjtl).Take(10).ToList();
            var cjpm_je = temp_data_cj.OrderByDescending(m => m.cjje).Take(10).ToList();
            System.Data.DataTable cjpm = new System.Data.DataTable();
            cjpm.Columns.Add("序号");
            cjpm.Columns.Add("项目名称1");
            cjpm.Columns.Add("套数");
            cjpm.Columns.Add("项目名称2");
            cjpm.Columns.Add("成交面积");
            cjpm.Columns.Add("项目名称3");
            cjpm.Columns.Add("成交金额");
            for (int i = 0; i < 10; i++)
            {
                DataRow dr = cjpm.NewRow();
                dr["序号"] = i + 1;
                if (cjpm_ts.Count() > i)
                {
                    dr["项目名称1"] = cjpm_ts[i].lpmc;
                    dr["套数"] = cjpm_ts[i].cjts;
                }
                else
                {
                    dr["项目名称"] = "";
                    dr["套数"] = "";
                }

                if (cjpm_mj.Count() > i)
                {
                    dr["项目名称2"] = cjpm_ts[i].lpmc;
                    dr["成交面积"] = cjpm_mj[i].cjtl.ints();
                }
                else
                {
                    dr["项目名称2"] = "";
                    dr["成交面积"] = "";
                }

                if (cjpm_je.Count() > i)
                {
                    dr["项目名称3"] = cjpm_ts[i].lpmc;
                    dr["成交金额"] = cjpm_je[i].cjje.je_wy();
                }
                else
                {
                    dr["项目名称3"] = "";
                    dr["成交金额"] = "";
                }
                cjpm.Rows.Add(dr);
            }
            var cjpmp_page = pages[index2[0]];


            Office_Tables.SetChart(cjpmp_page, cjpm, index2[1], null, null);
            if(index4!=null)
            {
                IAutoShape test1 = (IAutoShape)cjpmp_page.Shapes[index4[0]];
                test1.TextFrame.Text = string.Format(test1.TextFrame.Text, qy);

                IAutoShape test2 = (IAutoShape)cjpmp_page.Shapes[index4[1]];
                test2.TextFrame.Text = string.Format(test2.TextFrame.Text, Base_date.GET_ZCMC(Base_date.bn, Base_date.bz));

            }
            else
            {
                IAutoShape cjpmwz = (IAutoShape)cjpmp_page.Shapes[2];
                cjpmwz.TextFrame.Text = string.Format(cjpmwz.TextFrame.Text, Base_date.GET_ZCMC(Base_date.bn, Base_date.bz));
            }
            

            t.AddClone(cjpmp_page);
            #endregion

            return t;
        }

        public ISlideCollection ztzdpm(string str, int[] index1, int[] index2, string qy)
        {
           return  ztzdpm(str, index1, index2, null,null, qy);
        }
    }
}
