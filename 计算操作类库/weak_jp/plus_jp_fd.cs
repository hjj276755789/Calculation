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
    public class plus_jp_fd : weak
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
        ///  大业态循环
        /// </summary>
        /// <param name="str"></param>
        /// <param name="cjbh"></param>
        /// <returns></returns> 
        public ISlideCollection _plus_jp_fudi_4(string str, int cjbh)
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
                    var tp = new Presentation(str);
                    var temp = tp.Slides;
                    var page = temp[0];
                    IAutoShape text1 = (IAutoShape)page.Shapes[2];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                    //数据
                    System.Data.DataTable jzgjt = new System.Data.DataTable();
                    jzgjt.Columns.Add("");
                    jzgjt.Columns.Add("成交套数", typeof(int));
                    jzgjt.Columns.Add("建面均价", typeof(double));
                    //图表
                    IChart chart = (IChart)page.Shapes[3];
                    #region 本案
                    var bacjxx = Cache_data_cjjl.bz.AsEnumerable().Where(a => a["lpmc"].ToString() == item.lpcs[0] && a["yt"].ToString() == item.ytcs[0]);

                    DataRow dr = jzgjt.NewRow();
                    dr[0] = item.lpcs[0] + item.ytcs[0];
                    dr[1] = bacjxx.Sum(m => m["ts"].ints());
                    dr[2] = bacjxx.Sum(m => m["cjje"].ints()) / bacjxx.Sum(m => m["jzmj"].doubls());
                    jzgjt.Rows.Add(dr);
                    #endregion
                    #region 竞争项目
                    foreach (var item_jp in item.jpxmlb)
                    {
                        string jpyt = item_jp.ytcs == null ? item.ytcs[0] : item_jp.ytcs[0];
                        var jpcjxx = Cache_data_cjjl.bz.AsEnumerable().Where(a => a["lpmc"].ToString() == item_jp.lpcs[0] && a["yt"].ToString() == jpyt);

                        DataRow dr1 = jzgjt.NewRow();
                        dr1[0] = item_jp.lpcs[0] + "(" + item.ytcs[0] + ")";
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
                string path = ConfigurationManager.AppSettings["DgPath"] + Base_date.bn + "\\" + Base_date.bz;

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
                        Base_Log.Log(Path.Combine(path, item.lpcs[0] + ".jpg") + "文件不存在");
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
                            Base_Log.Log(Path.Combine(path, item.lpcs[0] + ".jpg") + "文件不存在");
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
                                    var tp1 = new Presentation(str);
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
                                    var tp1 = new Presentation(str);
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
                                    var tp1 = new Presentation(str);
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
                                var tp1 = new Presentation(str);
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
        public ISlideCollection _plus_jp_fudi_5(string str, int cjbh)
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
                    if(item.ytcs[0]=="别墅"|| item.ytcs[0] == "商务") {
                        if (item.xfytcs != null && item.xfytcs.Length > 0)
                        {
                            for (int i = 0; i < item.xfytcs.Length; i++)
                            {
                                var tp = new Presentation(str);
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
                #region P3
                //string path = ConfigurationManager.AppSettings["DgPath"] + Base_date.bn + "\\" + Base_date.bz;

                //foreach (var item in param)
                //{

                //    List<Zb_Jp_Tgtp_Model> tgtplb = new List<Zb_Jp_Tgtp_Model>();
                //    try
                //    {
                //        Image img = (Image)new Bitmap(Path.Combine(path, item.lpcs[0] + ".jpg"));
                //        if ((img.Width / 1.0) / img.Height < zd)
                //        {
                //            Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                //            tgtp.img = img;
                //            tgtp.xmmc = item.lpcs[0];
                //            tgtp.tplx = Models.Enums.TP_LX.窄图;
                //            tgtplb.Add(tgtp);
                //        }
                //        else if ((img.Width / 1.0) / img.Height > zd && (img.Width / 1.0) / img.Height < kd)
                //        {
                //            Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                //            tgtp.img = img;
                //            tgtp.xmmc = item.lpcs[0];
                //            tgtp.tplx = Models.Enums.TP_LX.方图;
                //            tgtplb.Add(tgtp);
                //        }
                //        else
                //        {
                //            Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                //            tgtp.img = img;
                //            tgtp.xmmc = item.lpcs[0];
                //            tgtp.tplx = Models.Enums.TP_LX.宽图;
                //            tgtplb.Add(tgtp);
                //        }
                //    }
                //    catch
                //    {
                //        Base_Log.Log(Path.Combine(path, item.lpcs[0] + ".jpg") + "文件不存在");
                //    }
                //    foreach (var item_jp in item.jpxmlb)
                //    {
                //        try
                //        {
                //            Image img = (Image)new Bitmap(Path.Combine(path, item_jp.lpcs[0] + ".jpg"));
                //            if ((img.Width / 1.0) / img.Height < zd)
                //            {
                //                Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                //                tgtp.img = img;
                //                tgtp.xmmc = item_jp.lpcs[0];
                //                tgtp.tplx = Models.Enums.TP_LX.窄图;
                //                tgtplb.Add(tgtp);
                //            }
                //            else if ((img.Width / 1.0) / img.Height > zd && (img.Width / 1.0) / img.Height < kd)
                //            {
                //                Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                //                tgtp.img = img;
                //                tgtp.xmmc = item_jp.lpcs[0];
                //                tgtp.tplx = Models.Enums.TP_LX.方图;
                //                tgtplb.Add(tgtp);
                //            }
                //            else
                //            {
                //                Zb_Jp_Tgtp_Model tgtp = new Zb_Jp_Tgtp_Model();
                //                tgtp.img = img;
                //                tgtp.xmmc = item_jp.lpcs[0];
                //                tgtp.tplx = Models.Enums.TP_LX.宽图;
                //                tgtplb.Add(tgtp);
                //            }
                //        }
                //        catch
                //        {
                //            Base_Log.Log(Path.Combine(path, item.lpcs[0] + ".jpg") + "文件不存在");
                //        }
                //    }
                //    if (tgtplb.Count > 0)
                //    {
                //        List<Zb_Jp_Tgtp_Model> zt_pic = new List<Zb_Jp_Tgtp_Model>();
                //        List<Zb_Jp_Tgtp_Model> ft_pic = new List<Zb_Jp_Tgtp_Model>();
                //        var zt = tgtplb.Where(m => m.tplx == Models.Enums.TP_LX.窄图);
                //        var ft = tgtplb.Where(m => m.tplx == Models.Enums.TP_LX.方图);
                //        var kt = tgtplb.Where(m => m.tplx == Models.Enums.TP_LX.宽图);
                //        if (zt != null && zt.Count() > 0)
                //        {
                //            var ztlist = zt.ToList();
                //            for (int i = 0; i < ztlist.Count; i++)
                //            {
                //                zt_pic.Add(ztlist[i]);
                //                if ((i + 1) % 2 == 0 || i + 1 >= ztlist.Count)
                //                {
                //                    var tp1 = new Presentation(str);
                //                    var temp1 = tp1.Slides;
                //                    for (int j = 0; j < zt_pic.Count; j++)
                //                    {
                //                        IAutoShape text = temp1[2].Shapes.AddAutoShape(ShapeType.Rectangle, 20 + (220 * j), 130, 210, 40);
                //                        text.TextFrame.Text = zt_pic[j].xmmc;
                //                        text.ShapeStyle.FontColor.Color = Color.Black;
                //                        text.FillFormat.FillType = FillType.NoFill;
                //                        text.ShapeStyle.LineColor.Color = Color.White;
                //                        IPPImage img1 = tp1.Images.AddImage(zt_pic[j].img);
                //                        int height = (img1.Height * 210 / img1.Width);
                //                        temp1[2].Shapes.AddPictureFrame(ShapeType.Rectangle, 20 + (220 * j), 170, 210, height, img1);
                //                    }
                //                    t.AddClone(temp1[2]);
                //                    zt_pic.Clear();
                //                }
                //            }
                //        }
                //        if (ft != null && ft.Count() > 0)
                //        {
                //            var ftlist = ft.ToList();
                //            for (int i = 0; i < ftlist.Count; i++)
                //            {

                //                ft_pic.Add(ftlist[i]);
                //                if ((i + 1) % 2 == 0)
                //                {
                //                    var tp1 = new Presentation(str);
                //                    var temp1 = tp1.Slides;
                //                    for (int j = 0; j < ft_pic.Count; j++)
                //                    {
                //                        IAutoShape text = temp1[2].Shapes.AddAutoShape(ShapeType.Rectangle, 20 + (280 * j), 130, 210, 40);
                //                        text.TextFrame.Text = ft_pic[j].xmmc;
                //                        text.ShapeStyle.FontColor.Color = Color.Black;
                //                        text.FillFormat.FillType = FillType.NoFill;
                //                        text.ShapeStyle.LineColor.Color = Color.White;
                //                        IPPImage img1 = tp1.Images.AddImage(ft_pic[j].img);
                //                        int height = (img1.Height * 270 / img1.Width);
                //                        temp1[2].Shapes.AddPictureFrame(ShapeType.Rectangle, 20 + (280 * j), 170, 270, height, img1);
                //                    }
                //                    t.AddClone(temp1[2]);
                //                    ft_pic.Clear();
                //                }
                //                else if (i + 1 >= ftlist.Count)
                //                {
                //                    var tp1 = new Presentation(str);
                //                    var temp1 = tp1.Slides;
                //                    for (int j = 0; j < ft_pic.Count; j++)
                //                    {
                //                        IAutoShape text = temp1[2].Shapes.AddAutoShape(ShapeType.Rectangle, 20 + (670 - 280) / 2, 130, 210, 40);
                //                        text.TextFrame.Text = ft_pic[j].xmmc;
                //                        text.ShapeStyle.FontColor.Color = Color.Black;
                //                        text.FillFormat.FillType = FillType.NoFill;
                //                        text.ShapeStyle.LineColor.Color = Color.White;
                //                        IPPImage img1 = tp1.Images.AddImage(ft_pic[j].img);
                //                        int height = (img1.Height * 270 / img1.Width);
                //                        temp1[2].Shapes.AddPictureFrame(ShapeType.Rectangle, 20 + (670 - 280) / 2, 170, 270, height, img1);
                //                    }
                //                    t.AddClone(temp1[2]);
                //                    ft_pic.Clear();
                //                }
                //            }
                //        }
                //        if (kt != null && kt.Count() > 0)
                //        {
                //            var ktlist = kt.ToList();
                //            for (int i = 0; i < ktlist.Count; i++)
                //            {
                //                var tp1 = new Presentation(str);
                //                var temp1 = tp1.Slides;
                //                IAutoShape text = temp1[2].Shapes.AddAutoShape(ShapeType.Rectangle, 20 + (670 - 440) / 2, 130, 440, 40);
                //                text.TextFrame.Text = ktlist[i].xmmc;
                //                text.ShapeStyle.FontColor.Color = Color.Black;
                //                text.FillFormat.FillType = FillType.NoFill;
                //                text.ShapeStyle.LineColor.Color = Color.White;
                //                IPPImage img1 = tp1.Images.AddImage(ktlist[i].img);
                //                int height = (img1.Height * 430 / img1.Width);
                //                temp1[2].Shapes.AddPictureFrame(ShapeType.Rectangle, 20 + (670 - 440) / 2, 170, 440, height, img1);
                //                t.AddClone(temp1[2]);
                //            }
                //        }
                //    }
                //}


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
