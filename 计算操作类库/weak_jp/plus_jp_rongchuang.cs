using Aspose.Slides;
using Calculation.Base;
using Calculation.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    public class plus_jp_rongchuang: plus_jp_base
    {

        public ISlideCollection _plus_jp_rongchuang_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);


                #region 

                foreach (var item in param)
                {
                    var tp = new Presentation(str);
                    var temp = tp.Slides;
                  

                    if(item.ytcs.Contains("商业"))
                    {
                        var page = temp[1];
                        IAutoShape text1 = (IAutoShape)page.Shapes[0];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);

                        dt.Columns.Add(Base_Config_Cjba.上周_备案套数);   //上周认购套数
                        dt.Columns.Add(Base_Config_Cjba.上周_建筑面积); //上周认购套内均价

                        dt.Columns.Add(Base_Config_Cjba.上周_成交金额);   //上周认购套数
                        dt.Columns.Add(Base_Config_Cjba.上周_建面均价); //上周认购套内均价


                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数);   //上周认购套数
                        dt.Columns.Add(Base_Config_Cjba.本周_建筑面积); //上周认购套内均价
                        dt.Columns.Add(Base_Config_Cjba.本周_成交金额);   //上周认购套数
                        dt.Columns.Add(Base_Config_Cjba.本周_建面均价); //上周认购套内均价
                        dt.Columns.Add("备注"); //上周认购套内均价
                        dt = GET_JPBA_BX(dt, item);
                        if(item.jpxmlb!=null&&item.jpxmlb.Count>0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_RONGCHUANG_2_Table(page, dt, 1, null, null);
                            t.AddClone(page);
                        }
                    }
                    else if(item.ytcs.Contains("车库"))
                    {
                        var page = temp[2];
                        IAutoShape text1 = (IAutoShape)page.Shapes[0];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);

                        dt.Columns.Add(Base_Config_Cjba.上周_备案套数);   //上周认购套数
                        dt.Columns.Add(Base_Config_Cjba.上周_建筑面积); //上周认购套内均价

                        dt.Columns.Add(Base_Config_Cjba.上周_成交金额);   //上周认购套数
                        dt.Columns.Add(Base_Config_Cjba.上周_套均总价); //上周认购套内均价


                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数);   //上周认购套数
                        dt.Columns.Add(Base_Config_Cjba.本周_建筑面积); //上周认购套内均价
                        dt.Columns.Add(Base_Config_Cjba.本周_成交金额);   //上周认购套数
                        dt.Columns.Add(Base_Config_Cjba.本周_套均总价); //上周认购套内均价
                        dt = GET_JPBA_BX(dt, item);
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_RONGCHUANG_3_Table(page, dt, 1, null, null);
                            t.AddClone(page);
                        }
                    }
                    else
                    {
                        var page = temp[0];
                        IAutoShape text1 = (IAutoShape)page.Shapes[0];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt.Columns.Add(Base_Config_Jzgj.业态);

                        dt.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间); 
                        dt.Columns.Add("总面积段");


                        dt.Columns.Add(Base_Config_Cjba.上周_备案套数);   //上周认购套数
                        dt.Columns.Add(Base_Config_Cjba.上周_建面均价); //上周认购套内均价

                        dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);   //上周认购套数
                        dt.Columns.Add(Base_Config_Rgsj.上周_认购建面均价); //上周认购套内均价

                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数);   //上周认购套数
                        dt.Columns.Add(Base_Config_Cjba.本周_建面均价); //上周认购套内均价
                                                      
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);   //上周认购套数
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面均价); //上周认购套内均价

                        dt.Columns.Add("备注");  //变化原因
                        dt.Columns.Add(Base_Config_Rgsj.本周_营销动作);
                        dt.Columns.Add(Base_Config_Rgsj.本周_优惠);
                        dt = GET_JPBA_BX(dt, item);
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_RONGCHUANG_1_Table(page, dt, 1, null, null);
                            t.AddClone(page);
                        }
                    }
                       
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
        /// 竞品项目表现 本案与认购数据
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="jpxm"></param>
        /// <returns></returns>
        public System.Data.DataTable GET_JPXM_BX(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {
                if (item.ytcs[0] == "别墅")
                {
                    if (item.xfytcs != null)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {

                            DataRow dr1 = dt.NewRow();
                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);

                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态  
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        //本周本案认购数据
                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else if (item.ytcs[0] == "商务")
                {

                    //商务业态：如果户型参数为空，则使用细分业态参数，若细分业态参数业为空，直接使用主业态：商务
                    if (item.hxcs != null && item.hxcs.Length > 0)
                    {
                        for (int i = 0; i < item.hxcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[i]);

                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[i]);


                            //本周本案认购数据
                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.hxcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                        }
                    }
                    else if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态

                            var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                            var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);

                            var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);


                            //本周本案认购数据
                            var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                            var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);

                        //本周本案认购数据
                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else if (item.ytcs[0] == "商铺")
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态

                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);


                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);

                    //本周本案认购数据
                    var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));

                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    //本周本案认购数据
                    var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));


                }
            }


            return dt;
        }
        public System.Data.DataTable GET_JPBA_BX(System.Data.DataTable dt, JP_BA_INFO item)
        {

            if (item.ytcs[0] == "别墅")
            {
                if (item.xfytcs != null && item.xfytcs.Contains("别墅"))
                {
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {

                        DataRow dr1 = dt.NewRow();
                        #region 数据准备
                        //竞品业态
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);


                        var temp_rgsj_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_rgsj_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        #endregion

                        dt.Rows.Add(GET_ROW_BA(item.xfytcs[i], dr1, dt, null, null, temp_cjba_bz, temp_cjba_sz, item));

                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态  
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    var temp_rgsj_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                    //本周本案认购数据
                    #endregion

                    dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, null, null, temp_cjba_bz, temp_cjba_sz, item));
                }
            }
            else if (item.ytcs[0] == "商务")
            {

                //商务业态：如果户型参数为空，则使用细分业态参数，若细分业态参数业为空，直接使用主业态：商务
                if (item.hxcs != null && item.hxcs.Length > 0)
                {
                    for (int i = 0; i < item.hxcs.Length; i++)
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[0]);

                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[0]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.hxcs[0]);


                        //本周本案认购数据
                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW_BA(item.hxcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else if (item.xfytcs != null && item.xfytcs.Length > 0)
                {
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态

                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[0]);
                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[0]);

                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[0]);
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[0]);


                        //本周本案认购数据
                        var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                        #endregion

                        dt.Rows.Add(GET_ROW_BA(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态

                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);


                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);

                    //本周本案认购数据
                    var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                    #endregion

                    dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
                }
            }
            else if (item.ytcs[0] == "商业")
            {
                DataRow dr1 = dt.NewRow();

                #region 数据准备
                //竞品业态

                var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                var temp_rgsj_sz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);


                var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);
                var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.ytcs[0]);

                //本周本案认购数据
                var temp_rg_bz = temp_rgsj_bz.FirstOrDefault();
                var temp_rg_sz = temp_rgsj_sz.FirstOrDefault();

                #endregion

                dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, temp_rg_bz, temp_rg_sz, temp_cjba_bz, temp_cjba_sz, item));
            }
            else
            {
                DataRow dr1 = dt.NewRow();

                #region 数据准备
                //竞品业态
                var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);

                var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                //本周本案认购数据
                var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                #endregion

                dt.Rows.Add(GET_ROW_BA(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_cjba_bz, temp_cjba_sz, item));
            }

            return dt;
        }
    }
}
