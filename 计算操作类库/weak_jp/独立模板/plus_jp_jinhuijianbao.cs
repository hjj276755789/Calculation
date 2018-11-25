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
    public class plus_jp_jinghuijianbao : plus_jp_base
    {
        public ISlideCollection _plus_jp_jinghuijianbao_1(string str, int cjbh)
        {
            try
            {

                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);
                foreach (var item in param)
                {
                    var tp = new Presentation(str);
                    var temp = tp.Slides;
                    var page1 = temp[0];

                    IAutoShape text1 = (IAutoShape)page1.Shapes[0];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, string.Join(",", item.ytcs[0]));

                    DataTable dt_jpbasj = new DataTable();
                    dt_jpbasj.Columns.Add(Base_Config_Jzgj.业态);
                    dt_jpbasj.Columns.Add(Base_Config_Jzgj.项目名称);

                    dt_jpbasj.Columns.Add(Base_Config_Cjba.上上上周_备案套数);
                    dt_jpbasj.Columns.Add("上上上周实际销售套数");
                    dt_jpbasj.Columns.Add(Base_Config_Cjba.上上周_备案套数);
                    dt_jpbasj.Columns.Add("上上周实际销售套数");
                    dt_jpbasj.Columns.Add(Base_Config_Cjba.上周_备案套数);
                    dt_jpbasj.Columns.Add("上周实际销售套数");
                    dt_jpbasj.Columns.Add(Base_Config_Cjba.本周_备案套数);
                    dt_jpbasj.Columns.Add("本周实际销售套数");


                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        //获取竞品项目数据
                        dt_jpbasj = GET_JPXM_BX(dt_jpbasj, item.jpxmlb);
                        Office_Tables.SetJP_JINHUIJIANBAO_Table(page1, dt_jpbasj, 1, null, null);
                        t.AddClone(page1);
                    }


                    var page2 = temp[1];
                    IAutoShape text2 = (IAutoShape)page2.Shapes[0];
                    text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc, item.ytcs[0]);

                    DataTable dt_2 = new DataTable();
                    dt_2.Columns.Add(Base_Config_Jzgj.业态);
                    dt_2.Columns.Add(Base_Config_Jzgj.项目名称);
                    dt_2.Columns.Add("在售楼栋");
                    dt_2.Columns.Add("面积区间");
                    dt_2.Columns.Add(Base_Config_Jzgj.竞争格局_主力面积区间);
                    dt_2.Columns.Add(Base_Config_Cjba.本周_备案套数);
                    dt_2.Columns.Add(Base_Config_Cjba.本周_建面均价);
                    dt_2.Columns.Add("总价范围");
                    dt_2.Columns.Add("主力总价");
                    dt_2.Columns.Add(Base_Config_Rgsj.本周_营销动作);

                    //获取本案数据
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        //获取竞品项目数据
                        dt_2 = GET_JPXM_BX_RG(dt_2, item.jpxmlb);
                        Office_Tables.SetJP_JINHUIJIANBAO_1_Table(page2, dt_2, 1, null, null);
                        t.AddClone(page2);
                    }
                }

                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log(e.Message);
                return null;
            }
        }

        public DataTable GET_JPXM_BX(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {

                if (item.ytcs[0] == "别墅")
                {
                    if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {

                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                           
                            #endregion

                            dt.Rows.Add(GET_ROW_BA_SZ(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));

                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        #endregion

                        dt.Rows.Add(GET_ROW_BA_SZ(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    if (item.hxcs != null & item.hxcs.Length > 0)
                    {
                        for (int i = 0; i < item.hxcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            
                            #endregion

                            dt.Rows.Add(GET_ROW_BA_SZ(item.hxcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                        }
                    }
                    else if(item.xfytcs!=null&&item.xfytcs.Length>0){
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            //本周本案认购数据
                            #endregion

                            dt.Rows.Add(GET_ROW_BA_SZ(item.xfytcs[i], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                        }

                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains( m["yt"].ToString()));
                        var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        //本周本案认购数据
                        #endregion

                        dt.Rows.Add(GET_ROW_BA_SZ(string.Join(",",item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    //竞品业态
                    var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_ba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_ba_ssz = Cache_data_cjjl.ssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_ba_sssz = Cache_data_cjjl.sssz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
 
                    #endregion

                    dt.Rows.Add(GET_ROW_BA_SZ(string.Join(",", item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                }


            }


            return dt;
        }

        public DataTable GET_JPXM_BX_RG(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {

                if (item.ytcs[0] == "别墅")
                {
                    if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {

                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_rg_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]).FirstOrDefault();
                            //本周本案认购数据
                           
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_rg_bz, null, temp_ba_bz, null, item));

                        }
                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_rg_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString())).FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs.Join(), dr1, dt, temp_rg_bz, null, temp_ba_bz, null, item));
                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    if (item.hxcs != null & item.hxcs.Length > 0)
                    {
                        for (int i = 0; i < item.hxcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                            var temp_rg_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString()==item.hxcs[i]).FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.hxcs[i], dr1, dt, temp_rg_bz, null, temp_ba_bz, null, item));
                        }
                    }
                    else if (item.xfytcs != null && item.xfytcs.Length > 0)
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["xfyt"].ToString() == item.xfytcs[i]);
                            var temp_rg_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]).FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_rg_bz, null, temp_ba_bz, null, item));
                        }

                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString()));
                        var temp_rg_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString())).FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs.Join(), dr1, dt, temp_rg_bz, null, temp_ba_bz, null, item));
                    }
                
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    //竞品业态
                    var temp_ba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rg_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString())).FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs.Join(), dr1, dt, temp_rg_bz, null, temp_ba_bz, null, item));
                }


            }


            return dt;
        }

    }
}
