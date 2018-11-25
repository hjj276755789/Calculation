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

    /// <summary>
    /// 启迪协信
    /// </summary>
    public class plus_jp_qidixiexin :plus_jp_base
    {
        public ISlideCollection _plus_jp_qidixiexin_1(string str, int cjbh)
        {
            try
            {

                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);
                foreach (var item in param)
                {

                     var query = from a in item.jpxmlb
                                group a by new { jzgjid = a.jzgjid } into m
                                select new
                                {
                                    jzgjid = m.Key.jzgjid,
                                };
                    List< List<JP_JPXM_INFO> >list = new List<List<JP_JPXM_INFO>>();
                    foreach (var jzgjid in query)
                    {
                        List<JP_JPXM_INFO> jpxm = item.jpxmlb.Where(m => m.jzgjid == jzgjid.jzgjid).ToList();
                        list.Add(jpxm);
                    }
                    foreach (var jpxmlb in list)
                    {
                        var tp = new Presentation(str);
                        var temp = tp.Slides;
                        var page1 = temp[0];

                        IAutoShape text1 = (IAutoShape)page1.Shapes[0];
                        text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc);

                        DataTable dt_jpbasj = new DataTable();
                        dt_jpbasj.Columns.Add(Base_Config_Jzgj.项目名称);
                        dt_jpbasj.Columns.Add(Base_Config_Jzgj.业态);


                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.上周_新开套数);
                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.上周_主力建面区间);
                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.上周_认购建面均价);

                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.本周_新开套数);
                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.本周_主力建面区间);
                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);

                        dt_jpbasj.Columns.Add("本周存量");
                        dt_jpbasj.Columns.Add("本月认购套数");
                        dt_jpbasj.Columns.Add("本月建面均价");
                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.本周_营销动作);
                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.本周_本周到访量);
                        dt_jpbasj.Columns.Add(Base_Config_Rgsj.本周_下周加推预计);
                        dt_jpbasj.Columns.Add("加推套数建面均价");

                        if (jpxmlb.Count > 0)
                        {
                            //获取竞品项目数据
                            dt_jpbasj = GET_JPXM_BX_RG(dt_jpbasj, jpxmlb);
                            Office_Tables.SetJP_QIDIXIEXIN_1_Table(page1, dt_jpbasj, 1, null, null);
                            t.AddClone(page1);
                        }
                        
                    }
                    var tp1 = new Presentation(str);
                    var temp1 = tp1.Slides;
                    var page2 = temp1[1];
                    IAutoShape text2 = (IAutoShape)page2.Shapes[0];
                    text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc);

                    DataTable dt_2 = new DataTable();
                    dt_2.Columns.Add(Base_Config_Jzgj.业态);
                    dt_2.Columns.Add("推出库存");
                    dt_2.Columns.Add("推出区划周期");
                    dt_2.Columns.Add("未推库存");
                    dt_2.Columns.Add("中期库存");
                    dt_2.Columns.Add("中期去化周期");

                    var jpyt=  from a in item.jpxmlb
                               group a by new { ytcs = a.ytcs.Join() } into m
                               select new
                               {
                                   yt = m.Key.ytcs,
                               };
                    foreach (var yt in jpyt)
                    {
                        DataRow dr = dt_2.NewRow();
                        dr[Base_Config_Jzgj.业态] = yt.yt;
                        dt_2.Rows.Add(dr);
                    }
                    Office_Tables.SetJP_QIDIXIEXIN_2_Table(page2, dt_2, 1, null, null);
                    t.AddClone(page2);

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
                    else if (item.xfytcs != null && item.xfytcs.Length > 0)
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
                        //本周本案认购数据
                        #endregion

                        dt.Rows.Add(GET_ROW_BA_SZ(string.Join(",", item.ytcs), dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
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
            var list = jpxm.OrderBy(m => m.lpcs.Join()).ToList();
            foreach (var item in list)
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
                    if (item.hxcs.IsNotNull())
                    {
                        for (int i = 0; i < item.hxcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_rg_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]).FirstOrDefault();
                            var temp_rg_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]).FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.hxcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, null, null, item));
                        }
                    }
                    else if (item.xfytcs.IsNotNull())
                    {
                        for (int i = 0; i < item.xfytcs.Length; i++)
                        {
                            DataRow dr1 = dt.NewRow();

                            #region 数据准备
                            //竞品业态
                            var temp_rg_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]).FirstOrDefault();
                            var temp_rg_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]).FirstOrDefault();
                            #endregion

                            dt.Rows.Add(GET_ROW(item.xfytcs[i], dr1, dt, temp_rg_bz, temp_rg_sz, null, null, item));
                        }

                    }
                    else
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                      
                        var temp_rg_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString())).FirstOrDefault();
                        var temp_rg_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString())).FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs.Join(), dr1, dt, temp_rg_bz, temp_rg_sz, null, null, item));
                    }

                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态

                    var temp_rg_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString())).FirstOrDefault();
                    var temp_rg_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && item.ytcs.Contains(m["yt"].ToString())).FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs.Join(), dr1, dt, temp_rg_bz, temp_rg_sz, null, null, item));
                }


            }


            return dt;
        }
    }
}
