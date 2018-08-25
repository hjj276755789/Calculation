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
    public class plus_jp_xianggangzhidi : plus_jp_base
    {
        public ISlideCollection _plus_jp_xianggangzhidi_1(string str, int cjbh)
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

                    #region 格局统计
                    var page = temp[1];
                    DataTable dt = new DataTable();
                    dt.Columns.Add(Base_Config_Jzgj.业态);
                    dt.Columns.Add(Base_Config_Jzgj.组团);
                    dt.Columns.Add(Base_Config_Jzgj.项目名称);

                    dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                    dt.Columns.Add(Base_Config_Rgsj.本周_认购套内均价);
                    dt.Columns.Add(Base_Config_Rgsj.本周_套均总价);

                    dt.Columns.Add(Base_Config_Rgsj.上周_认购套数);
                    dt.Columns.Add(Base_Config_Rgsj.上周_认购套内均价);
                    dt.Columns.Add(Base_Config_Rgsj.上周_套均总价);


                    dt.Columns.Add("ssz_rgts");//上上周_备案套数
                    dt.Columns.Add("ssz_rgtnjj");//上上周_套内均价
                    dt.Columns.Add("ssz_tjzj");//上上周_套内均价

                    dt.Columns.Add("sssz_rgts");//上上上周_备案套数
                    dt.Columns.Add("sssz_rgtnjj");//上上上周_套内均价
                    dt.Columns.Add("sssz_tjzj");//上上上周_套内均价

                    dt.Columns.Add(Base_Config_Rgsj.变化原因);
                    #endregion
                    if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    {
                        dt = GET_JPXM_BX(dt, item.jpxmlb);
                        Office_Tables.SetJP_XiangGangZhiDi_JPBX_Table(page, dt, 1, null, null);
                        t.AddClone(page);
                    }
                }
                return t;
            }
            catch (Exception e)
            {
                return null;
            }

        }

        /// <summary>
        /// 竞品表现
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
                    for (int i = 0; i < item.xfytcs.Length; i++)
                    {

                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_rgsj_ssz = (Cache_data_rgsj.jbz.Select("zc=" + (Base_date.bz - 2)).CopyToDataTable()).AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        var temp_rgsj_sssz = (Cache_data_rgsj.jbz.Select("zc=" + (Base_date.bz - 3)).CopyToDataTable()).AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.xfytcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        var temp_ba_ssz = temp_rgsj_ssz.FirstOrDefault();
                        var temp_ba_sssz = temp_rgsj_sssz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));

                    }
                }
                else if (item.ytcs[0] == "商务")
                {
                    for (int i = 0; i < item.hxcs.Length; i++)
                    {
                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        var temp_rgsj_ssz = (Cache_data_rgsj.jbz.Select("zc=" + (Base_date.bz - 2)).CopyToDataTable()).AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        var temp_rgsj_sssz = (Cache_data_rgsj.jbz.Select("zc=" + (Base_date.bz - 3)).CopyToDataTable()).AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.hxcs[i]);
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                        var temp_ba_ssz = temp_rgsj_ssz.FirstOrDefault();
                        var temp_ba_sssz = temp_rgsj_sssz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                    }
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_rgsj_ssz = Cache_data_rgsj.jbz.Select("zc=" + (Base_date.bz - 2));
                    var temp_rgsj_sssz = Cache_data_rgsj.jbz.Select("zc=" + (Base_date.bz - 3));
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    var temp_ba_sz = temp_rgsj_sz.FirstOrDefault();
                    var temp_ba_ssz = temp_rgsj_ssz!=null && temp_rgsj_ssz.Length > 0 ? temp_rgsj_ssz.CopyToDataTable().AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]).FirstOrDefault():null;
                    var temp_ba_sssz = temp_rgsj_sssz!=null&& temp_rgsj_sssz.Length>0 ? temp_rgsj_sssz.CopyToDataTable().AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]).FirstOrDefault() : null ;
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_ba_sz, temp_ba_ssz, temp_ba_sssz, item));
                }


            }


            return dt;
        }


        public DataRow GET_ROW(string yt, DataRow dr1, System.Data.DataTable dt,
                              DataRow temp_ba_bz,
                              DataRow temp_ba_sz,
                              DataRow temp_ba_ssz,
                              DataRow temp_ba_sssz,
                              JP_JPXM_INFO item)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {


                if (Base_Config_Rgsj._认购数据.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Rgsj.本周_新开销售套数:
                        case Base_Config_Rgsj.本周_新开套数:
                        case Base_Config_Rgsj.本周_认购套数:
                        case Base_Config_Rgsj.本周_认购套内均价:
                        case Base_Config_Rgsj.本周_认购建面均价:
                        case Base_Config_Rgsj.本周_认购套内体量:
                        case Base_Config_Rgsj.本周_认购建面体量:
                        case Base_Config_Rgsj.本周_认购金额:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz!= null? temp_ba_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()]:0;
                            }; break;
                        case Base_Config_Rgsj.本周_套均总价: {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null&&temp_ba_bz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls()!=0? (temp_ba_bz[Base_Config_Rgsj.本周_认购套内均价._ConfigRgsjMc()].doubls()* temp_ba_bz[Base_Config_Rgsj.本周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_bz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls()).je_wy():0;
                            };break;
                        case Base_Config_Rgsj.上周_新开销售套数:
                        case Base_Config_Rgsj.上周_新开套数:
                        case Base_Config_Rgsj.上周_认购套数:
                        case Base_Config_Rgsj.上周_认购套内均价:
                        case Base_Config_Rgsj.上周_认购建面均价:
                        case Base_Config_Rgsj.上周_认购套内体量:
                        case Base_Config_Rgsj.上周_认购建面体量:
                        case Base_Config_Rgsj.上周_认购金额:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_sz != null ? temp_ba_sz[dt.Columns[j].ColumnName._ConfigRgsjMc()]:0;
                            }; break;
                        case Base_Config_Rgsj.上周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_sz != null&&temp_ba_sz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_ba_sz[Base_Config_Rgsj.本周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_sz[Base_Config_Rgsj.本周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_sz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
                            }; break;
                        default:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null? temp_ba_bz[dt.Columns[j].ColumnName]:"-";
                            }; break;
                    }
                }
                else if (Base_Config_Cjba._备案数据.Contains(dt.Columns[j].ColumnName))
                {
                    
                }
                else if (Base_Config_Jzgj._竞争格局参数名称.Contains(dt.Columns[j].ColumnName))
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case Base_Config_Jzgj.组团: { dr1[dt.Columns[j].ColumnName] = item.ztcs[0]; }; break;
                        case Base_Config_Jzgj.项目名称: { dr1[dt.Columns[j].ColumnName] = item.lpcs[0]; }; break;
                        case Base_Config_Jzgj.业态: { dr1[dt.Columns[j].ColumnName] = yt; }; break;
                        case Base_Config_Jzgj.竞争格局名称: { dr1[dt.Columns[j].ColumnName] = item.jzgjmc; }; break;
                        case Base_Config_Jzgj.竞争格局_主力面积区间: { dr1[dt.Columns[j].ColumnName] = item.zlmjqj; }; break;
                        default: { dr1[dt.Columns[j].ColumnName] = ""; }; break;
                    }

                }
                else
                {
                    switch (dt.Columns[j].ColumnName)
                    {
                        case "sssz_rgts": {dr1[dt.Columns[j].ColumnName] = temp_ba_sssz != null ? temp_ba_sssz["rgts"].ints():0; }; break;
                        case "sssz_rgtnjj": { dr1[dt.Columns[j].ColumnName] = temp_ba_sssz != null ? temp_ba_sssz["rgtnjj"].ints():0; }; break;
                        case "sssz_tjzj": { dr1[dt.Columns[j].ColumnName] = temp_ba_sssz != null && temp_ba_sssz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() != 0 ? temp_ba_sssz[Base_Config_Rgsj.本周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_sssz[Base_Config_Rgsj.本周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_sssz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() : 0; }; break;
                        case "ssz_rgts": { dr1[dt.Columns[j].ColumnName] = temp_ba_ssz!=null?temp_ba_ssz["rgts"].ints():0; }; break;
                        case "ssz_rgtnjj": { dr1[dt.Columns[j].ColumnName] = temp_ba_ssz != null ? temp_ba_ssz["rgtnjj"].ints():0; }; break;
                        case "ssz_tjzj": { dr1[dt.Columns[j].ColumnName] = temp_ba_ssz != null && temp_ba_ssz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() != 0 ? temp_ba_ssz[Base_Config_Rgsj.本周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_ssz[Base_Config_Rgsj.本周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_ssz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() : 0; }; break;
                    }
                }
            }

            return dr1;
        }
    }
}
