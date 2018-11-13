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
    public class plus_jp_base_jpmb1 : weak
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="yt"></param>
        /// <param name="dr1"></param>
        /// <param name="dt"></param>
        /// <param name="temp_ba_bz"></param>
        /// <param name="temp_ba_sz"></param>
        /// <param name="temp_ba_ssz"></param>
        /// <param name="temp_ba_sssz"></param>
        /// <param name="item"></param>
        /// <returns></returns>
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
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null ? temp_ba_bz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.本周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null && temp_ba_bz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_ba_bz[Base_Config_Rgsj.本周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_bz[Base_Config_Rgsj.本周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_bz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
                            }; break;
                        case Base_Config_Rgsj.上周_新开销售套数:
                        case Base_Config_Rgsj.上周_新开套数:
                        case Base_Config_Rgsj.上周_认购套数:
                        case Base_Config_Rgsj.上周_认购套内均价:
                        case Base_Config_Rgsj.上周_认购建面均价:
                        case Base_Config_Rgsj.上周_认购套内体量:
                        case Base_Config_Rgsj.上周_认购建面体量:
                        case Base_Config_Rgsj.上周_认购金额:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_sz != null ? temp_ba_sz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.上周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_sz != null && temp_ba_sz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_ba_sz[Base_Config_Rgsj.本周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_sz[Base_Config_Rgsj.本周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_sz[Base_Config_Rgsj.本周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
                            }; break;
                        case Base_Config_Rgsj.上上周_新开销售套数:
                        case Base_Config_Rgsj.上上周_新开套数:
                        case Base_Config_Rgsj.上上周_认购套数:
                        case Base_Config_Rgsj.上上周_认购套内均价:
                        case Base_Config_Rgsj.上上周_认购建面均价:
                        case Base_Config_Rgsj.上上周_认购套内体量:
                        case Base_Config_Rgsj.上上周_认购建面体量:
                        case Base_Config_Rgsj.上上周_认购金额:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_ssz != null ? temp_ba_ssz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.上上周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_ssz != null && temp_ba_ssz[Base_Config_Rgsj.上上周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_ba_ssz[Base_Config_Rgsj.上上周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_ssz[Base_Config_Rgsj.上上周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_ssz[Base_Config_Rgsj.上上周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
                            }; break;
                        case Base_Config_Rgsj.上上上周_新开销售套数:
                        case Base_Config_Rgsj.上上上周_新开套数:
                        case Base_Config_Rgsj.上上上周_认购套数:
                        case Base_Config_Rgsj.上上上周_认购套内均价:
                        case Base_Config_Rgsj.上上上周_认购建面均价:
                        case Base_Config_Rgsj.上上上周_认购套内体量:
                        case Base_Config_Rgsj.上上上周_认购建面体量:
                        case Base_Config_Rgsj.上上上周_认购金额:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_sssz != null ? temp_ba_sssz[dt.Columns[j].ColumnName._ConfigRgsjMc()] : 0;
                            }; break;
                        case Base_Config_Rgsj.上上上周_套均总价:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_sssz != null && temp_ba_sssz[Base_Config_Rgsj.上上上周_认购套数._ConfigRgsjMc()].doubls() != 0 ? (temp_ba_sssz[Base_Config_Rgsj.上上上周_认购套内均价._ConfigRgsjMc()].doubls() * temp_ba_sssz[Base_Config_Rgsj.上上上周_认购套内体量._ConfigRgsjMc()].doubls() / temp_ba_sssz[Base_Config_Rgsj.上上上周_认购套数._ConfigRgsjMc()].doubls()).je_wy() : 0;
                            }; break;
                        default:
                            {
                                dr1[dt.Columns[j].ColumnName] = temp_ba_bz != null ? temp_ba_bz[dt.Columns[j].ColumnName] : "-";
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
                
            }

            return dr1;
        }
        /// <summary>
        /// 近4周整体市场表现
        /// </summary>
        /// <param name="str"></param>
        /// <param name="ztmc"></param>
        /// <returns></returns>
        public DataTable JSZZTSCBX(string [] ztmc)
        {
        
            System.Data.DataTable zzsc = new System.Data.DataTable();
            zzsc.Columns.Add("时间");
            zzsc.Columns.Add("供应量（万方）");
            zzsc.Columns.Add("成交量（万方）");
            zzsc.Columns.Add("建面均价（元 /㎡）");

            var jbz_cjba = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                            where ztmc.Contains(a["zt"].ToString()) 
                            group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                            select new
                            {
                                zc = s.Key.zc,
                                zcmc = s.Key.zcmc,
                                cjje = s.Sum(a => a["cjje"].longs()),
                                jzmj = s.Sum(a => a["jzmj"].doubls()),
                            }).OrderBy(m => m.zc).ToList();
            var jbz_xzys = (from a in Cache_data_xzys.jbz.AsEnumerable()
                            where ztmc.Contains(a["zt"].ToString()) 
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
                         }).OrderBy(m=>m.zc).Skip(4).ToList();
            for (int i = 0; i < temp6.Count(); i++)
            {
                DataRow dr = zzsc.NewRow();
                dr[0] = temp6[i].zcmc;
                dr[1] = temp6[i].xzgyl.mj_wf();
                dr[2] = temp6[i].cjmj.mj_wf();
                dr[3] = temp6[i].jmjj.je_y();
                zzsc.Rows.Add(dr);
            }
            return zzsc;
        }
        /// <summary>
        /// 近4周区域整体市场表现
        /// </summary>
        /// <param name="str"></param>
        /// <param name="ztmc"></param>
        /// <returns></returns>
        public DataTable JSZQYSCBX(string[] qymc)
        {

            System.Data.DataTable zzsc = new System.Data.DataTable();
            zzsc.Columns.Add("时间");
            zzsc.Columns.Add("供应量（万方）");
            zzsc.Columns.Add("成交量（万方）");
            zzsc.Columns.Add("建面均价（元 /㎡）");

            var jbz_cjba = (from a in Cache_data_cjjl.jbz.AsEnumerable()
                            where qymc.Contains(a["qy"].ToString())
                            group a by new { zc = a["zc"], zcmc = a["zcmc"] } into s
                            select new
                            {
                                zc = s.Key.zc,
                                zcmc = s.Key.zcmc,
                                cjje = s.Sum(a => a["cjje"].longs()),
                                jzmj = s.Sum(a => a["jzmj"].doubls()),
                            }).OrderBy(m => m.zc).ToList();
            var jbz_xzys = (from a in Cache_data_xzys.jbz.AsEnumerable()
                            where qymc.Contains(a["qx1"].ToString())
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
                         }).OrderBy(m => m.zc).Skip(4).ToList();
            for (int i = 0; i < temp6.Count(); i++)
            {
                DataRow dr = zzsc.NewRow();
                dr[0] = temp6[i].zcmc;
                dr[1] = temp6[i].xzgyl.mj_wf();
                dr[2] = temp6[i].cjmj.mj_wf();
                dr[3] = temp6[i].jmjj.je_y();
                zzsc.Rows.Add(dr);
            }
            return zzsc;
        }
        public DataTable GET_JPXM_ZT_CJJE(DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {
                var temp_sssz = Cache_data_rgsj.sssz.AsEnumerable().Where(m => item.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
                var temp_ssz = Cache_data_rgsj.ssz.AsEnumerable().Where(m => item.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
                var temp_sz = Cache_data_rgsj.sz.AsEnumerable().Where(m => item.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
                var temp_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => item.kfs.Contains(m["qymc"])).Sum(m => m["rgje"].longs());
                DataRow dr = dt.NewRow();
                dr["kfs"] = string.Join(",", item.kfs);
                dr["hj"] = temp_sssz + temp_ssz + temp_sz + temp_bz;
                dr["sssz_cjje"] = temp_sssz;
                dr["ssz_cjje"] = temp_ssz;
                dr["sz_cjje"] = temp_sz;
                dr["bz_cjje"] = temp_bz;
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public DataTable GET_JPXM_XF_CJJE(DataTable dt, JP_JPXM_INFO jpxm)
        {
            string sql = "zc >=" + (Base_date.bz - 3) + " and zc<=" + Base_date.bz;
            var query = (from t in Cache_data_rgsj.jbz.Select(sql).AsEnumerable()
                         where jpxm.kfs.Contains(t["qymc"])
                         group t by new { xm = t["xm"], yt = t["yt"] } into m
                         select new
                         {
                             xm = m.Key.xm + "(" + m.Key.yt + ")",
                             hj = m.Sum(n => n["rgje"].longs()),
                             sssz = m.Where(a => a["zc"].ints() == (Base_date.bz - 3)).Sum(n => n["rgje"].longs()),
                             ssz = m.Where(a => a["zc"].ints() == (Base_date.bz - 2)).Sum(n => n["rgje"].longs()),
                             sz = m.Where(a => a["zc"].ints() == (Base_date.bz - 1)).Sum(n => n["rgje"].longs()),
                             bz = m.Where(a => a["zc"].ints() == Base_date.bz).Sum(n => n["rgje"].longs()),
                         }).ToList();
            foreach (var item in query)
            {
                DataRow dr = dt.NewRow();
                dr["kfs"] = item.xm;
                dr["hj"] = item.hj;
                dr["sssz"] = item.sssz;
                dr["ssz"] = item.ssz;
                dr["sz"] = item.sz;
                dr["bz"] = item.bz;
                dt.Rows.Add(dr);
            }

            return dt;
        }
    }
}
