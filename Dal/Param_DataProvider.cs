using Calculation.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Calculation.Base;
using System.Data;

namespace Calculation.Dal
{
    public class Param_DataProvider
    {
        public static List<ParamModels> GET_MBCJCSLB(int mbid)
        {
            string sql = @"select d.csid,d.cjid,c.cjmc,d.csms,d.cslx,d.sfbl
                            from calculation.xtgl_bbmb a, calculation.xtgl_bbmbcj b, calculation.xtgl_bbcjb c, calculation.xtgl_cj_cjcs d
                            where a.mbid = b.mbid and b.cjbh = c.cjbh and b.cjbh = d.cjid and a.mbid =@mbid order by d.px
                ";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid) };
            return Models.Modelhelper.类列表赋值<ParamModels>(new ParamModels(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }

        public static List<ParamValueModel> GET_MBCJCSNR(int mbid, int nf, int zc)
        {
            string sql = "select a.csid,b.rwid,a.cjid,b.csnr from  calculation.xtgl_cj_cjcs a left join calculation.xtgl_cj_rwcs b on  a.csid = b.csid left join calculation.xtgl_bbrw c on b.rwid =c.rwid where c.mbid=@mbid and c.nf=@nf and c.zc =@zc ";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid), new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
            return Models.Modelhelper.类列表赋值<ParamValueModel>(new ParamValueModel(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }

        /// <summary>
        /// 模板插件参数内容
        /// </summary>
        /// <param name="mbid"></param>
        /// <param name="nf"></param>
        /// <param name="zc"></param>
        /// <returns></returns>
        public static List<ParamValueModel> GET_MBCJCSNR(int rwid)
        {
            string sql = @"select a.csid,b.rwid,a.cjid,b.csnr,a.csms from  
calculation.xtgl_cj_cjcs a
left join calculation.xtgl_cj_rwcs b on a.csid = b.csid
left join calculation.xtgl_bbrw c on b.rwid = c.rwid
where c.rwid = @rwid ";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid) };
            return Models.Modelhelper.类列表赋值<ParamValueModel>(new ParamValueModel(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }

        public static List<ParamValueModel> GET_RWCSNR(int rwid, int csid)
        {
            string sql = @"select  b.id rwcsid,a.csid,b.rwid,a.cjid,b.csnr,a.csms from  
calculation.xtgl_cj_cjcs a
left join calculation.xtgl_cj_rwcs b on a.csid = b.csid
left join calculation.xtgl_bbrw c on b.rwid = c.rwid
where c.rwid = @rwid and a.csid =@csid";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("csid", csid) };
            return Models.Modelhelper.类列表赋值<ParamValueModel>(new ParamValueModel(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }

        public static int SET_RWCJCS(int rwid, int csid, string csnr)
        {
            string sql = "insert into calculation.xtgl_cj_rwcs(rwid,csid,csnr)values(@rwid,@csid,@csnr)";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("csid", csid), new MySqlParameter("csnr", csnr) };
            if (MySqlDbhelper.ExecuteNonQuery(sql, p) > 0)
                return MySqlDbhelper.ExecuteScalar("select max(id) from calculation.xtgl_cj_rwcs").ints();
            else return -1;
        }
        public static int RESET_RWCJCS(int rwid, int csid, string csnr)
        {
            DEL_RWCJCS(rwid, csid);
            return SET_RWCJCS(rwid, csid, csnr);
        }

        /// <summary>
        /// 任务定稿文件
        /// </summary>
        /// <param name="rwid"></param>
        /// <param name="wjdz"></param>
        /// <returns></returns>
        public static bool SET_RWDGWJ(int rwid, string wjdz)
        {
            string sql = "update calculation.xtgl_bbrw set xzdz2 =@xzdz2 where rwid=@rwid";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("xzdz2", wjdz) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
        public static bool DEL_RWCJCS(int rwid, int csid)
        {
            string sql = "delete from xtgl_cj_rwcs where rwid=@rwid and csid =@csid";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("csid", csid) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
        public static bool DEL_RWCJCS(int rwcsid)
        {
            string sql = "delete from xtgl_cj_rwcs where id=@id";
            MySqlParameter[] p = { new MySqlParameter("id", rwcsid) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }

        #region 竞品
        public static IPageList<Data_Cjba_Default> GET_JP_CJBAXX(int nf, int zc, int pagesize, int pagenow)
        {
            string sql = "select * from calculation.xtgl_data_zb_cjba where nf=@nf and zc=@zc";
            MySqlParameter[] p = { new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
            return MySqlDbhelper.GetPagedList<Data_Cjba_Default>(sql, p, pagesize, pagenow);
        }
        public static IPageList<Data_Cjba_Default> GET_JP_CJBAXX(int nf, int zc, JP_ParamValueModel param, int pagesize, int pagenow)
        {

            string sql = "select * from calculation.xtgl_data_zb_cjba where nf=@nf and zc=@zc";
            string tempsql = "";
            if (param.qy != null)
            {
                tempsql += " and qy in ('" + string.Join("','", param.qy) + "')";
            }
            if (param.zt != null)
            {
                tempsql += " and zt in ('" + string.Join("','", param.zt) + "')";
            }
            if (param.lpmc != null)
            {
                tempsql += " and lpmc in ('" + string.Join("','", param.lpmc) + "')";
            }
            if (param.yt != null)
            {
                tempsql += " and yt in ('" + string.Join("','", param.yt) + "')";
            }
            if (param.xfyt != null)
            {
                tempsql += " and xfyt in ('" + string.Join("','", param.xfyt) + "')";
            }
            if (param.hx != null)
            {
                tempsql += " and  hx in ('" + string.Join("','", param.hx) + "')";
            }
            sql += tempsql;
            MySqlParameter[] p = { new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
            return MySqlDbhelper.GetPagedList<Data_Cjba_Default>(sql, p, pagesize, pagenow);
        }


        public static bool ADD_JP_BA(int rwid, string bamc)
        {
            string sql = @"INSERT INTO calculation.xtgl_param_jpba
                        ( rwid, bamc)
                        values (@rwid,@bamc)";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("bamc", bamc) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }

        public static List<JP_BA> GET_JP_BA(int rwid)
        {
            string sql = "select * from calculation.xtgl_param_jpba where rwid=@rwid";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid) };
            return Models.Modelhelper.类列表赋值<JP_BA>(new JP_BA(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        public static JP_BA GET_JP_BA_XQ(int id)
        {
            string sql = "select * from calculation.xtgl_param_jpba where id=@id";
            MySqlParameter[] p = { new MySqlParameter("id", id) };
            return Models.Modelhelper.类对象赋值<JP_BA>(new JP_BA(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        public static bool DEL_JP_BA(int id)
        {
            string sql = "delete from calculation.xtgl_param_jpba where id=@id";
            MySqlParameter[] p = { new MySqlParameter("id", id) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }

        public static bool SAVE_JP_JPXMCS(int id, JP_ParamValueModel p)
        {
            string sql = @"update calculation.xtgl_param_jpgj set ztcs=@ztcs,qycs=@qycs,lpcs=@lpcs,ytcs=@ytcs,xfytcs=@xfytcs,hxcs=@hxcs where id=@id";
            MySqlParameter[] par = { new MySqlParameter("id", id),
                new MySqlParameter("ztcs",p.zt==null||p.zt.Count()==0?"": string.Join("," ,p.zt)),
                new MySqlParameter("qycs",p.qy==null||p.qy.Count()==0?"": string.Join("," ,p.qy)),
                new MySqlParameter("lpcs",p.lpmc==null||p.lpmc.Count()==0?"": string.Join("," ,p.lpmc)),
                new MySqlParameter("ytcs",p.yt==null||p.yt.Count()==0?"": string.Join("," ,p.yt)),
                new MySqlParameter("xfytcs",p.xfyt==null||p.xfyt.Count()==0?"": string.Join("," ,p.xfyt)),
                new MySqlParameter("hxcs",p.hx==null||p.hx.Count()==0?"": string.Join("," ,p.hx)),};
            return MySqlDbhelper.ExecuteNonQuery(sql, par) > 0;
        }
        public static bool SAVE_JP_BAXMCS(int id, JP_ParamValueModel p)
        {
            string sql = @"update calculation.xtgl_param_jpba set ztcs=@ztcs,qycs=@qycs,lpcs=@lpcs,ytcs=@ytcs,xfytcs=@xfytcs,hxcs=@hxcs where id=@id";
            MySqlParameter[] par = { new MySqlParameter("id", id),
                new MySqlParameter("ztcs",p.zt==null||p.zt.Count()==0?"": string.Join("," ,p.zt)),
                new MySqlParameter("qycs",p.qy==null||p.qy.Count()==0?"": string.Join("," ,p.qy)),
                new MySqlParameter("lpcs",p.lpmc==null||p.lpmc.Count()==0?"": string.Join("," ,p.lpmc)),
                new MySqlParameter("ytcs",p.yt==null||p.yt.Count()==0?"": string.Join("," ,p.yt)),
                new MySqlParameter("xfytcs",p.xfyt==null||p.xfyt.Count()==0?"": string.Join("," ,p.xfyt)),
                new MySqlParameter("hxcs",p.hx==null||p.hx.Count()==0?"": string.Join("," ,p.hx)),};
            return MySqlDbhelper.ExecuteNonQuery(sql, par) > 0;
        }

        public static List<JP_JPXM> GET_JP_JPXM(int baid)
        {
            string sql = "select t1.*,t2.jzgjmc,t2.px from calculation.xtgl_param_jpgj t1 , calculation. dmb_jzgj t2   where t1.jzgjid=t2.id  and t1.baid = @baid order by px";
            MySqlParameter[] p = { new MySqlParameter("baid", baid) };
            return Models.Modelhelper.类列表赋值<JP_JPXM>(new JP_JPXM(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        public static JP_JPXM GET_JP_JPXM_XQ(int id)
        {
            string sql = "select * from  calculation.xtgl_param_jpgj t1  where id=@id";
            MySqlParameter[] p = { new MySqlParameter("id", id) };
            return Models.Modelhelper.类对象赋值<JP_JPXM>(new JP_JPXM(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        public static bool add_jp_jpxm(int baid, int jzgjid)
        {
            string sql = "insert into  calculation.xtgl_param_jpgj (baid,jzgjid) values(@baid,@jzgjid)";
            MySqlParameter[] p = { new MySqlParameter("baid", baid), new MySqlParameter("jzgjid", jzgjid) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }

        public static bool del_jp_jpxm(int id)
        {
            string sql = "delete from calculation.xtgl_param_jpgj where id =@id";
            MySqlParameter[] p = { new MySqlParameter("id", id) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }

        public static bool jcszsz(int rwid, int mbid, int nf, int zc)
        {
            DataTable batable = GET_JP_BAXX_REAL(mbid, nf, zc - 1);
            DataTable jptable = GET_JP_JPXMXX_REAL(mbid, nf, zc - 1);
            List<JP_BA_INFO> list = new List<JP_BA_INFO>();

            #region 清空本周参数
            del_jp_jpxx(rwid);
            del_jp_ba(rwid);
            #endregion




            #region 查询上周参数


            foreach (DataRow item in batable.Rows)
            {
                JP_BA_INFO jp = new JP_BA_INFO();
                jp.id = item["id"].ints();
                jp.bamc = item["bamc"].ToString();
                jp.rwid = item["rwid"].ints();
                jp.qycs = item["qycs"] == null || string.IsNullOrEmpty(item["qycs"].ToString()) ? null : item["qycs"].ToString().Split(',');
                jp.ztcs = item["ztcs"] == null || string.IsNullOrEmpty(item["ztcs"].ToString()) ? null : item["ztcs"].ToString().Split(',');
                jp.lpcs = item["lpcs"] == null || string.IsNullOrEmpty(item["lpcs"].ToString()) ? null : item["lpcs"].ToString().Split(',');
                jp.ytcs = item["ytcs"] == null || string.IsNullOrEmpty(item["ytcs"].ToString()) ? null : item["ytcs"].ToString().Split(',');
                jp.xfytcs = item["xfytcs"] == null || string.IsNullOrEmpty(item["xfytcs"].ToString()) ? null : item["xfytcs"].ToString().Split(',');
                jp.hxcs = item["hxcs"] == null || string.IsNullOrEmpty(item["hxcs"].ToString()) ? null : item["hxcs"].ToString().Split(',');
                jp.jpxmlb = new List<JP_JPXM_INFO>();
                var xmlist = jptable.AsEnumerable().Where(m => m["baid"].ints() == item["id"].ints());
                foreach (var xm in xmlist)
                {
                    JP_JPXM_INFO jpxm = new JP_JPXM_INFO();
                    jpxm.id = xm["id"].ints();
                    jpxm.baid = xm["baid"].ints();
                    jpxm.jzgjid = xm["jzgjid"].ints();
                    jpxm.qycs = xm["qycs"] == null || string.IsNullOrEmpty(xm["qycs"].ToString()) ? null : xm["qycs"].ToString().Split(',');
                    jpxm.ztcs = xm["ztcs"] == null || string.IsNullOrEmpty(xm["ztcs"].ToString()) ? null : xm["ztcs"].ToString().Split(',');
                    jpxm.lpcs = xm["lpcs"] == null || string.IsNullOrEmpty(xm["lpcs"].ToString()) ? null : xm["lpcs"].ToString().Split(',');
                    jpxm.ytcs = xm["ytcs"] == null || string.IsNullOrEmpty(xm["ytcs"].ToString()) ? null : xm["ytcs"].ToString().Split(',');
                    jpxm.xfytcs = xm["xfytcs"] == null || string.IsNullOrEmpty(item["xfytcs"].ToString()) ? null : xm["xfytcs"].ToString().Split(',');
                    jpxm.hxcs = xm["hxcs"] == null || string.IsNullOrEmpty(xm["hxcs"].ToString()) ? null : xm["hxcs"].ToString().Split(',');
                    jp.jpxmlb.Add(jpxm);
                }
                list.Add(jp);
            }
            #endregion
            foreach (var item in list)
            {
                string sql = "insert into calculation.xtgl_param_jpba (rwid,bamc,qycs,ztcs,lpcs,ytcs,xfytcs,hxcs) values(@rwid,@bamc,@qycs,@ztcs,@lpcs,@ytcs,@xfytcs,@hxcs)";
                MySqlParameter[] p = { new MySqlParameter("rwid",rwid),
                                        new MySqlParameter("bamc",item.bamc),
                                        new MySqlParameter("qycs",item.qycs==null?"":string.Join(",", item.qycs)),
                                        new MySqlParameter("ztcs",item.ztcs==null?"":string.Join(",", item.ztcs)),
                                        new MySqlParameter("lpcs",item.lpcs==null?"":string.Join(",", item.lpcs)),
                                        new MySqlParameter("ytcs",item.ytcs==null?"":string.Join(",", item.ytcs)),
                                        new MySqlParameter("xfytcs",item.xfytcs==null?"":string.Join(",", item.xfytcs)),
                                        new MySqlParameter("hxcs",item.hxcs==null?"":string.Join(",", item.hxcs)),
                };
                MySqlDbhelper.ExecuteNonQuery(sql, p);
                if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                {
                    StringBuilder sqlbuilder = new StringBuilder("insert into calculation.xtgl_param_jpgj(baid,jzgjid,qycs,ztcs,lpcs,ytcs,xfytcs,hxcs) values ");
                    foreach (var item_jp in item.jpxmlb)
                    {

                        sqlbuilder.Append(string.Format(@"((select max(id) from calculation.xtgl_param_jpba),'{0}','{1}','{2}','{3}','{4}','{5}','{6}'),",
                              item_jp.jzgjid,
                              item_jp.qycs == null ? "" : string.Join(",", item_jp.qycs),
                              item_jp.ztcs == null ? "" : string.Join(",", item_jp.ztcs),
                              item_jp.lpcs == null ? "" : string.Join(",", item_jp.lpcs),
                              item_jp.ytcs == null ? "" : string.Join(",", item_jp.ytcs),
                              item_jp.xfytcs == null ? "" : string.Join(",", item_jp.xfytcs),
                              item_jp.hxcs == null ? "" : string.Join(",", item_jp.hxcs)));
                    }
                    string sql1 = sqlbuilder.ToString();
                    MySqlDbhelper.ExecuteNonQuery(sql1.Substring(0, sql1.Length - 1));
                }
            }
            return true;
        }

        public static bool del_jp_ba(int rwid)
        {
            string sql = " delete from calculation.xtgl_param_jpba where rwid = @rwid";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p)>0;

        }
        public static bool del_jp_jpxx(int rwid)
        {
            string sql = " delete from calculation. xtgl_param_jpgj where baid in(select id from calculation.xtgl_param_jpba where rwid=@rwid)";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;

        }

        public static DataTable GET_JP_BAXX_REAL(int mbid, int nf, int zc)
        {
            string sql = @"select t2.* from calculation.xtgl_bbrw  t1 ,calculation.xtgl_param_jpba t2 
                where t1.rwid=t2.rwid 
                and t1.mbid=@mbid and nf=@nf and zc=@zc";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid), new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
            return MySqlDbhelper.GetDataSet(sql, p).Tables[0];

        }
        public static DataTable GET_JP_JPXMXX_REAL(int mbid, int nf, int zc)
        {
            string sql = @"select t3.* from calculation.xtgl_bbrw  t1 ,calculation.xtgl_param_jpba t2 ,calculation.xtgl_param_jpgj t3 
                    where t1.rwid=t2.rwid and t2.id =t3.baid
                    and t1.mbid=@mbid and nf=@nf and zc=@zc
                ";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid), new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
            return MySqlDbhelper.GetDataSet(sql, p).Tables[0];

        }

        #region 竞品-导出
        public static DataTable GET_JP_BAXX(int mbid, int nf, int zc)
        {
            string sql = @"select t2.*,t3.cjbh cjid from calculation.xtgl_bbrw  t1 ,calculation.xtgl_param_jpba t2 ,calculation.xtgl_bbmbcj t3
                where t1.rwid=t2.rwid and t1.mbid =t3.mbid 
                and t1.mbid=@mbid and nf=@nf and zc=@zc
                ";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid), new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
            return MySqlDbhelper.GetDataSet(sql, p).Tables[0];

        }
        public static DataTable GET_JP_JPXMXX(int mbid, int nf, int zc)
        {
            string sql = @"select t3.* ,t4.jzgjmc,t4.px from calculation.xtgl_bbrw  t1 ,calculation.xtgl_param_jpba t2 ,calculation.xtgl_param_jpgj t3 ,calculation.dmb_jzgj t4
                    where t1.rwid=t2.rwid and t2.id =t3.baid and t3.jzgjid =t4.id order by t3.baid,t4.px
                    and t1.mbid=@mbid and nf=@nf and zc=@zc
                ";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid), new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
            return MySqlDbhelper.GetDataSet(sql, p).Tables[0];

        }


        #endregion
        #endregion
    }
}
