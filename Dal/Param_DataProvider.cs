using Calculation.Models.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Calculation.Base;

namespace Calculation.Dal
{
    public class Param_DataProvider
    {
        public static List<ParamModels> GET_MBCJCSLB(int mbbh)
        {
            string sql = @"select d.csid,d.cjid,c.cjmc,d.csms,d.cslx,d.sfbl
                            from calculation.xtgl_bbmb a, calculation.xtgl_bbmbcj b, calculation.xtgl_bbcjb c, calculation.xtgl_cj_cjcs d
                            where a.mbbh = b.mbbh and b.cjbh = c.cjbh and b.cjbh = d.cjid and a.mbbh = @mbbh order by d.px
                ";
            MySqlParameter[] p = { new MySqlParameter("mbbh", mbbh) };
            return Models.Modelhelper.类列表赋值< ParamModels >(new ParamModels(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }

        public static List<ParamValueModel> GET_MBCJCSNR(int mbid,int nf,int zc)
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


        public static int SET_RWCJCS(int rwid,int csid,string csnr)
        {
            string sql = "insert into calculation.xtgl_cj_rwcs(rwid,csid,csnr)values(@rwid,@csid,@csnr)";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid),new MySqlParameter("csid", csid),new MySqlParameter("csnr", csnr) };
            if (MySqlDbhelper.ExecuteNonQuery(sql, p) > 0)
                return MySqlDbhelper.ExecuteScalar("select max(id) from calculation.xtgl_cj_rwcs").ints();
            else return -1;
        }
        public static int RESET_RWCJCS(int rwid, int csid, string csnr)
        {
            DEL_RWCJCS(rwid, csid);
            return SET_RWCJCS(rwid, csid, csnr);
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
    }
}
