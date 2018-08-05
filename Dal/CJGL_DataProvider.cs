using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Dal
{

    public class CJGL_DataProvider
    {
        
        /// <summary>
        /// 通过报表编号获取插件列表
        /// </summary>
        /// <returns></returns>
        public static DataTable GET_CJLB_BB(int mbid)
        {
            string sql = @"select a.mbmc,b.px,c.cjbh,c.cjmc,c.cjdz,c.cjps,c.cjclass,c.cjmethod,c.sfsx 
                            from calculation.xtgl_bbmb a, calculation.xtgl_bbmbcj b, calculation.xtgl_bbcjb c 
                            where a.mbid =b.mbid and b.cjbh =c.cjbh and a.mbid=@mbid order by px
                                ";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid) };
            return MySqlDbhelper.GetDataSet(sql,p).Tables[0];
        }


      
    }
}
