using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Dal
{
    public class MBGL_DataProvider
    {
        public static void ADD_MB(string mbmc,string mblx,string xflx,string cjmc,string cjdz,string cjlm,string cjffm)
        {
            string sql = "insert into calculation.xtgl_bbmb values ('"+ mbmc + "','" + mblx+"','"+xflx+"')";
            MySqlDbhelper.ExecuteNonQuery(sql);
            string mbid = MySqlDbhelper.ExecuteScalar("select max(id) from calculation.xtgl_bbmb").ToString() ;
            string sql1 = "insert into calculation.xtgl_bbcjb ('" + cjmc + "','"+cjdz+"','" +cjlm +"','"+ cjffm+ "')";
            MySqlDbhelper.ExecuteNonQuery(sql1);
            string cjid = MySqlDbhelper.ExecuteScalar("select max(id) from calculation.xtgl_bbcjb").ToString();
            MySqlDbhelper.ExecuteNonQuery("insert into calculation.xtgl_bbmbcj values(" + mbid+","+cjid+",0)");

        }
    }
}
