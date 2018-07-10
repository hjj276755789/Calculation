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
        /// 获取所有插件列表
        /// </summary>
        /// <returns></returns>
        public static DataTable GET_CJLB()
        {
            string sql = "select a.* ,case when b.px is not null then 1 else 0 end sfxz from calculation.xtgl_bbcjb a  left join calculation.xtgl_bbmbcj b on a.cjbh = b.cjbh order by cjbh";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }
        /// <summary>
        /// 设置报表插件
        /// </summary>
        /// <param name="cjbh"></param>
        /// <returns></returns>
        public static int SET_BBCJ(List<string> cjbh)
        {
            string sql = "delete from calculation.xtgl_bbmbcj where mbbh = 1 ";
            
            MySqlDbhelper.ExecuteNonQuery(sql);

            StringBuilder sql1 = new StringBuilder(@"insert into calculation.xtgl_bbmbcj (mbbh,cjbh,px) values ");

            for (int i = 0; i < cjbh.Count; i++)
            {
                sql1.Append(string.Format(@"('{0}','{1}','{2}'),", 1, cjbh[i], i));
            }
            string str = sql1.ToString();
            return MySqlDbhelper.ExecuteNonQuery(str.Substring(0, str.Length - 1));

        }

        /// <summary>
        /// 通过报表编号获取插件列表
        /// </summary>
        /// <returns></returns>
        public static DataTable GET_CJLB_BB(int bbbh)
        {
            string sql = @"select a.mbmc,b.px,c.cjbh,c.cjmc,c.cjdz,c.cjps,c.cjclass,c.cjmethod,c.sfsx 
                            from calculation.xtgl_bbmb a, calculation.xtgl_bbmbcj b, calculation.xtgl_bbcjb c 
                            where a.mbbh =b.mbbh and b.cjbh =c.cjbh and a.mbbh=@mbbh order by px
                                ";
            MySqlParameter[] p = { new MySqlParameter("mbbh", bbbh) };
            return MySqlDbhelper.GetDataSet(sql,p).Tables[0];
        }


      
    }
}
