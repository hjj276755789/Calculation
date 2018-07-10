using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Dal
{
    public class Jsjg_zb_DataProvider
    {
        public static int index = 0;
        private static int count = 0;

        /// <summary>
        /// 新开盘走势
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static int ADD_XKPZS(DataTable dt)
        {
            StringBuilder sb = new StringBuilder("insert into calculation.xtgl_jsjg_xkqzs (tjnf,tjzc,gyts,rgts,rgl) values ");
            string sql = "";
            foreach (DataRow item in dt.Rows)
            {
                if (index != 0 && index % 10000 == 0)
                {
                    sql = sb.ToString();
                    count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                    sb = new StringBuilder("insert into calculation.xtgl_jsjg_xkqzs (tjnf,tjzc,gyts,rgts,rgl) values ");
                }
                sb.Append(string.Format(@"('{0}','{1}','{2}','{3}','{4}'),", item[0], item[1], item[2], item[3], item[4]));
                index++;
            }
            sql = sb.ToString();
            return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
        }

        public static DataTable GET_XKPZS()
        {
            string sql = @"select CONCAT('第',tjzc,'周') '统计周次',gyts'供应套数',rgts '认购套数',cast(rgl as decimal(9,2)) '认购率' from calculation.xtgl_jsjg_xkqzs where tjzc % 4 = 1";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }

    }
}
