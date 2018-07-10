using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Dal
{
    public class Jsjg_yb_DataProvider
    {
        public static int index =0;
        private static int count = 0;
        public static int ADD_SCGXFX(DataTable dt)
        {
            StringBuilder sb = new StringBuilder("insert into calculation.xtgl_jsjg_scgxfx (tjyf,cjmj,jmjj) values ");
            string sql = "";
            foreach (DataRow item in dt.Rows)
            {
                if (index != 0 && index % 10000 == 0)
                {
                    sql = sb.ToString();
                    count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                    sb = new StringBuilder("insert into  calculation.xtgl_jsjg_scgxfx (tjyf,cjmj,jmjj) values ");
                }
                sb.Append(string.Format(@"('{0}','{1}','{2}'),", item[0], item[1], item[2]));
                index++;
            }
            sql = sb.ToString();
            return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
        }
        public static int ADD_SCGXFX_PSB(DataTable dt)
        {
            StringBuilder sb = new StringBuilder("insert into calculation.xtgl_jsjg_scgxfx_psb (tjyf,gymj,cjmj,psb) values ");
            string sql = "";
            foreach (DataRow item in dt.Rows)
            {
                if (index != 0 && index % 10000 == 0)
                {
                    sql = sb.ToString();
                    count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                    sb = new StringBuilder("insert into  calculation.xtgl_jsjg_scgxfx (tjyf,gymj,cjmj,psb) values ");
                }
                sb.Append(string.Format(@"('{0}','{1}','{2}','{3}'),", item[0], item[1], item[2], item[3]));
                index++;
            }
            sql = sb.ToString();
            return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
        }

        public static int ADD_SCJGFX(DataTable dt)
        {
            StringBuilder sb = new StringBuilder("insert into calculation.xtgl_jsjg_scjgfx (nf,1y,2y,3y,4y,5y,6y,7y,8y,9y,10y,11y,12y) values ");
            string sql = "";
            foreach (DataRow item in dt.Rows)
            {
                if (index != 0 && index % 10000 == 0)
                {
                    sql = sb.ToString();
                    count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                    sb = new StringBuilder("insert into  calculation.xtgl_jsjg_scjgfx (nf,1y,2y,3y,4y,5y,6y,7y,8y,9y,10y,11y,12y) values ");
                }
                sb.Append(string.Format(@"('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}'),", item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9], item[10], item[11], item[12]));
                index++;
            }
            sql = sb.ToString();
            return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
        }




        /// <summary>
        /// 计算结果_市场供需分析
        /// </summary>
        /// <returns></returns>
        public static DataTable GET_SCGXFX()
        {
            string sql = "select tjyf '列1', cast(cjmj as decimal(9,2)) '成交面积（单位：万方）',floor(jmjj) '建面均价（单位：元/㎡）' from calculation.xtgl_jsjg_scgxfx";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }
        /// <summary>
        /// 计算结果_市场供需分析_批售比
        /// </summary>
        /// <returns></returns>
        public static DataTable GET_SCGXFX_PSB()
        {
            string sql = "select tjyf '列1', cast(gymj as decimal(9,2)) '供应面积(万平米)',floor(cjmj) '成交面积（万平米）', cast(psb as decimal(9,2)) '批售比' from calculation.xtgl_jsjg_scgxfx_psb";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }

        public static DataTable GET_SCJGFX()
        {
            string sql = @"select nf '列1',floor(1y) '1月',floor(2y) '2月',floor(3y) '3月',floor(4y) '4月',floor(5y) '5月',floor(6y) '6月',floor(7y) '7月',floor(8y) '8月',floor(9y) '9月',floor(10y) '10月',floor(11y) '11月',floor(12y) '12月'
 from calculation.xtgl_jsjg_scjgfx";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }
    }
}
