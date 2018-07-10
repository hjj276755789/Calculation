using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using MySql.Data.MySqlClient;

namespace Calculation.Dal
{
    /// <summary>
    /// 成交备案
    /// </summary>
    public class ZB_Data_CJBA_DataProvider
    {

        public static int Insert(DataTable dt, int nf, int zc,string zcmc)
        {
            if (dt.Columns.Count != 12)
            {
                //数据结构出错
                return 0;
            }
            StringBuilder sb = new StringBuilder("insert into xtgl_data_zb_cjba (cjrq,qy,zt,kfs,lpmc,yt,xfyt,hx,jzmj,tnmj,cjje,ts,nf,zc,zcmc) values ");
            string sql = "";
            foreach (DataRow item in dt.Rows)
            {
                if (index != 0 && index % 10000 == 0)
                {
                    sql = sb.ToString();
                    count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                    sb = new StringBuilder("insert into xtgl_data_zb_cjba (cjrq,qy,zt,kfs,lpmc,yt,xfyt,hx,jzmj,tnmj,cjje,ts,nf,zc,zcmc) values ");
                }
                sb.Append(string.Format(@"('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}'),",
                     item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9], item[10], item[11], nf, zc,zcmc));
                index++;
            }
            sql = sb.ToString();
            return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
        }



        public static int index = 0;
        public static int count = 0;

        public static DataTable GET_YB(DateTime first,DateTime end)
        {
            string sql = @"select * from calculation.xtgl_data_zb_cjba where  unix_timestamp(cjrq)
between unix_timestamp('" + first.ToString("yyyy/MM/dd") + "') and unix_timestamp('" + end.ToString("yyyy/MM/dd") + "')";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }
        public static DataTable GET_ZB(DateTime first, DateTime end)
        {
            string sql = @"select * from calculation.xtgl_data_zb_cjba where  unix_timestamp(cjrq)
between unix_timestamp('" + first.ToString("yyyy/MM/dd") + "') and unix_timestamp('" + end.ToString("yyyy/MM/dd") + "')";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }
        public static DataTable GET_JBZ(int dqz)
        {
            string sql = @"select * from calculation.xtgl_data_zb_cjba where zc between (@dqz - 7) and @dqz";
            MySqlParameter[] p = {  new MySqlParameter("dqz", dqz) };
            return MySqlDbhelper.GetDataSet(sql,p).Tables[0];
        }
    }
}
