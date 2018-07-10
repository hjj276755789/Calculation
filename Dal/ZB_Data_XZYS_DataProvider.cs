using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Dal
{
    /// <summary>
    /// 新增预售
    /// </summary>
    public class ZB_Data_XZYS_DataProvider
    {
        public static int Insert(DataTable dt)
        {
            StringBuilder sb = new StringBuilder("insert into xtgl_xzysb (qx1,qx2,ztmc,xmmc,wylx,tyyt,bh,fzsj,jzmj,fzzmj,yf,zc) values ");
            string sql = "";
            foreach (DataRow item in dt.Rows)
            {
                if (index != 0 && index % 10000 == 0)
                {
                    sql = sb.ToString();
                    count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                    sb = new StringBuilder("insert into xtgl_xzysb (qx1,qx2,ztmc,xmmc,wylx,tyyt,bh,fzsj,jzmj,fzzmj,yf,zc) values ");
                }
                sb.Append(string.Format(@"('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}'),", item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9], item[10], item[11]));
                index++;
            }
            sql = sb.ToString();
            return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
        }

        public static int Insert(DataTable dt,int nf,int zc)
        {
            StringBuilder sb = new StringBuilder("insert into xtgl_xzysb (qx1,qx2,ztmc,xmmc,wylx,tyyt,bh,fzsj,jzmj,fzzmj,yf,nf,zc) values ");
            string sql = "";
            foreach (DataRow item in dt.Rows)
            {
                if (index != 0 && index % 10000 == 0)
                {
                    sql = sb.ToString();
                    count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                    sb = new StringBuilder("insert into xtgl_xzysb (qx1,qx2,ztmc,xmmc,wylx,tyyt,bh,fzsj,jzmj,fzzmj,yf,nf,zc) values ");
                }
                sb.Append(string.Format(@"('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}'),", item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9], item[10],nf,zc));
                index++;
            }
            sql = sb.ToString();
            return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
        }

        public static int Insert(DataTable dt, int nf, int zc,string zcmc)
        {
            StringBuilder sb = new StringBuilder("insert into xtgl_data_zb_xzys (qx1,qx2,zt,xmmc,wylx,tyyt,zjbh,fzsj,jzmj,fzzmj,nf,zc,zcmc) values ");
            string sql = "";
            foreach (DataRow item in dt.Rows)
            {
                if (index != 0 && index % 10000 == 0)
                {
                    sql = sb.ToString();
                    count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                    sb = new StringBuilder("insert into xtgl_data_zb_xzys (qx1,qx2,zt,xmmc,wylx,tyyt,zjbh,fzsj,jzmj,fzzmj,nf,zc,zcmc) values ");
                }
                sb.Append(string.Format(@"('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}'),", item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9], nf, zc,zcmc));
                index++;
            }
            sql = sb.ToString();
            return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
        }
        public static int index = 0;
        public static int count = 0;

        public static DataTable GET_BY(DateTime first, DateTime end)
        {
            string sql = @"select * from calculation.xtgl_data_zb_xzys where unix_timestamp(fzsj)
between unix_timestamp('" + first.ToString("yyyy/MM/dd") + "') and unix_timestamp('" + end.ToString("yyyy/MM/dd") + "')";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }


        public static DataTable GET_ZB(DateTime first, DateTime end)
        {
            string sql = @"select * from calculation.xtgl_data_zb_xzys where unix_timestamp(fzsj)
between unix_timestamp('" + first.ToString("yyyy/MM/dd") + "') and unix_timestamp('" + end.ToString("yyyy/MM/dd") + "')";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }
        /// <summary>
        /// 获取进8周新增预售记录
        /// </summary>
        /// <returns></returns>
        public static DataTable GET_JBZ(int dqz)
        {
            string sql = "select * from calculation.xtgl_data_zb_xzys where zc between @qsz and @dqz";
            MySqlParameter[] p = { new MySqlParameter("qsz", dqz - 7),new MySqlParameter("dqz", dqz) };
            return MySqlDbhelper.GetDataSet(sql,p).Tables[0];
        }
    }
}
