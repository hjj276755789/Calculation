using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Calculation.Base;
using MySql.Data.MySqlClient;

namespace Calculation.Dal
{
    /// <summary>
    /// 认购数据
    /// </summary>
    public class ZB_Data_RGSJ_DataProvider
    {

        public static int Insert(DataTable data, int nf, int zc, string zcmc)
        {
            StringBuilder sb = new StringBuilder(@"insert into calculation.xtgl_data_zb_rgsj
(qymc,qy,zt,xm,yt,zx,
xkts,xkxsts,zljmqj,zltnqj,xktnjj,xkjmjj,
bats,batnjj,jmtl,tntl,
rgts,rgtnjj,rgjmjj,rgtntl,rgjmtl,rgje,
cjtshb,tnjjhb,bhyy,bzkc,
bzld,bzdfl,yh,yxdz, hd,bm,xzjtyj,nf,zc,zcmc) values ");
            string sql = "";
            int index = 0;
            int count = 0;
            if (remove("xtgl_data_zb_rgsj", nf, zc) != -1)
            {
                foreach (DataRow item in data.Rows)
                {
                    if (index != 0 && index % 3000 == 0)
                    {
                        sql = sb.ToString();
                        count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                        sb = new StringBuilder(@"insert into calculation.xtgl_data_zb_rgsj 
(qymc,qy,zt,xm,yt,zx,
xkts,xkxsts,zljmqj,zltnqj,xktnjj,xkjmjj,
bats,batnjj,jmtl,tntl,
rgts,rgtnjj,rgjmjj,rgtntl,rgjmtl,rgje,
cjtshb,tnjjhb,bhyy,bzkc,
bzld,bzdfl,yh,yxdz, hd,bm,xzjtyj,nf,zc,zcmc) values ");
                    }
                    sb.Append(string.Format(@"('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}','{30}','{31}','{32}','{33}','{34}','{35}'),",
                         item["企业"], item["区域"], item["组团"], item["项目"], item["业态"], item["装修"],
                         item["新开套数"].ToString().dw_zs(), item["新开销售套数"].ToString().dw_zs(), item["主力建面区间"].ToString(), item["主力套内面积区间"].ToString(), item["新开套内均价"].ToString().dw_xs(),item["新开建面均价"].ToString().dw_zs(),
                         item["备案套数"].ToString().dw_xs(), item["备案均价"].ToString().dw_xs(), item["建面体量"].ToString().dw_xs(), item["套内体量"].ToString().dw_xs(), 
                         item["认购套数"].ToString().dw_zs(), item["认购套内均价"].ToString().dw_zs(), item["认购建面均价"].ToString().dw_zs(), item["认购套内体量"].ToString().dw_zs(), item["认购建面体量"].ToString().dw_zs(), item["认购金额（万）"].ToString().dw_zs(),
                         item["成交套数环比"].ToString(), item["套内均价环比"].ToString(), item["变化原因"].ToString(), item["本周库存"].ToString().dw_zs(),
                         item["本周来电"].ToString().dw_zs(), item["本周到访量"].ToString().dw_zs(), item["优惠"], item["营销动作"], item["活动"], item["报媒"], item["下周加推预计"], nf, zc, zcmc));
                    index++;
                }
                sql = sb.ToString();
                count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                return count;
            } 
            else return 0;
        }

        public static DataTable GET_JBZ(int dqz)
        {
            string sql = @"select * from calculation.xtgl_data_zb_rgsj where zc between (@dqz - 7) and @dqz";
            MySqlParameter[] p = { new MySqlParameter("dqz", dqz) };
            return MySqlDbhelper.GetDataSet(sql, p).Tables[0];
        }
       
        public static int remove(string bm, int nf, int zc)
        {
            try
            {
                string sql = "delete from " + bm + " where nf=@nf and zc=@zc";
                MySqlParameter[] p = { new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
                return MySqlDbhelper.ExecuteNonQuery(sql, p);
            }
            catch (Exception)
            {

                return -1;
            }

        }
    }
}
