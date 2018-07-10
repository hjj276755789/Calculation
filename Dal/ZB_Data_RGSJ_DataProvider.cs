﻿using System;
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
        
        public static int Insert(DataTable data,int nf,int zc,string zcmc)
        {
            StringBuilder sb = new StringBuilder(@"insert into calculation.xtgl_data_zb_rgsj
(qymc,qy,zt,xm,yt,hx,
tjtnzlmj,xkts,xkxsts,xktnjj,
xkjmjj,rgts,rgtnjj,rgjmjj,rgtntl,
rgje,cjtshb,tnjjhb,bhyy,bzkc,
bzld,bzdfl,yh,yxdz,hd,
xzjtyj,nf,zc,zcmc) values ");
            string sql = "";
            int index = 0;
            int count = 0;
            foreach (DataRow item in data.Rows)
            {
                if (index != 0 && index % 10000 == 0)
                {
                    sql = sb.ToString();
                    count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                    sb = new StringBuilder("insert into calculation.xtgl_data_zb_rgsj (qymc,qy,zt,xm,yt,hx,tjtnzlmj,xkts,xkxsts,xktnjj,xkjmjj,rgts,rgtnjj,rgjmjj,rgtntl,rgje,cjtshb,tnjjhb,bhyy,bzkc,bzld,bzdfl,yh,yxdz,hd,xzjtyj,nf,zc,zcmc) values ");
                }
                sb.Append(string.Format(@"('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}'),",
                     item[0], item[1], item[2], item[3], item[4], item[5], 
                     item[6].ToString().dw_zs(), item[7].ToString().dw_zs(), item[8].ToString().dw_zs(), item[9].ToString().dw_xs(), item[10].ToString().dw_xs(),
                     item[11].ToString().dw_zs(), item[12].ToString().dw_xs(), item[13].ToString().dw_xs(), item[14].ToString().dw_xs(), item[15].ToString().dw_xs(), item[16], item[17], item[18],
                     item[19].ToString().dw_zs(), item[20].ToString().dw_zs(), item[21].ToString().dw_zs(), 
                     item[22], item[23], item[24], item[25], nf, zc,zcmc));
                index++;
            }
            sql = sb.ToString();
            count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
            return count;
        }

        public static DataTable GET_JBZ(int dqz)
        {
            string sql = @"select * from calculation.xtgl_data_zb_rgsj where zc between (@dqz - 7) and @dqz";
            MySqlParameter[] p = { new MySqlParameter("dqz", dqz) };
            return MySqlDbhelper.GetDataSet(sql, p).Tables[0];
        }
    }
}
