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
    /// 土地成交记录
    /// </summary>
    public class ZB_Data_TDCJ_DataProvider
    {
        public static int Insert(DataTable dt)
        {
            StringBuilder sb = new StringBuilder(@"insert into xtgl_tdcjjl (xzq, syq, sjq, bk, dkmc, ggbh, dkbh, ggr, jyr,ggzj,
                                                                            ggdj, qplmj, bzj,cjfs, yjl, cjzj, cjdj, cjlmj, jdzrr, jdr, 
                                                                            sjkzr, jjlc, cpqys, cyjpqy, dkwz, yt, xz, szb, dksz, xg,
                                                                            jdzt, zyd_m, zyd_wm, jsyd, rjl, kjtl_wf, kjtl_f, sytl, zztl, xm,
                                                                            sfhz, kpsj, bz, sfdd, cjzt, zdjzmd, ldl, jmxzyq, fkjz, xftj, 
                                                                            sfxf, zc, gdf, gdsfcg
                                                                            ) values ");
            string sql = "";
           
            foreach (DataRow item in dt.Rows)
            {
                if (index != 0 && index % 10000 == 0)
                {
                    sql = sb.ToString();
                    count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                    sb = new StringBuilder(@"insert into xtgl_tdcjjl (xzq, syq, sjq, bk, dkmc, ggbh, dkbh, ggr, jyr,ggzj,
                                                                            ggdj, qplmj, bzj,cjfs, yjl, cjzj, cjdj, cjlmj, jdzrr, jdr, 
                                                                            sjkzr, jjlc, cpqys, cyjpqy, dkwz, yt, xz, szb, dksz, xg,
                                                                            jdzt, zyd_m, zyd_wm, jsyd, rjl, kjtl_wf, kjtl_f, sytl, zztl, xm,
                                                                            sfhz, kpsj, bz, sfdd, cjzt, zdjzmd, ldl, jmxzyq, fkjz, xftj, 
                                                                            sfxf, zc, gdf, gdsfcg
                                                                            ) values ");
                }
                sb.Append(string.Format(@"('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}'
                                                ,'{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}'
                                                ,'{21}','{22}','{2}','{24}','{25}','{26}','{27}','{28}','{29}','{30}'
                                                ,'{31}','{32}','{33}','{34}','{35}','{36}','{37}','{38}','{39}','{40}'
                                                ,'{41}','{42}','{43}','{44}','{45}','{46}','{47}','{48}','{49}','{50}'
                                                ,'{51}','{52}','{53}'),",
                                                item[0],
                                                item[1],
                                                item[2],
                                                item[3],
                                                item[4], 
                                                item[5], 
                                                item[6], 
                                                item[7],
                                                item[8], 
                                                item[9],
                                                item[10],
                                                item[11],
                                                item[12],
                                                item[13],
                                                item[14],
                                                item[15],
                                                item[16],
                                                item[17],
                                                item[18],
                                                item[19],
                                                item[20],
                                                item[21],
                                                item[22],
                                                item[23],
                                                item[24],
                                                item[25],
                                                item[26],
                                                item[27],
                                                item[28],
                                                item[29],
                                                item[30],
                                                item[31],
                                                item[32],
                                                item[33],
                                                item[34],
                                                item[35],
                                                item[36],
                                                item[37],
                                                item[38],
                                                item[39],
                                                item[40],
                                                item[41],
                                                item[42],
                                                item[43],
                                                item[44],
                                                item[45],
                                                item[46],
                                                item[47],
                                                item[48],
                                                item[49],
                                                item[50],
                                                item[51],
                                                item[52],
                                                item[53]));
                index++;
            }
            sql = sb.ToString();
            return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
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
        public static int Insert(DataTable dt,int nf,int zc)
        {
            StringBuilder sb = new StringBuilder(@"insert into xtgl_data_zb_tdcj (xzq, syq, sjq, bk, dkmc, ggbh, dkbh, ggr, jyr,ggzj,
                                                                            ggdj, qplmj, bzj,cjfs, yjl, cjzj, cjdj, cjlmj, jdzrr, jdr, 
                                                                            sjkzr, jjlc, cpqys, cyjpqy, dkwz, yt, xz, szb, dksz, xg,
                                                                            jdzt, zyd_m, zyd_wm, jsyd, rjl, kjtl_wf, kjtl_f, sytl, zztl, xm,
                                                                            sfhz, kpsj, bz, sfdd, cjzt, zdjzmd, ldl, jmxzyq, fkjz, xftj, 
                                                                            sfxf, zichi, gdf, gdsfcg,nf,zc
                                                                            ) values ");
            string sql = "";
            if (remove("xtgl_data_zb_tdcj", nf, zc)!=-1) {
                foreach (DataRow item in dt.Rows)
                {
                    if (index != 0 && index % 10000 == 0)
                    {
                        sql = sb.ToString();
                        count += MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
                        sb = new StringBuilder(@"insert into xtgl_data_zb_tdcj (xzq, syq, sjq, bk, dkmc, ggbh, dkbh, ggr, jyr,ggzj,
                                                                            ggdj, qplmj, bzj,cjfs, yjl, cjzj, cjdj, cjlmj, jdzrr, jdr, 
                                                                            sjkzr, jjlc, cpqys, cyjpqy, dkwz, yt, xz, szb, dksz, xg,
                                                                            jdzt, zyd_m, zyd_wm, jsyd, rjl, kjtl_wf, kjtl_f, sytl, zztl, xm,
                                                                            sfhz, kpsj, bz, sfdd, cjzt, zdjzmd, ldl, jmxzyq, fkjz, xftj, 
                                                                            sfxf, zichi, gdf, gdsfcg,nf,zc
                                                                            ) values ");
                    }
                    sb.Append(string.Format(@"('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}'
                                                ,'{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}'
                                                ,'{21}','{22}','{2}','{24}','{25}','{26}','{27}','{28}','{29}','{30}'
                                                ,'{31}','{32}','{33}','{34}','{35}','{36}','{37}','{38}','{39}','{40}'
                                                ,'{41}','{42}','{43}','{44}','{45}','{46}','{47}','{48}','{49}','{50}'
                                                ,'{51}','{52}','{53}','{54}',{55}),",
                                                    item[0],
                                                    item[1],
                                                    item[2],
                                                    item[3],
                                                    item[4],
                                                    item[5],
                                                    item[6],
                                                    item[7],
                                                    item[8],
                                                    item[9],
                                                    item[10],
                                                    item[11],
                                                    item[12],
                                                    item[13],
                                                    item[14],
                                                    item[15],
                                                    item[16],
                                                    item[17],
                                                    item[18],
                                                    item[19],
                                                    item[20],
                                                    item[21],
                                                    item[22],
                                                    item[23],
                                                    item[24],
                                                    item[25],
                                                    item[26],
                                                    item[27],
                                                    item[28],
                                                    item[29],
                                                    item[30],
                                                    item[31],
                                                    item[32],
                                                    item[33],
                                                    item[34],
                                                    item[35],
                                                    item[36],
                                                    item[37],
                                                    item[38],
                                                    item[39],
                                                    item[40],
                                                    item[41],
                                                    item[42],
                                                    item[43],
                                                    item[44],
                                                    item[45],
                                                    item[46],
                                                    item[47],
                                                    item[48],
                                                    item[49],
                                                    item[50],
                                                    item[51],
                                                    item[52],
                                                    item[53],
                                                    nf.ToString(),
                                                    zc
                                                    ));
                    index++;
                }
                sql = sb.ToString();
                return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
            }
            else
            {
                return 0;
            }
        }

        public static int index = 0;
        public static int count = 0;

        public static DataTable GET_ZB(DateTime first, DateTime end)
        {
            string sql = @"select syq,bk,xz,zyd_m,kjtl_wf,cjzj,cjfs,nf,zc,zcmc from calculation.xtgl_data_zb_tdcj where  unix_timestamp(jyr)
between unix_timestamp('" + first.ToString("yyyy/MM/dd") + "') and unix_timestamp('" + end.ToString("yyyy/MM/dd") + "')";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }
        public static DataTable GET_ZB(int  nf, int zc)
        {
            string sql = @"select syq,bk,xz,zyd_m,kjtl_wf,cjzj,cjfs,nf,zc,zcmc from calculation.xtgl_data_zb_tdcj where nf=@nf and zc=@zc";
            return MySqlDbhelper.GetDataSet(sql).Tables[0];
        }
        public static DataTable GET_JBZ(int nf,int dqz)
        {
            string sql = @"select syq,bk,xz,zyd_m,kjtl_wf,cjzj,cjfs,nf,zc,zcmc from calculation.xtgl_data_zb_tdcj where nf=@nf and (zc between (@dqz - 7) and @dqz)";
            MySqlParameter[] p = { new MySqlParameter("nf", nf),new MySqlParameter("dqz", dqz) };
            return MySqlDbhelper.GetDataSet(sql, p).Tables[0];
        }
    }
}
