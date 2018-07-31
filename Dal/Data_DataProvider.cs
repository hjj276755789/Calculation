using Calculation.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Dal
{
    public class Data_DataProvider
    {
        public int ADD_JH(int nf,List<int> zc)
        {
            
            StringBuilder sb = new StringBuilder("insert into  calculation. xtgl_sjrwjhb (nf,zc) values ");
            string sql = "";
            foreach (int item in zc)
            {
                sb.Append(string.Format(@"('{0}','{1}'),", nf, item));
            }
            sql = sb.ToString();
            return MySqlDbhelper.ExecuteNonQuery(sql.Substring(0, sql.Length - 1));
        }
        public List<Data_JHSQXQ> GET_JHXQ(int nf)
        {
            string sql = @"select jhbh,t1.nf,t1.zc,
case when t2.sl is not null then t2.sl else 0 end cjjl ,
case when t3.sl is not null then t3.sl else 0 end xzys ,
case when t4.sl is not null then t4.sl else 0 end tdcj ,
case when t5.sl is not null then t5.sl else 0 end rgsj  from  calculation. xtgl_sjrwjhb t1 
left join (select nf,zc,count(*) sl from calculation.xtgl_data_zb_cjba  group by nf,zc) t2 on t1.nf=t2.nf and t1.zc =t2.zc
left join (select nf,zc,count(*) sl from calculation.xtgl_data_zb_xzys  group by nf,zc) t3 on t1.nf=t3.nf and t1.zc =t3.zc
left join (select nf,zc,count(*) sl from calculation.xtgl_data_zb_tdcj  group by nf,zc) t4 on t1.nf=t4.nf and t1.zc =t4.zc
left join (select nf,zc,count(*) sl from calculation.xtgl_data_zb_rgsj  group by nf,zc) t5 on t1.nf=t5.nf and t1.zc =t5.zc
 where t1.nf = @nf
 order by t1.nf,t1.zc
";
            MySqlParameter[] p = { new MySqlParameter("nf", nf)};
            return Models.Modelhelper.类列表赋值(new Data_JHSQXQ(), MySqlDbhelper.GetDataSet(sql,p).Tables[0]);
        }

        public Data_JHSQXQ GET_JHXQ(int nf,int zc)
        {
            string sql = @"select jhbh,t1.nf,t1.zc,
case when t2.sl is not null then t2.sl else 0 end cjjl ,
case when t3.sl is not null then t3.sl else 0 end xzys ,
case when t4.sl is not null then t4.sl else 0 end tdcj ,
case when t5.sl is not null then t5.sl else 0 end rgsj  from  calculation. xtgl_sjrwjhb t1 
left join (select nf,zc,count(*) sl from calculation.xtgl_data_zb_cjba  group by nf,zc) t2 on t1.nf=t2.nf and t1.zc =t2.zc
left join (select nf,zc,count(*) sl from calculation.xtgl_data_zb_xzys  group by nf,zc) t3 on t1.nf=t3.nf and t1.zc =t3.zc
left join (select nf,zc,count(*) sl from calculation.xtgl_data_zb_tdcj  group by nf,zc) t4 on t1.nf=t4.nf and t1.zc =t4.zc
left join (select nf,zc,count(*) sl from calculation.xtgl_data_zb_rgsj  group by nf,zc) t5 on t1.nf=t4.nf and t1.zc =t4.zc

 where t1.nf = @nf and t1.zc= @zc
 order by t1.nf,t1.zc
";
            MySqlParameter[] p = { new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
            return Models.Modelhelper.类对象赋值(new Data_JHSQXQ(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }

    }
}
