using Calculation.Base;
using Calculation.Models;
using Calculation.Models.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Dal
{
    public class Dg_DataProvider : MySqlDbhelper
    {
        public IPageList<DGModels> GET_DG(string tj,string nf,string zc, int pagesize, int pagenow)
        {
            string sql = @"select b1.kfsbh,b1.kfsmc,b2.* from 
(select a1.*,a2.mbbh from xtgl_kfs_xx a1,xtgl_kfs_kfsmb a2 where a1.kfsbh=a2.kfsbh ) b1,(
select t2.nf,t2.zc,t2.mbid,t2.mbmc,t3.rwmc,case when t3.xzdz2 is not null then 1 else 0 end kfxz,t3.xzdz2 from(
select * from xtgl_sjrwjhb,xtgl_bbmb)t2 left join xtgl_bbrw t3 on t2.mbid =t3.mbid and t2.nf=t3.nf and t2.zc=t3.zc
where  t2.nf=@nf and t2.zc=@zc ";
            if (tj.IsNull())
            {
                sql += ") b2 where b1.mbbh =b2.mbid order by b1.kfsbh";
                MySqlParameter[] p = { new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
                return GetPagedList<DGModels>(sql, p, pagesize, pagenow);
               
            }
            else
            {
                sql += " ) b2 where  b1.mbbh =b2.mbid and b1.kfsmc like @kfsmc order by b1.kfsbh";
                MySqlParameter[] p = { new MySqlParameter("kfsmc", "%" + tj + "%"), new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
                return GetPagedList<DGModels>(sql, p, pagesize, pagenow);
            }
        
           
        }
    }
}
