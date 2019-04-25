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
    /// <summary>
    /// 开发商模块
    /// </summary>
    public class FW_KFS_DataProvider:MySqlDbhelper
    {
        public IPageList<KfsModels> FIND_KFS_XX(string tj,int pagesize,int pagenow)
        {
            string sql = "select * from xtgl_kfs_xx";
            if (!tj.IsNull())
            {
                sql += " where kfsmc like @kfsmc ";
                MySqlParameter[] p = { new MySqlParameter("kfsmc", "%" + tj + "%") };
                return GetPagedList<KfsModels>(sql, p, pagesize, pagenow);
            }
            else
            {
                return GetPagedList<KfsModels>(sql, null, pagesize, pagenow);
            }
            
        }

        public bool ADD_KFS(string kfsmc, string kfslxr,string kfslxrdh,string bz)
        {
            string sql = "insert into xtgl_kfs_xx (kfsmc,kfslxr,kfslxrdh,kfscjsj,kfszt,bz) values (@kfsmc,@kfslxr,@kfslxrdh,@kfscjsj,@kfszt,@bz)";
            MySqlParameter[] p = { new MySqlParameter("kfsmc", kfsmc), new MySqlParameter("kfslxr", kfslxr), new MySqlParameter("kfslxrdh", kfslxrdh), new MySqlParameter("kfscjsj", DateTime.Now.ToDateStr()), new MySqlParameter("kfszt", 1), new MySqlParameter("bz", bz) };
            return ExecuteNonQuery(sql, p) > 0;
        }


        /// <summary>
        /// 删除开发商信息
        /// </summary>
        /// <param name="kfsbh"></param>
        /// <returns></returns>
        public bool DEL_KFS(string kfsbh)
        {
            string sql = "delete from xtgl_kfs_xx where kfsbh =@kfsbh";
            MySqlParameter[] p = { new MySqlParameter("kfsbh", kfsbh) };
            return ExecuteNonQuery(sql, p) > 0;
        }

        /// <summary>
        /// 删除开发商模板
        /// </summary>
        /// <param name="kfsbh"></param>
        /// <returns></returns>
        public bool DEL_KFS_FZR(string kfsbh)
        {
            string sql = "delete from xtgl_kfs_kfsmb where kfsbh =@kfsbh";
            MySqlParameter[] p = { new MySqlParameter("kfsbh", kfsbh) };
            return ExecuteNonQuery(sql, p) > 0;
        }
        /// <summary>
        /// 删除用户负责开发商
        /// </summary>
        /// <param name="kfsbh"></param>
        /// <returns></returns>
        public bool DEL_KFS_YHFZKFS(string kfsbh)
        {
            string sql = "delete from xtgl_fw_yhfzkfs where kfsbh =@kfsbh";
            MySqlParameter[] p = { new MySqlParameter("kfsbh", kfsbh) };
            return ExecuteNonQuery(sql, p) > 0;
        }
        public IPageList<KFSMBModels> FIND_KFS_MB(string kfsbh,int pagesize,int pagenow)
        {
            string sql = @"select * from (
                        select a1.mbid,a1.mbmc,count(a1.mbid) rwcs from xtgl_bbmb a1 left join xtgl_bbrw a2 on a1.mbid = a2.mbid group by a1.mbid,a1.mbmc
                        ) t1, xtgl_kfs_kfsmb t2
                        where t1.mbid = t2.mbbh and t2.kfsbh =@kfsbh";
            MySqlParameter[] p = { new MySqlParameter("kfsbh", kfsbh) };
            return GetPagedList<KFSMBModels>(sql, p, pagesize, pagenow);

        }

        public IPageList<YHFZKFSModels> FIND_YHFZKFSBH(string yhbh, string tj, int pagesize, int pagenow)
        {
            string sql = @"select t1.*,case when t2.kfsbh is not null then 1 else 0 end sffp from 
                xtgl_kfs_xx t1 left join 
              (select kfsbh from xtgl_fw_yhfzkfs a1 where a1.yhbh = @yhbh ) t2 on t1.kfsbh =t2.kfsbh ";
            if (!tj.IsNull())
            {
                sql += " where kfsmc like @kfsmc ";
                MySqlParameter[] p = { new MySqlParameter("yhbh",yhbh),new MySqlParameter("kfsmc", "%" + tj + "%") };
                return GetPagedList<YHFZKFSModels>(sql, p, pagesize, pagenow);
            }
            else
            {
                MySqlParameter[] p = { new MySqlParameter("yhbh", yhbh) };
                return GetPagedList<YHFZKFSModels>(sql, p, pagesize, pagenow);
            }

        }
    }
}
