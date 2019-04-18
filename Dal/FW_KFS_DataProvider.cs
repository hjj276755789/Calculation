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
    }
}
