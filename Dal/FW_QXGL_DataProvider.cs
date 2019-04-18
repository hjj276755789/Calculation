using Calculation.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Calculation.Base;
using Calculation.Models.Enums;

namespace Calculation.Dal
{
    public class FW_QXGL_DataProvider
    {
        //登录
        public YHXX CHECK_LOGIN(string yhmc, string yhmm)
        {
            string sql = "select * from xtgl_fw_yh where yhmc=@yhmc and yhmm=@yhmm";
            MySqlParameter[] p = { new MySqlParameter("yhmc", yhmc), new MySqlParameter("yhmm", yhmm) };
            return Modelhelper.类对象赋值<YHXX>(new YHXX(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }


        public List<YHXX> GET_YHLB()
        {
            string sql = "select * from xtgl_fw_yh";
            return Modelhelper.类列表赋值<YHXX>(new YHXX(), MySqlDbhelper.GetDataSet(sql).Tables[0]);
        }

        public bool ADD_USER(string yhmc,string yhmm,YH_LX yhlx)
        {
            string sql = "insert into xtgl_fw_yh(yhmc,yhmm,yhlx) values(@yhmc,@yhmm,@yhlx)";
            MySqlParameter[] p = { new MySqlParameter("yhmc", yhmc), new MySqlParameter("yhmm", yhmm), new MySqlParameter("yhlx", yhlx.ints()) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }

        //获取角色列表
        public List<JSXX> GET_JSLB()
        {
            string sql = "select * from xtgl_fw_js";
            return Modelhelper.类列表赋值<JSXX>(new JSXX(), MySqlDbhelper.GetDataSet(sql).Tables[0]);
        }
        public List<JSXX> GET_JSLB(int yhbh)
        {
            string sql = "select * from xtgl_fw_js t,xtgl_fw_yhjs t1 where t.jsbh =t1.jsbh and t1.yhbh=@yhbh ";
            MySqlParameter[] p = { new MySqlParameter("yhbh", yhbh) };
            return Modelhelper.类列表赋值<JSXX>(new JSXX(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        //获取权限列表
        public List<QXXX> GET_QXLB(int jsbh)
        {
            string sql = "select t.* from xtgl_fw_qxxx t,xtgl_fw_jsqx t1 where t.qxbh =t1.qxbh  and t1.jsbh=@jsbh";
            MySqlParameter[] p = { new MySqlParameter("jsbh", jsbh) };
            return Modelhelper.类列表赋值<QXXX>(new QXXX(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        /// <summary>
        /// 获取根权限列表
        /// </summary>
        public List<QXXX> GET_GQXLB()
        {
            string sql = "select * from xtgl_fw_qxxx where fqxbh is null";
            return Modelhelper.类列表赋值<QXXX>(new QXXX(), MySqlDbhelper.GetDataSet(sql).Tables[0]);
        }
        //获取权限列表
        public List<QXXX> GET_GQXLB(int jsid)
        {
            string sql = "select t.* from calculation.fw_qxb t,calculation.fw_jsqxb t1 where t.id =t1.qxid  and t1.jsid=@jsid and t.fid is null";
            MySqlParameter[] p = { new MySqlParameter("jsid", jsid) };
            return Modelhelper.类列表赋值<QXXX>(new QXXX(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        /// <summary>
        /// 获取所有权限列表
        /// </summary>
        /// <returns></returns>
        public List<QXXX> GET_QXLB()
        {
            string sql = "select * from calculation.fw_qxb";
            return Modelhelper.类列表赋值<QXXX>(new QXXX(), MySqlDbhelper.GetDataSet(sql).Tables[0]);
        }





        //获取用户角色权限
        public List<QXXX> GET_YHQX(string yhbh)
        {
            string sql = @"select  distinct e.qxbh,e.qxmc,e.qxkzq,e.qxst,e.fqxbh,e.qxlx,e.tb
from xtgl_fw_yh a,xtgl_fw_js b,xtgl_fw_yhjs c,xtgl_fw_jsqx d,xtgl_fw_qxxx e
where a.yhbh=c.yhbh and b.jsbh =c.jsbh and d.jsbh=b.jsbh and  e.qxbh=d.qxbh
and a.yhbh=@yhbh";
            MySqlParameter[] p = { new MySqlParameter("yhbh", yhbh) };
            return Modelhelper.类列表赋值<QXXX>(new QXXX(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        //设置用户角色
        public bool ADD_YHJS(int yhbh, int jsbh)
        {
            string sql = "insert into xtgl_fw_yhjs values(@yhbh,@jsbh)";
            MySqlParameter[] p = { new MySqlParameter("yhbh", yhbh), new MySqlParameter("jsbh", jsbh) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
        public bool DEL_YHJS(int yhbh, int jsbh)
        {
            string sql = "delete from xtgl_fw_yhjs where yhbh=@yhbh and jsbh=@jsbh";
            MySqlParameter[] p = { new MySqlParameter("yhbh", yhbh), new MySqlParameter("jsbh", jsbh) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }

        //设置角色权限
        public bool ADD_JSQX(int jsbh, int fqxbh)
        {
            string sql = @"insert into xtgl_fw_jsqx select @jsbh,@fqxbh union select @jsbh,qxbh from xtgl_fw_qxxx where fqxbh = @fqxbh";
            MySqlParameter[] p = { new MySqlParameter("jsbh", jsbh), new MySqlParameter("fqxbh", fqxbh) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
        public bool DEL_JSQX(int jsbh, int fqxbh)
        {
            string sql = @"DELETE FROM xtgl_fw_jsqx WHERE  jsbh=@jsbh AND qxbh in (select qxbh from xtgl_fw_qxxx where fqxbh = @fqxbh or qxbh=@fqxbh)";
            MySqlParameter[] p = { new MySqlParameter("jsbh", jsbh), new MySqlParameter("fqxbh", fqxbh) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }



        /// <summary>
        /// 是否拥有权限
        /// </summary>
        /// <param name="yhid"></param>
        /// <param name="controllerName"></param>
        /// <param name="actionName"></param>
        /// <returns></returns>

        public bool HAS_POWER(string yhid, string qxkzq, string qxst)
        {
            string sql = @"select count(t1.yhbh) ct from xtgl_fw_yh t1
join xtgl_fw_yhjs t2 on t1.yhbh=t2.yhbh
join xtgl_fw_js t3 on t3.jsbh= t2.jsbh
join xtgl_fw_jsqx t4 on t3.jsbh =t4.jsbh
join xtgl_fw_qxxx t5 on t4.qxbh =t5.qxbh 
where t1.yhbh=@yhbh and t5.qxkzq =@qxkzq and t5.qxst = @qxst";
            MySqlParameter[] p = { new MySqlParameter("yhbh", yhid), new MySqlParameter("qxkzq", qxkzq), new MySqlParameter("qxst", qxst) };
            var obj = MySqlDbhelper.ExecuteScalar(sql, p);
            if (obj != null)
                return int.Parse(obj.ToString()) > 0;
            else
                return false;
        }




        public bool DEL_USER(string yhbh)
        {
            string sql = "delete from xtgl_fw_yh where yhbh=@yhbh";
            MySqlParameter[] p = { new MySqlParameter("yhbh", yhbh) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
    }
}
