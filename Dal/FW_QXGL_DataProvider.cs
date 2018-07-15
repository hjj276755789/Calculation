using Calculation.Models.Models;
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
        public bool CHECK_LOGIN(string yhm,string yhmm)
        {
            string sql = "select count(*) from calculation.fw_yhb where yhm=@yhm and yhmm=@yhmm";
            MySqlParameter[] p = { new MySqlParameter("yhm", yhm), new MySqlParameter("yhmm", yhmm) };
            return MySqlDbhelper.ExecuteScalar(sql, p).ints() > 0;
        }

        
        public List<YHXX> GET_YHLB()
        {
            string sql = "select * from calculation.fw_yhb";
            return Models.Modelhelper.类列表赋值<YHXX>(new YHXX(), MySqlDbhelper.GetDataSet(sql).Tables[0]);
        }

        //获取角色列表
        public List<JSXX> GET_JSLB()
        {
            string sql = "select * from calculation.fw_jsb";
            return Models.Modelhelper.类列表赋值<JSXX>(new JSXX(), MySqlDbhelper.GetDataSet(sql).Tables[0]);
        }
        public List<JSXX> GET_JSLB(int yhid)
        {
            string sql = "select * from calculation.fw_jsb t,calculation.fw_yhjsb t1 where t.id =t1.jsid and t1.yhid=@yhid ";
            MySqlParameter[] p = { new MySqlParameter("yhid", yhid) };
            return Models.Modelhelper.类列表赋值<JSXX>(new JSXX(), MySqlDbhelper.GetDataSet(sql,p).Tables[0]);
        }
        //获取权限列表
        public List<QXXX> GET_QXLB(int jsid)
        {
            string sql = "select t.* from calculation.fw_qxb t,calculation.fw_jsqxb t1 where t.id =t1.qxid  and t1.jsid=@jsid";
            MySqlParameter[] p = { new MySqlParameter("jsid", jsid) };
            return Models.Modelhelper.类列表赋值<QXXX>(new QXXX(), MySqlDbhelper.GetDataSet(sql,p).Tables[0]);
        }
        /// <summary>
        /// 获取所有权限列表
        /// </summary>
        /// <returns></returns>
        public List<QXXX> GET_QXLB()
        {
            string sql = "select * from calculation.fw_qxb";
            return Models.Modelhelper.类列表赋值<QXXX>(new QXXX(), MySqlDbhelper.GetDataSet(sql).Tables[0]);
        }
        //获取用户角色权限
        public List<QXXX> GET_YHQX(string yhm)
        {
            string sql = @"select distinct t5.id,t5.qxmc,t5.qxkzq,t5.qxst from calculation.fw_yhb t1
join calculation.fw_yhjsb t2 on t1.id=t2.yhid
join calculation.fw_jsb t3 on t2.yhid= t1.id
join calculation.fw_jsqxb t4 on t3.id =t4.jsid
join calculation.fw_qxb t5 on t4.qxid =t5.id 
where t1.yhm=@yhm and t5.fid is null";
            MySqlParameter[] p = { new MySqlParameter("yhm", yhm) };
            return Models.Modelhelper.类列表赋值<QXXX>(new QXXX(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        //设置用户角色
        //设置角色权限





        /// <summary>
        /// 是否拥有权限
        /// </summary>
        /// <param name="yhid"></param>
        /// <param name="controllerName"></param>
        /// <param name="actionName"></param>
        /// <returns></returns>

        public bool HAS_POWER(string yhm, string qxkzq, string qxst)
        {
            string sql = @"select count(t1.id) from calculation.fw_yhb t1
join calculation.fw_yhjsb t2 on t1.id=t2.yhid
join calculation.fw_jsb t3 on t2.yhid= t1.id
join calculation.fw_jsqxb t4 on t3.id =t4.jsid
join calculation.fw_qxb t5 on t4.qxid =t5.id 
where t1.yhm=@yhm and t5.qxkzq =@qxkzq and t5.qxst = @qxst";
            MySqlParameter[] p = { new MySqlParameter("yhm", yhm), new MySqlParameter("qxkzq", qxkzq), new MySqlParameter("qxst", qxst) };
            var obj = MySqlDbhelper.ExecuteScalar(sql, p);
            if (obj != null)
                return int.Parse(obj.ToString()) > 0;
            else
                return false;
        }



        public bool ADD_USER(string yhm,string yhmm, YH_LX yhlx)
        {
            string sql = "insert into calculation.fw_yhb (yhm,yhmm,yhlx) values (@yhm,@yhmm,@yhlx)";
            MySqlParameter[] p = { new MySqlParameter("yhm", yhm), new MySqlParameter("yhmm", yhmm), new MySqlParameter("yhlx", yhlx) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
        public bool DEL_USER(string id)
        {
            string sql = "delete from calculation.fw_yhb where id=@id";
            MySqlParameter[] p = { new MySqlParameter("id", id)};
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
    }
}
