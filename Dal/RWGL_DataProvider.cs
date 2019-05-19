using Calculation.Models;
using Calculation.Models.Enums;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Calculation.Dal
{
    
    public class RWGL_DataProvider
    {
        #region 周报


        public bool Add_ZB(string rwmc, int mbid, int nf, int zc)
        {
            string sql = @"insert into calculation.xtgl_bbrw(rwmc,mbid,nf,zc,zt) values (@rwmc, @mbid, @nf, @zc, @zt)"; 
            MySqlParameter[] p = { new MySqlParameter("rwmc", rwmc), new MySqlParameter("mbid", mbid), new MySqlParameter("nf", nf), new MySqlParameter("zc", zc), new MySqlParameter("zt", 0) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p)>0;
        }
        public bool Del_ZB(int rwid)
        {
            string sql1 = @"delete from calculation.xtgl_bbrw where rwid=@rwid";
            string sql2 = "delete from calculation.xtgl_cj_rwcs where rwid =@rwid";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid)};
            int a = MySqlDbhelper.ExecuteNonQuery(sql1, p) ;
            int b= MySqlDbhelper.ExecuteNonQuery(sql2, p);
            return a+b>0;
        }
        public List<Rw_List> GET_ZB_RWLB(int mbid, int pagesize,int pagenow)
        {
            string sql = @"select * from calculation.xtgl_bbrw where mbid=@mbid limit @f,@e";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid), new MySqlParameter("f", pagesize * (pagenow - 1)), new MySqlParameter("e", pagesize * pagenow) };
            return Modelhelper.类列表赋值(new Rw_List(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        public IPageList<Zb_Item_Model> GET_ZB_LB(string yhbh,int pagesize, int pagenow,string mbmc,string kfsbh)
        {
            string sql = @"select t4.* from xtgl_fw_yhfzkfs t1,xtgl_kfs_xx t2,xtgl_kfs_kfsmb t3 ,xtgl_bbmb t4
where t1.kfsbh = t2.kfsbh and t2.kfsbh = t3.kfsbh and t3.mbbh = t4.mbid and t1.yhbh = @yhbh and t2.kfsbh =@kfsbh ";
            if (!string.IsNullOrEmpty(mbmc)) { 
                sql += @" and mblx = @mblx and mbmc like @mbmc ";
                MySqlParameter[] p = { new MySqlParameter("kfsbh", kfsbh), new MySqlParameter("yhbh", yhbh), new MySqlParameter("f", pagesize * (pagenow - 1)), new MySqlParameter("e", pagesize * pagenow), new MySqlParameter("mblx",(int) MB_Enums.周报), new MySqlParameter("mbmc", "%" + mbmc + "%") };
                return MySqlDbhelper.GetPagedList<Zb_Item_Model>(sql, p,pagesize,pagenow);
            }
            else{
                sql += @" and mblx = @mblx ";
                MySqlParameter[] p = { new MySqlParameter("kfsbh", kfsbh), new MySqlParameter("yhbh", yhbh),new MySqlParameter("f", pagesize * (pagenow - 1)), new MySqlParameter("e", pagesize * pagenow), new MySqlParameter("mblx",(int) MB_Enums.周报)};
                return MySqlDbhelper.GetPagedList<Zb_Item_Model>(sql, p, pagesize, pagenow);
            }
        }

        public Rw_Cofirm_data GET_RWZT(int rwid)
        {
            string sql = @"select t1.rwid,t1.rwmc,t1.nf,t1.zc,t2.cjjl_zt,t2.xzys_zt,t2.tdcj_zt,t2.rgsj_zt from calculation.xtgl_bbrw t1 left join  calculation.xtgl_ConfirmData t2 
on t1.rwid = t2.rwid where t1.rwid = @rwid";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid)};
            return Modelhelper.类对象赋值(new Rw_Cofirm_data(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        public Rw_Cofirm_data GET_RWZT(int rwid,int nf,int zc)
        {
            string sql = @"select t1.rwid,t1.rwmc,t1.nf,t1.zc,t2.cjjl_zt,t2.xzys_zt,t2.tdcj_zt,t2.rgsj_zt from calculation.xtgl_bbrw t1 left join  calculation.xtgl_ConfirmData t2 
on t1.rwid = t2.rwid where t1.rwid=@rwid and  t1.nf = @nf and t1.zc=@zc";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
            return Modelhelper.类对象赋值(new Rw_Cofirm_data(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        /// <summary>
        /// 确认成交数据
        /// </summary>
        /// <param name="rwid"></param>
        /// <param name="zt"></param>
        /// <returns></returns>
        public bool SET_DATA_ZT_CJ(int rwid, DATA_ZT zt)
        {
            string sql = "INSERT INTO calculation.xtgl_ConfirmData(rwid, cjjl_zt) VALUES(@rwid,@cjjl_zt) ON DUPLICATE KEY UPDATE cjjl_zt =@cjjl_zt";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("cjjl_zt", zt) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
        /// <summary>
        /// 确认新增预售
        /// </summary>
        /// <param name="rwid"></param>
        /// <param name="zt"></param>
        /// <returns></returns>
        public bool SET_DATA_ZT_XZ(int rwid, DATA_ZT zt)
        {

            string sql = "INSERT INTO calculation.xtgl_ConfirmData(rwid, xzys_zt) VALUES(@rwid,@xzys_zt) ON DUPLICATE KEY UPDATE xzys_zt =@xzys_zt";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("xzys_zt", zt) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
        /// <summary>
        /// 确认土地交易
        /// </summary>
        /// <param name="rwid"></param>
        /// <param name="zt"></param>
        /// <returns></returns>
        public bool SET_DATA_ZT_TD(int rwid, DATA_ZT zt)
        {

            string sql = "INSERT INTO calculation.xtgl_ConfirmData(rwid, tdcj_zt) VALUES(@rwid,@tdcj_zt) ON DUPLICATE KEY UPDATE tdcj_zt =@tdcj_zt";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("tdcj_zt", zt) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
        /// <summary>
        /// 确认认购数据
        /// </summary>
        /// <param name="rwid"></param>
        /// <param name="zt"></param>
        /// <returns></returns>
        public bool SET_DATA_ZT_RG(int rwid, DATA_ZT zt)
        {       
            string sql = "INSERT INTO calculation.xtgl_ConfirmData(rwid, rgsj_zt) VALUES(@rwid,@rgsj_zt) ON DUPLICATE KEY UPDATE rgsj_zt =@rgsj_zt";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("rgsj_zt", zt) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
        public bool SET_RWZT(int rwid,RW_ZT rwzt)
        {
            string sql = "update calculation.xtgl_bbrw  set zt =@zt where rwid=@rwid";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid), new MySqlParameter("zt", rwzt) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }
        public bool SET_RWZT(int mbid,int nf,int zc, RW_ZT rwzt,string xzdz)
        {
            string sql = "update calculation.xtgl_bbrw  set zt =@zt ,xzdz=@xzdz where mbid=@mbid and nf=@nf and zc=@zc";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid), new MySqlParameter("xzdz", xzdz), new MySqlParameter("nf", nf), new MySqlParameter("zc", zc), new MySqlParameter("zt", rwzt) };
            return MySqlDbhelper.ExecuteNonQuery(sql, p) > 0;
        }

        public Rw_Item_Model GET_RWXQ(int rwid)
        {
            string sql = "select * from  calculation.xtgl_bbrw where rwid=@rwid";
            MySqlParameter[] p = { new MySqlParameter("rwid", rwid) };
            return Modelhelper.类对象赋值<Rw_Item_Model>(new Rw_Item_Model(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        public Rw_Item_Model GET_RWXQ(int mbid,int nf,int zc)
        {
            string sql = "select * from  calculation.xtgl_bbrw where mbid= @mbid and nf=@nf and zc=@zc";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid), new MySqlParameter("nf", nf), new MySqlParameter("zc", zc) };
            return Modelhelper.类对象赋值<Rw_Item_Model>(new Rw_Item_Model(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }

        public ZB_ZJ_CS GET_ZJ_ZB_CS(int mbid)
        {
            string sql = "select rwid,nf,zc,mbid from xtgl_bbrw where (zt = 3 or zt =4) and mbid =@mbid order by nf desc,zc desc limit 1";
            MySqlParameter[] p = { new MySqlParameter("mbid", mbid) };
            return Modelhelper.类对象赋值<ZB_ZJ_CS>(new ZB_ZJ_CS(), MySqlDbhelper.GetDataSet(sql, p).Tables[0]);
        }
        #endregion
    }
}
