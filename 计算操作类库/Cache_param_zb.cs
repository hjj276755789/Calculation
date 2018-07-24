using Calculation.Base;
using Calculation.Models.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    public class Cache_param_zb
    {
        private static Cache_param_zb uniqueInstance;
        public static Cache_param_zb ini_zb(int mbid,int nf,int zc)
        {
            if (uniqueInstance == null)
            {
                uniqueInstance = new Cache_param_zb();
                value = Dal.Param_DataProvider.GET_MBCJCSNR(mbid,nf,zc);
                _param_jp = _param_jp_helper(mbid, nf, zc);
            }
            return uniqueInstance;
        }

        public static List<ParamValueModel> value { get; set; }


        public static List<JP_BA_INFO> _param_jp { get; set; }


        #region 内部方法
        private static List<JP_BA_INFO> _param_jp_helper(int mbid,int nf,int zc)
        {
            DataTable batable = Dal.Param_DataProvider.GET_JP_BAXX(mbid, nf, zc);
            DataTable jptable = Dal.Param_DataProvider.GET_JP_JPXMXX(mbid, nf, zc);
            List<JP_BA_INFO> list = new List<JP_BA_INFO>();
            foreach (DataRow item in batable.Rows)
            {
                JP_BA_INFO jp = new JP_BA_INFO();
                jp.cjid = item["cjid"].ints();
                jp.id = item["id"].ints();
                jp.bamc = item["bamc"].ToString();
                jp.rwid = item["rwid"].ints();
                jp.qycs = item["qycs"].ToString().Split(',');
                jp.ztcs = item["ztcs"].ToString().Split(',');
                jp.lpcs = item["lpcs"].ToString().Split(',');
                jp.ytcs = item["ytcs"].ToString().Split(',');
                jp.xfytcs = item["xfytcs"].ToString().Split(',');
                jp.hxcs = item["hxcs"].ToString().Split(',');
                jp.jpxmlb = new List<JP_JPXM_INFO>();
                var xmlist = jptable.AsEnumerable().Where(m => m["baid"].ints() == item["id"].ints());
                foreach (var xm in xmlist)
                {
                    JP_JPXM_INFO jpxm = new JP_JPXM_INFO();
                    jpxm.id = xm["id"].ints();
                    jpxm.baid = xm["baid"].ints();
                    jpxm.jzgjid = xm["jzgjid"].ints();
                    jpxm.qycs = xm["qycs"].ToString().Split(',');
                    jpxm.ztcs = xm["ztcs"].ToString().Split(',');
                    jpxm.lpcs = xm["lpcs"].ToString().Split(',');
                    jpxm.ytcs = xm["ytcs"].ToString().Split(',');
                    jpxm.xfytcs = xm["xfytcs"].ToString().Split(',');
                    jpxm.hxcs = xm["hxcs"].ToString().Split(',');
                    jp.jpxmlb.Add(jpxm);
                }
                list.Add(jp);

            }
            return list;
        }
        #endregion
    }
}
