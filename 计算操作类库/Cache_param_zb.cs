using Calculation.Base;
using Calculation.Models;
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
        private static dateTask nowdatetask;
        private static Cache_param_zb uniqueInstance;
        public static Cache_param_zb ini_zb(int mbid,int nf,int zc)
        {
            if (uniqueInstance == null)
            {
                uniqueInstance = new Cache_param_zb();
                value = Dal.Param_DataProvider.GET_MBCJCSNR(mbid,nf,zc);
                _param_jp = _param_jp_helper(mbid, nf, zc);
                nowdatetask = new dateTask(mbid, Base_date.bn, null, Base_date.bz);
            }
            else
            {
                if (nowdatetask != null && (nowdatetask.mbid !=mbid|| nowdatetask.nf != Base_date.bn || nowdatetask.zc != Base_date.bz))
                {
                    value = Dal.Param_DataProvider.GET_MBCJCSNR(mbid, nf, zc);
                    _param_jp = _param_jp_helper(mbid, nf, zc);
                    nowdatetask = new dateTask(null, Base_date.bn, null, Base_date.bz);
                }
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
                jp.qycs = item["qycs"] == null || string.IsNullOrEmpty(item["qycs"].ToString()) ? null : item["qycs"].ToString().Split(',');
                jp.ztcs = item["ztcs"] == null || string.IsNullOrEmpty(item["ztcs"].ToString()) ? null : item["ztcs"].ToString().Split(',');
                jp.lpcs = item["lpcs"] == null || string.IsNullOrEmpty(item["lpcs"].ToString()) ? null : item["lpcs"].ToString().Split(',');
                jp.ytcs = item["ytcs"] == null || string.IsNullOrEmpty(item["ytcs"].ToString()) ? null : item["ytcs"].ToString().Split(',');
                jp.xfytcs = item["xfytcs"] == null || string.IsNullOrEmpty(item["xfytcs"].ToString()) ? null : item["xfytcs"].ToString().Split(',');
                jp.hxcs = item["hxcs"] == null || string.IsNullOrEmpty(item["hxcs"].ToString()) ? null : item["hxcs"].ToString().Split(',');
                jp.zlmjqj = item["zlmjqj"].ToString() ;
                jp.jpxmlb = new List<JP_JPXM_INFO>();
                var xmlist = jptable.AsEnumerable().Where(m => m["baid"].ints() == item["id"].ints()).OrderBy(m => m["px"]) ;
                foreach (var xm in xmlist)
                {
                    JP_JPXM_INFO jpxm = new JP_JPXM_INFO();
                    jpxm.id = xm["id"].ints();
                    jpxm.baid = xm["baid"].ints();
                    jpxm.jzgjid = xm["jzgjid"].ints();
                    jpxm.jzgjmc = xm["jzgjmc"].ToString();
                    jpxm.px = xm["px"].ints();
                    jpxm.qycs = xm["qycs"] == null || string.IsNullOrEmpty(xm["qycs"].ToString()) ? null : xm["qycs"].ToString().Split(',');
                    jpxm.ztcs = xm["ztcs"] == null || string.IsNullOrEmpty(xm["ztcs"].ToString()) ? null : xm["ztcs"].ToString().Split(',');
                    jpxm.lpcs = xm["lpcs"] == null || string.IsNullOrEmpty(xm["lpcs"].ToString()) ? null : xm["lpcs"].ToString().Split(',');
                    jpxm.ytcs = xm["ytcs"] == null || string.IsNullOrEmpty(xm["ytcs"].ToString()) ? null : xm["ytcs"].ToString().Split(',');
                    jpxm.xfytcs = xm["xfytcs"] == null || string.IsNullOrEmpty(xm["xfytcs"].ToString()) ? null : xm["xfytcs"].ToString().Split(',');
                    jpxm.hxcs = xm["hxcs"] == null || string.IsNullOrEmpty(xm["hxcs"].ToString()) ? null : xm["hxcs"].ToString().Split(',');
                    jpxm.zlmjqj = xm["zlmjqj"].ToString();
                    jp.jpxmlb.Add(jpxm);
                }
                list.Add(jp);

            }
            return list;
        }
        #endregion
    }
}
