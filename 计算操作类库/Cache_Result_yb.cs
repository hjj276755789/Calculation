using Calculation.Dal;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    /// <summary>
    ///  缓存_月报
    /// </summary>
    public class Cache_Result_yb
    {

        private static Cache_Result_yb uniqueInstance;

        public static Cache_Result_yb ini()
        {

            if (uniqueInstance == null)
            {
                uniqueInstance = new Cache_Result_yb();
                jsjg_scgxfx = Jsjg_yb_DataProvider.GET_SCGXFX();
                jsjg_scgxfx_psb = Jsjg_yb_DataProvider.GET_SCGXFX_PSB();
                jsjg_scjgfx = Jsjg_yb_DataProvider.GET_SCJGFX();
            }
            return uniqueInstance;
        }

        #region 计算结果

       
        /// <summary>
        /// 计算结果_市场供需分析
        /// </summary>
        public static DataTable jsjg_scgxfx
        {
            get; set;
        }

        /// <summary>
        /// 计算结果_市场供需分析_批售比
        /// </summary>
        public static DataTable jsjg_scgxfx_psb
        {
            get; set;
        }
        /// <summary>
        /// 计算结果_市场结构分析
        /// </summary>
        public static DataTable jsjg_scjgfx
        {
            get; set;
        }
        #endregion

        #region 土地
        #region 土地计算结果
        private static double _td_by_zyd { get; set; }
        private static double _td_by_kjtl { get; set; }
        private static double _td_by_cjje { get; set; }

        private static double _td_sy_zyd { get; set; }
        private static double _td_sy_kjtl { get; set; }
        private static double _td_sy_cjje { get; set; }


        #endregion
        //-------------------------本月----------------------------//
        /// <summary>
        /// 成交土地面积（总用地）
        /// </summary>
        public static double td_by_zyd
        {
            get {
                if (_td_by_zyd == 0)
                {
                    _td_by_zyd = Cache_data_tdjyjl.by.AsEnumerable().Sum(m => double.Parse(m[3].ToString()));
                }
                return _td_by_zyd;
            }
        }
        /// <summary>
        /// 可建体量
        /// </summary>
        public static double td_by_kjtl
        {
            get {
                if (_td_by_kjtl == 0)
                {
                    _td_by_kjtl = Cache_data_tdjyjl.by.AsEnumerable().Sum(m => double.Parse(m[4].ToString()));
                }
                return _td_by_kjtl;
            }
        }
        /// <summary>
        /// 土地综合出让金
        /// </summary>
        public static double td_by_cjje
        {
            get
            {
                if (_td_by_cjje == 0)
                {
                    _td_by_cjje = Cache_data_tdjyjl.by.AsEnumerable().Sum(m => double.Parse(m[5].ToString()));
                }
                return _td_by_cjje;
            }
        }
        //----------------------上月-------------------------//
        public static double td_sy_zyd
        {
            get
            {
                if (_td_sy_zyd == 0)
                {
                    _td_sy_zyd = Cache_data_tdjyjl.sy.AsEnumerable().Sum(m => double.Parse(m[3].ToString()));
                }
                return _td_sy_zyd;
            }
        }
        /// <summary>
        /// 可建体量
        /// </summary>
        public static double td_sy_kjtl
        {
            get
            {
                if (_td_sy_kjtl == 0)
                {
                    _td_sy_kjtl = Cache_data_tdjyjl.sy.AsEnumerable().Sum(m => double.Parse(m[4].ToString()));
                }
                return _td_sy_kjtl;
            }
        }
        /// <summary>
        /// 土地综合出让金
        /// </summary>
        public static double td_sy_cjje
        {
            get
            {
                if (_td_sy_cjje == 0)
                {
                    _td_sy_cjje = Cache_data_tdjyjl.sy.AsEnumerable().Sum(m => double.Parse(m[5].ToString()));
                }
                return _td_sy_cjje;
            }
        }
        #endregion

        #region 成交记录
        #region 成交记录计算结果
        //----------------------本月-------------------------//
        private static double _by_cj_jzmj_xzys { get; set; }
        private static double _by_cj_jzmj_fzz_xzys { get; set; }
        private static double _by_cj_jzmj { get; set; }
        private static double _by_cj_tnmj { get; set; }
        private static double _by_cj_cjje { get; set; }
        private static EnumerableRowCollection<DataRow> _by_cj_czz { get; set; }
        //----------------------上月-------------------------//
        private static double _sy_cj_jzmj_xzys { get; set; }
        private static double _sy_cj_jzmj_fzz_xzys { get; set; }
        private static double _sy_cj_jzmj { get; set; }
        private static double _sy_cj_tnmj { get; set; }
        private static double _sy_cj_cjje { get; set; }
        private static EnumerableRowCollection<DataRow> _sy_cj_czz { get; set; }
        #endregion

        //----------------------本月-------------------------//
        /// <summary>
        /// 成交记录-建筑面积-新增预售
        /// <summary>
        public static double by_cj_jzmj_xzys
        {
            get { 
                if(_by_cj_jzmj_xzys ==0)
                {
                    _by_cj_jzmj_xzys = Cache_data_xzys.by.AsEnumerable().Sum(m => double.Parse(m[10].ToString()));
                }
                return _by_cj_jzmj_xzys;
            }
        }
        /// <summary>
        /// 成交记录-建筑面积-非住宅-新增预售
        /// </summary>
        public static double by_cj_jzmj_fzz_xzys
        {
            get
            {
                if (_by_cj_jzmj_fzz_xzys == 0)
                {
                    _by_cj_jzmj_fzz_xzys = Cache_data_xzys.by.AsEnumerable().Sum(m => string.IsNullOrEmpty(m[11].ToString()) ? 0 : double.Parse(m[11].ToString()));
                }
                return _by_cj_jzmj_fzz_xzys;
            }
        }
        /// <summary>
        /// 成交记录-建筑面积
        /// </summary>
        public static double by_cj_jzmj
        {
            get
            {
                if (_by_cj_jzmj == 0)
                {
                    _by_cj_jzmj = Cache_data_cjjl.by.AsEnumerable().Sum(m => double.Parse(m["jzmj"].ToString()));
                }
                return _by_cj_jzmj;
            }
        }
        /// <summary>
        /// 成交记录-套内面积
        /// </summary>
        public static double by_cj_tnmj
        {
            get
            {
                if (_by_cj_tnmj == 0)
                {
                    _by_cj_tnmj = Cache_data_cjjl.by.AsEnumerable().Sum(m => double.Parse(m["tnmj"].ToString()));
                }
                return _by_cj_tnmj;
            }
        }
        /// <summary>
        /// 成交记录-成交金额
        /// </summary>
        public static double by_cj_cjje
        {
            get
            {
                if (_by_cj_cjje == 0)
                {
                    _by_cj_cjje = Cache_data_cjjl.by.AsEnumerable().Sum(m => double.Parse(m["cjje"].ToString()));
                }
                return _by_cj_cjje;
            }
        }

        public static EnumerableRowCollection<DataRow> by_cj_czz
        {
            get
            {
                if (_by_cj_czz == null)
                {
                    _by_cj_czz = Cache_data_cjjl.by.AsEnumerable().Where(m => m["ytmc"].ToString() == "别墅" || m["ytmc"].ToString() == "高层" || m["ytmc"].ToString() == "小高层" || m["ytmc"].ToString() == "洋房" || m["ytmc"].ToString() == "洋楼");
                }
                return _by_cj_czz;
            }
        }


        ////----------------------上月-------------------------//
        public static double sy_cj_jzmj_xzys
        {
            get
            {
                if (_sy_cj_jzmj_xzys == 0)
                {
                    _sy_cj_jzmj_xzys = Cache_data_xzys.sy.AsEnumerable().Sum(m => double.Parse(m["jzmj"].ToString()));
                }
                return _sy_cj_jzmj_xzys;
            }
        }
        /// <summary>
        /// 成交记录-建筑面积-非住宅-新增预售
        /// </summary>
        public static double sy_cj_jzmj_fzz_xzys
        {
            get
            {
                if (_sy_cj_jzmj_fzz_xzys == 0)
                {
                    _sy_cj_jzmj_fzz_xzys = Cache_data_xzys.sy.AsEnumerable().Sum(m => string.IsNullOrEmpty(m["fzzmj"].ToString()) ? 0 : double.Parse(m["fzzmj"].ToString()));
                }
                return _sy_cj_jzmj_fzz_xzys;
            }
        }
        /// <summary>
        /// 成交记录-建筑面积
        /// </summary>
        public static double sy_cj_jzmj
        {
            get
            {
                if (_sy_cj_jzmj == 0)
                {
                    _sy_cj_jzmj = Cache_data_cjjl.sy.AsEnumerable().Sum(m => double.Parse(m["jzmj"].ToString()));
                }
                return _sy_cj_jzmj;
            }
        }
        /// <summary>
        /// 成交记录-套内面积
        /// </summary>
        public static double sy_cj_tnmj
        {
            get
            {
                if (_sy_cj_tnmj == 0)
                {
                    _sy_cj_tnmj = Cache_data_cjjl.sy.AsEnumerable().Sum(m => double.Parse(m["tnmj"].ToString()));
                }
                return _sy_cj_tnmj;
            }
        }
        /// <summary>
        /// 成交记录-成交金额
        /// </summary>
        public static double sy_cj_cjje
        {
            get
            {
                if (_sy_cj_cjje == 0)
                {
                    _sy_cj_cjje = Cache_data_cjjl.sy.AsEnumerable().Sum(m => double.Parse(m["cjje"].ToString()));
                }
                return _sy_cj_cjje;
            }
        }

        public static EnumerableRowCollection<DataRow> sy_cj_czz
        {
            get
            {
                if (_sy_cj_czz == null)
                {
                    _sy_cj_czz = Cache_data_cjjl.sy.AsEnumerable().Where(m => m["ytmc"].ToString() == "别墅" || m["ytmc"].ToString() == "高层" || m["ytmc"].ToString() == "小高层" || m["ytmc"].ToString() == "洋房" || m["ytmc"].ToString() == "洋楼");
                }
                return _sy_cj_czz;
            }
        }

        #endregion

    }
}
