using Calculation.Dal;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Calculation.Base;

namespace Calculation.JS
{

    /// <summary>
    /// 周报缓存数据
    /// </summary>
    public class Cache_Result_zb
    {
        private static Cache_Result_zb uniqueInstance;

        public static Cache_Result_zb ini()
        {

            if (uniqueInstance == null)
            {
                uniqueInstance = new Cache_Result_zb();
                jsjg_xkpzs = Jsjg_zb_DataProvider.GET_XKPZS();
            }
            return uniqueInstance;
        }

        #region 计算结果
        public static DataTable jsjg_xkpzs { get; set; }
        #endregion

        #region 土地
        #region 土地计算结果
        private static double _td_bz_zyd { get; set; }
        private static double _td_bz_kjtl { get; set; }
        private static long _td_bz_cjje { get; set; }
        private static int _td_bz_cjsl { get; set; }
        private static double _td_sz_zyd { get; set; }
        private static double _td_sz_kjtl { get; set; }
        private static double _td_sz_cjje { get; set; }
        private static double _td_sz_cjsl { get; set; }


        #endregion
        //-------------------------本月----------------------------//
        /// <summary>
        /// 成交土地面积（总用地）
        /// </summary>
        public static double td_bz_zyd
        {
            get
            {
                if (_td_bz_zyd == 0)
                {
                    _td_bz_zyd = Cache_data_tdjyjl.bz.AsEnumerable().Sum(m => double.Parse(m["zyd_m"].ToString()));
                }
                return _td_bz_zyd;
            }
        }
        /// <summary>
        /// 可建体量
        /// </summary>
        public static double td_bz_kjtl
        {
            get
            {
                if (_td_bz_kjtl == 0)
                {
                    _td_bz_kjtl = Cache_data_tdjyjl.bz.AsEnumerable().Sum(m => double.Parse(m["kjtl_wf"].ToString()));
                }
                return _td_bz_kjtl;
            }
        }
        /// <summary>
        /// 土地综合出让金
        /// </summary>
        public static long td_bz_cjje
        {
            get
            {
                if (_td_bz_cjje == 0)
                {
                    _td_bz_cjje = Cache_data_tdjyjl.bz.AsEnumerable().Sum(m => m["cjzj"].longs());
                }
                return _td_bz_cjje;
            }
        }
        /// <summary>
        /// 土地成交数量（本周）
        /// </summary>
        public static int td_bz_cjsl
        {
            get
            {
                if (_td_bz_cjsl == 0)
                {
                    _td_bz_cjsl = Cache_data_tdjyjl.bz.AsEnumerable().Count();
                }
                return _td_bz_cjsl;
            }
        }
        //----------------------上周-------------------------//
        public static double td_sz_zyd
        {
            get
            {
                if (_td_sz_zyd == 0)
                {
                    _td_sz_zyd = Cache_data_tdjyjl.sz.AsEnumerable().Sum(m => double.Parse(m["zyd_m"].ToString()));
                }
                return _td_sz_zyd;
            }
        }
        /// <summary>
        /// 可建体量
        /// </summary>
        public static double td_sz_kjtl
        {
            get
            {
                if (_td_sz_kjtl == 0)
                {
                    _td_sz_kjtl = Cache_data_tdjyjl.sz.AsEnumerable().Sum(m => double.Parse(m["kjtl_wf"].ToString()));
                }
                return _td_sz_kjtl;
            }
        }
        /// <summary>
        /// 土地综合出让金
        /// </summary>
        public static double td_sz_cjje
        {
            get
            {
                if (_td_sz_cjje == 0)
                {
                    _td_sz_cjje = Cache_data_tdjyjl.sz.AsEnumerable().Sum(m => double.Parse(m["cjzj"].ToString()));
                }
                return _td_sz_cjje;
            }
        }
      
        #endregion

        #region 成交记录
        #region 成交记录计算结果
        //----------------------本周-------------------------//
        private static double _bz_cj_jzmj_xzys { get; set; }
        private static double _bz_cj_jzmj_fzz_xzys { get; set; }
        private static double _bz_cj_jzmj { get; set; }
        private static double _bz_cj_tnmj { get; set; }
        private static long _bz_cj_cjje { get; set; }
        private static EnumerableRowCollection<DataRow> _bz_cj_czz { get; set; }
        private static EnumerableRowCollection<DataRow> _bz_cj_czz_xzys { get; set; }
        //----------------------本周-------------------------//
        private static double _sz_cj_jzmj_xzys { get; set; }
        private static double _sz_cj_jzmj_fzz_xzys { get; set; }
        private static double _sz_cj_jzmj { get; set; }
        private static double _sz_cj_tnmj { get; set; }
        private static long _sz_cj_cjje { get; set; }
        private static EnumerableRowCollection<DataRow> _sz_cj_czz { get; set; }

        private static EnumerableRowCollection<DataRow> _sz_cj_czz_xzys { get; set; }

        //----------------------同周-------------------------//
        private static double _tz_cj_jzmj_xzys { get; set; }
        private static double _tz_cj_jzmj_fzz_xzys { get; set; }
        private static double _tz_cj_jzmj { get; set; }
        private static double _tz_cj_tnmj { get; set; }
        private static long _tz_cj_cjje { get; set; }
        private static EnumerableRowCollection<DataRow> _tz_cj_czz { get; set; }

        private static EnumerableRowCollection<DataRow> _tz_cj_czz_xzys { get; set; }
        #endregion

        //----------------------本周-------------------------//
        /// <summary>
        /// 成交记录-建筑面积-新增预售
        /// <summary>
        public static double bz_cj_jzmj_xzys
        {
            get
            {
                if (_bz_cj_jzmj_xzys == 0)
                {
                    _bz_cj_jzmj_xzys = Cache_data_xzys.bz.AsEnumerable().Sum(m =>m["jzmj"].doubls() +m["fzzmj"].doubls());
                }
                return _bz_cj_jzmj_xzys;
            }
        }
        /// <summary>
        /// 成交记录-建筑面积-非住宅-新增预售
        /// </summary>
        public static double bz_cj_jzmj_fzz_xzys
        {
            get
            {
                if (_bz_cj_jzmj_fzz_xzys == 0)
                {
                    _bz_cj_jzmj_fzz_xzys = Cache_data_xzys.bz.AsEnumerable().Sum(m => m["fzzmj"].doubls());
                }
                return _bz_cj_jzmj_fzz_xzys;
            }
        }
        /// <summary>
        /// 成交记录-建筑面积
        /// </summary>
        public static double bz_cj_jzmj
        {
            get
            {
                if (_bz_cj_jzmj == 0)
                {
                    _bz_cj_jzmj = Cache_data_cjjl.bz.AsEnumerable().Sum(m => m["jzmj"].doubls());
                }
                return _bz_cj_jzmj;
            }
        }
        /// <summary>
        /// 成交记录-套内面积
        /// </summary>
        public static double bz_cj_tnmj
        {
            get
            {
                if (_bz_cj_tnmj == 0)
                {
                    _bz_cj_tnmj = Cache_data_cjjl.bz.AsEnumerable().Sum(m => m["tnmj"].doubls());
                }
                return _bz_cj_tnmj;
            }
        }
        /// <summary>
        /// 成交记录-成交金额
        /// </summary>
        public static long bz_cj_cjje
        {
            get
            {
                if (_bz_cj_cjje == 0)
                {
                    _bz_cj_cjje = Cache_data_cjjl.bz.AsEnumerable().Sum(m => m["cjje"].longs());
                }
                return _bz_cj_cjje;
            }
        }

        public static EnumerableRowCollection<DataRow> bz_cj_czz
        {
            get
            {
                if (_bz_cj_czz == null)
                {
                    _bz_cj_czz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["yt"].ToString() == "别墅" || m["yt"].ToString() == "高层" || m["yt"].ToString() == "小高层" || m["yt"].ToString() == "洋房" || m["yt"].ToString() == "洋楼");
                }
                return _bz_cj_czz;
            }
        }

        public static EnumerableRowCollection<DataRow> bz_cj_czz_xzys
        {
            get
            {
                if (_bz_cj_czz_xzys == null)
                {
                    _bz_cj_czz_xzys = Cache_data_xzys.bz.AsEnumerable().Where(m => m["tyyt"].ToString() == "别墅" || m["tyyt"].ToString() == "高层" || m["tyyt"].ToString() == "小高层" || m["tyyt"].ToString() == "洋房" || m["tyyt"].ToString() == "洋楼");
                }
                return _bz_cj_czz_xzys;
            }
        }

        ////----------------------上周-------------------------//
        public static double sz_cj_jzmj_xzys
        {
            get
            {
                if (_sz_cj_jzmj_xzys == 0)
                {
                    _sz_cj_jzmj_xzys = Cache_data_xzys.sz.AsEnumerable().Sum(m => m["jzmj"].doubls());
                }
                return _sz_cj_jzmj_xzys;
            }
        }
        /// <summary>
        /// 成交记录-建筑面积-非住宅-新增预售
        /// </summary>
        public static double sz_cj_jzmj_fzz_xzys
        {
            get
            {
                if (_sz_cj_jzmj_fzz_xzys == 0)
                {
                    _sz_cj_jzmj_fzz_xzys = Cache_data_xzys.sz.AsEnumerable().Sum(m => m["fzzmj"].doubls());
                }
                return _sz_cj_jzmj_fzz_xzys;
            }
        }
        /// <summary>
        /// 成交记录-建筑面积
        /// </summary>
        public static double sz_cj_jzmj
        {
            get
            {
                if (_sz_cj_jzmj == 0)
                {
                    _sz_cj_jzmj = Cache_data_cjjl.sz.AsEnumerable().Sum(m => m["jzmj"].doubls());
                }
                return _sz_cj_jzmj;
            }
        }
        /// <summary>
        /// 成交记录-套内面积
        /// </summary>
        public static double sz_cj_tnmj
        {
            get
            {
                if (_sz_cj_tnmj == 0)
                {
                    _sz_cj_tnmj = Cache_data_cjjl.sz.AsEnumerable().Sum(m => m["tnmj"].doubls());
                }
                return _sz_cj_tnmj;
            }
        }
        /// <summary>
        /// 成交记录-成交金额
        /// </summary>
        public static double sz_cj_cjje
        {
            get
            {
                if (_sz_cj_cjje == 0)
                {
                    _sz_cj_cjje = Cache_data_cjjl.sz.AsEnumerable().Sum(m =>m["cjje"].longs());
                }
                return _sz_cj_cjje;
            }
        }

        public static EnumerableRowCollection<DataRow> sz_cj_czz
        {
            get
            {
                if (_sz_cj_czz == null)
                {
                    _sz_cj_czz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["yt"].ToString() == "别墅" || m["yt"].ToString() == "高层" || m["yt"].ToString() == "小高层" || m["yt"].ToString() == "洋房" || m["yt"].ToString() == "洋楼");
                }
                return _sz_cj_czz;
            }
        }
        public static EnumerableRowCollection<DataRow> sz_cj_czz_xzys
        {
            get
            {
                if (_sz_cj_czz_xzys == null)
                {
                    _sz_cj_czz_xzys = Cache_data_xzys.sz.AsEnumerable().Where(m => m["tyyt"].ToString() == "别墅" || m["tyyt"].ToString() == "高层" || m["tyyt"].ToString() == "小高层" || m["tyyt"].ToString() == "洋房" || m["tyyt"].ToString() == "洋楼");
                }
                return _sz_cj_czz_xzys;
            }
        }



        ////----------------------同周-------------------------//
        public static double tz_cj_jzmj_xzys
        {
            get
            {
                if (_tz_cj_jzmj_xzys == 0)
                {
                    _tz_cj_jzmj_xzys = Cache_data_xzys.tz.AsEnumerable().Sum(m => m["jzmj"].doubls());
                }
                return _tz_cj_jzmj_xzys;
            }
        }
        /// <summary>
        /// 成交记录-建筑面积-非住宅-新增预售
        /// </summary>
        public static double tz_cj_jzmj_fzz_xzys
        {
            get
            {
                if (_tz_cj_jzmj_fzz_xzys == 0)
                {
                    _tz_cj_jzmj_fzz_xzys = Cache_data_xzys.tz.AsEnumerable().Sum(m => m["fzzmj"].doubls());
                }
                return _tz_cj_jzmj_fzz_xzys;
            }
        }
        /// <summary>
        /// 成交记录-建筑面积
        /// </summary>
        public static double tz_cj_jzmj
        {
            get
            {
                if (_tz_cj_jzmj == 0)
                {
                    _tz_cj_jzmj = Cache_data_cjjl.tz.AsEnumerable().Sum(m => m["jzmj"].doubls());
                }
                return _tz_cj_jzmj;
            }
        }
        /// <summary>
        /// 成交记录-套内面积
        /// </summary>
        public static double tz_cj_tnmj
        {
            get
            {
                if (_tz_cj_tnmj == 0)
                {
                    _tz_cj_tnmj = Cache_data_cjjl.tz.AsEnumerable().Sum(m => m["tnmj"].doubls());
                }
                return _tz_cj_tnmj;
            }
        }
        /// <summary>
        /// 成交记录-成交金额
        /// </summary>
        public static double tz_cj_cjje
        {
            get
            {
                if (_tz_cj_cjje == 0)
                {
                    _tz_cj_cjje = Cache_data_cjjl.tz.AsEnumerable().Sum(m => m["cjje"].longs());
                }
                return _tz_cj_cjje;
            }
        }

        public static EnumerableRowCollection<DataRow> tz_cj_czz
        {
            get
            {
                if (_tz_cj_czz == null)
                {
                    _tz_cj_czz = Cache_data_cjjl.tz.AsEnumerable().Where(m => m["yt"].ToString() == "别墅" || m["yt"].ToString() == "高层" || m["yt"].ToString() == "小高层" || m["yt"].ToString() == "洋房" || m["yt"].ToString() == "洋楼");
                }
                return _tz_cj_czz;
            }
        }

        #endregion
    }
}
