using Calculation.Base;
using Calculation.Dal;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    public class Cache_data_tdjyjl
    {
        private static Cache_data_tdjyjl uniqueInstance;

        public static Cache_data_tdjyjl ini_yb()
        {
            if (uniqueInstance == null)
            {
                uniqueInstance = new Cache_data_tdjyjl();
                by = ZB_Data_TDCJ_DataProvider.GET_BY(Base_date.by_First, Base_date.by_Last);
                sy = ZB_Data_TDCJ_DataProvider.GET_BY(Base_date.sy_First, Base_date.sy_Last);
                ty = ZB_Data_TDCJ_DataProvider.GET_BY(Base_date.ty_first, Base_date.sy_Last);
            }
            return uniqueInstance;
        }

        public static Cache_data_tdjyjl ini_zb()
        {
            if (uniqueInstance == null)
            {
                uniqueInstance = new Cache_data_tdjyjl();
                bz = ZB_Data_TDCJ_DataProvider.GET_BY(Base_date.bz_first, Base_date.bz_Last);
                sz = ZB_Data_TDCJ_DataProvider.GET_BY(Base_date.sz_first, Base_date.sz_Last);
                tz = ZB_Data_TDCJ_DataProvider.GET_BY(Base_date.tz_first, Base_date.sz_Last);
            }
            return uniqueInstance;
        }
        /// <summary>
        /// 本月
        /// </summary>
        public static DataTable by
        {
            get; set;
        }
        public static DataTable sy
        {
            get; set;
        }
        public static DataTable ty
        {
            get; set;
        }

        public static DataTable bz
        {
            get; set;
        }
        public static DataTable sz
        {
            get; set;
        }
        public static DataTable tz
        {
            get; set;
        }
    }
}
