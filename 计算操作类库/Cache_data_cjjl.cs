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
    /// <summary>
    /// 成交记录一级缓存
    /// </summary>
    public class Cache_data_cjjl
    {
        private static dateTask nowdatetask;
        private static Cache_data_cjjl uniqueInstance;

        public static Cache_data_cjjl ini_yb(int year,int month)
        {
            if (uniqueInstance == null)
            { 
                
                uniqueInstance= new Cache_data_cjjl();
                by = ZB_Data_CJBA_DataProvider.GET_YB(Base_date.by_First, Base_date.by_Last);
                sy = ZB_Data_CJBA_DataProvider.GET_YB(Base_date.sy_First, Base_date.sy_Last);
            }
           
            return uniqueInstance;
        }

        public static Cache_data_cjjl ini_zb()
        {
            if (uniqueInstance == null)
            {
                uniqueInstance = new Cache_data_cjjl();
                jbz = ZB_Data_CJBA_DataProvider.GET_JBZ(Base_date.bn,Base_date.bz);
                var bztemp = jbz.Select("zc=" + Base_date.bz);
                bz = bztemp.Count() != 0 ? bztemp.CopyToDataTable():new DataTable();
                var sztemp = jbz.Select("zc=" + (Base_date.bz - 1));
                sz = sztemp.Count() != 0 ? sztemp.CopyToDataTable() : new DataTable();
                var ssztemp = jbz.Select("zc=" + (Base_date.bz - 2));
                ssz = ssztemp.Count() != 0 ? ssztemp.CopyToDataTable() : new DataTable();
                var sssztemp = jbz.Select("zc=" + (Base_date.bz - 3));
                sssz = sssztemp.Count() != 0 ? sssztemp.CopyToDataTable() : new DataTable();
                tz = ZB_Data_CJBA_DataProvider.GET_ZB(Base_date.tz_first, Base_date.tz_Last);
                nowdatetask = new dateTask(null, Base_date.bn, null, Base_date.bz);
            }
            else
            {
                if (nowdatetask != null && (nowdatetask.nf != Base_date.bn || nowdatetask.zc !=Base_date.bz))
                {
                    jbz = ZB_Data_CJBA_DataProvider.GET_JBZ(Base_date.bn, Base_date.bz);
                    var bztemp = jbz.Select("zc=" + Base_date.bz);
                    bz = bztemp.Count() != 0 ? bztemp.CopyToDataTable() : new DataTable();
                    var sztemp = jbz.Select("zc=" + (Base_date.bz - 1));
                    sz = sztemp.Count() != 0 ? sztemp.CopyToDataTable() : new DataTable();
                    var ssztemp = jbz.Select("zc=" + (Base_date.bz - 2));
                    ssz = ssztemp.Count() != 0 ? ssztemp.CopyToDataTable() : new DataTable();
                    var sssztemp = jbz.Select("zc=" + (Base_date.bz - 3));
                    sssz = sssztemp.Count() != 0 ? sssztemp.CopyToDataTable() : new DataTable();
                    tz = ZB_Data_CJBA_DataProvider.GET_ZB(Base_date.tz_first, Base_date.tz_Last);
                    nowdatetask = new dateTask(null,Base_date.bn, null, Base_date.bz);
                }

            }
            return uniqueInstance;
        }

        /// <summary>
        /// 本月
        /// </summary>
        public static DataTable by {
            get;set;
        }
        public static DataTable sy
        {
            get; set;
        }
        public static DataTable ty
        {
            get; set;
        }

        /// <summary>
        /// 本年
        /// </summary>
        public static DataTable bn { get; set; }
        /// <summary>
        /// 本周
        /// </summary>
        public static DataTable bz{ get; set; }
        /// <summary>
        /// 上周
        /// </summary>
        public static DataTable sz { get; set; }
        /// <summary>
        /// 上上周
        /// </summary>
        public static DataTable ssz { get; set; }
        /// <summary>
        /// 上上上周
        /// </summary>
        public static DataTable sssz { get; set; }
        /// <summary>
        /// 同周
        /// </summary>
        public static DataTable tz { get; set; }
        /// <summary>
        /// 近八周
        /// </summary>
        public static DataTable jbz { get; set; }

    }
}
