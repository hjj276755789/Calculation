using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public class Base_date
    {
        public static void init_zb(int year, int weak)
        {
            bn = year;
            bz = weak;

            bz_first= CalcWeekDay_first(year, weak);
            bz_Last = CalcWeekDay_last(year, weak);

            sz_first = CalcWeekDay_first(year, weak-1);
            sz_Last = CalcWeekDay_last(year, weak-1);

            tz_first = CalcWeekDay_first(year -1, weak );
            tz_Last = CalcWeekDay_last(year -1, weak);

            bzwz = Base_date.bz_first.ToString("MM.dd") + "_" + Base_date.bz_Last.ToString("MM.dd");
        }

        public static void init_yb(int year, int month)
        {
            bn = year;
            by = month;


            by_First = new DateTime(year, month,1);
            by_Last = by_First.AddDays(DateTime.DaysInMonth(year, month) - 1);

            sy_First = by_First.AddMonths(-1).AddDays(1 - (by_First.Day));
            sy_Last = sy_First.AddDays(DateTime.DaysInMonth(year, month) - 1);

            ty_first = by_First.AddYears( - 1).AddDays(1 - (by_First.Day)); ;
            ty_Last = ty_first.AddDays(DateTime.DaysInMonth(year, month) - 1);
        }

        #region 年
        public static int bn { get; set; }
        #endregion

        #region 月
        /// <summary>
        /// 本月（实际为上月）
        /// </summary>
        private static int by { get; set; }
        /// <summary>
        /// 本月第一天（实际为上月）
        /// </summary>
        public static DateTime by_First { get; set; }


        /// <summary>
        /// 本月最后一天（实际为上月）
        /// </summary>
        public static DateTime by_Last { get; set; }

        /// <summary>
        /// 上月第一天（实际为上上月）
        /// </summary>
        public static DateTime sy_First { get; set; }


        /// <summary>
        /// 上月最后一天（实际为上上月）
        /// </summary>
        public static DateTime sy_Last { get; set; }


        public static DateTime ty_first { get; set; }
        public static DateTime ty_Last { get; set; }
        #endregion

        #region 周

        /// <summary>
        /// 本周是第几周
        /// </summary>
        public static int bz { get; set; }

        public static DateTime bz_first { get; set; }
        public static DateTime bz_Last { get; set; }

        public static DateTime sz_first { get; set; }
        public static DateTime sz_Last { get; set; }


        public static DateTime tz_first { get; set; }
        public static DateTime tz_Last { get; set; }

        /// <summary>
        /// 本周文字
        /// </summary>
        public static string bzwz { get; set; }   
        #endregion


        #region 周计算方法
        private static DateTime CalcWeekDay_first(int year, int week)
        {
            DateTime first = DateTime.MinValue;
            //指定年范围  
            DateTime start = new DateTime(year, 1, 1);
            DateTime end = new DateTime(year, 12, 31);
            int startWeekDay = (int)start.DayOfWeek;
            //周的起始日期  
            first = start.AddDays((7 - startWeekDay) + (week - 2) * 7 +1);
            return first;
        }
        private static DateTime CalcWeekDay_last(int year, int week)
        {
            DateTime first = DateTime.MinValue;
            DateTime last = DateTime.MinValue;
            DateTime start = new DateTime(year, 1, 1);
            DateTime end = new DateTime(year, 12, 31);
            int startWeekDay = (int)start.DayOfWeek;
            //周的起始日期  
            first = start.AddDays((7 - startWeekDay) + (week - 2) * 7);
            last = first.AddDays(7);
            return last;

        }


        #endregion


        public static List<int> GET_Z_OF_Y(int year)
        {
            List<int> zc = new List<int>();
            DateTime start = new DateTime(year, 1, 1);
            DateTime end = new DateTime(year, 12, 31);
            int startWeekDay = (int)start.DayOfWeek;
            int dayofyear = (int)end.DayOfYear + startWeekDay + (7 -(int)end.DayOfWeek);
            
            for (int i = 1; i <= dayofyear; i++)
            {
                if (i % 7 == 0)
                    zc.Add(i / 7);
            }
            return zc;

        }

        public static string GET_ZCMC(int year,int weak)
        {
            DateTime dt_first = CalcWeekDay_first(year, weak);
            DateTime dt_Last = CalcWeekDay_last(year, weak);
            return dt_first.ToString("M.d") + "-" + dt_Last.ToString("M.d");
        }
        public static string GET_NFZCMC(int year, int weak)
        {
            DateTime dt_first = CalcWeekDay_first(year, weak);
            DateTime dt_Last = CalcWeekDay_last(year, weak);
            return year+"."+dt_first.ToString("M.d") + "-" + year + "." + dt_Last.ToString("M.d");
        }




        private static string numtoUpper(int num)
        {
            String str = num.ToString();
            string rstr = "";
            int n;
            for (int i = 0; i < str.Length; i++)
            {
                n = Convert.ToInt16(str[i].ToString());//char转数字,转换为字符串，再转数字
                switch (n)
                {
                    case 0: rstr = rstr + "〇"; break;
                    case 1: rstr = rstr + "一"; break;
                    case 2: rstr = rstr + "二"; break;
                    case 3: rstr = rstr + "三"; break;
                    case 4: rstr = rstr + "四"; break;
                    case 5: rstr = rstr + "五"; break;
                    case 6: rstr = rstr + "六"; break;
                    case 7: rstr = rstr + "七"; break;
                    case 8: rstr = rstr + "八"; break;
                    default: rstr = rstr + "九"; break;
                }

            }
            return rstr;
        }
        //月转化为大写
        private static string monthtoUpper(int month)
        {
            if (month < 10)
            {
                return numtoUpper(month);
            }
            else
                if (month == 10) { return "十"; }

            else
            {
                return "十" + numtoUpper(month - 10);
            }
        }
        //日转化为大写
        private static string daytoUpper(int day)
        {
            if (day < 20)
            {
                return monthtoUpper(day);
            }
            else
            {
                String str = day.ToString();
                if (str[1] == '0')
                {
                    return numtoUpper(Convert.ToInt16(str[0].ToString())) + "十";
                }
                else
                {
                    return numtoUpper(Convert.ToInt16(str[0].ToString())) + "十"
                        + numtoUpper(Convert.ToInt16(str[1].ToString()));
                }
            }
        }
        //日期转换为大写
        public static string dateToUpper(System.DateTime date)
        {
            int year = date.Year;
            int month = date.Month;
            int day = date.Day;
            return numtoUpper(year) + "年" + monthtoUpper(month) + "月" + daytoUpper(day) + "日";

        }
    }
    /// <summary>
    ///  
    /// </summary>
    public class dateTask
    {
        public dateTask(int ? mbid,int nf,int? yf,int? zc)
        {
            this.nf = nf;
            if (mbid.HasValue)
                this.mbid = mbid.Value;
            if (yf.HasValue)
                this.yf = yf.Value;
            if (zc.HasValue)
                this.zc = zc.Value;
        }
        public int nf { get; set; }
        public int yf { get; set; }
        public int zc { get; set; }

        public int mbid { get; set; }

    }

}
