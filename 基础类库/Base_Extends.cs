using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public static class Base_Extends
    {
        /// <summary>
        /// 单位整数
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static int dw_zs(this string target)
        {
            if (string.IsNullOrEmpty(target))
                return 0;
            try
            {
                return Int32.Parse(target);
            }
            catch (Exception)
            {
                return 0;
            }
        }
        /// <summary>
        /// 单位小数
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static double dw_xs(this string target)
        {
            if (string.IsNullOrEmpty(target))
                return 0.0;
            try
            {
                return Math.Round(double.Parse(target),2);
            }
            catch (Exception)
            {
                return 0.0;
            }
        }
        public static double dw_xs(this double target)
        {
         
            try
            {
                return Math.Round(target, 2);
            }
            catch (Exception)
            {
                return 0.0;
            }
        }
        public static bool IsNull(this string target)
        {
            if (target == null || target == "" || target.Length == 0)
                return true;
            else return false;
        }

        public static int ints(this object target)
        {
            if (target == null)
                return 0;
            try
            {

                return Convert.ToInt32(target);
            }
            catch (Exception e)
            {
                Base_Log.Log(e.Message);
                if (target != null)
                    return (int)target;
                else return 0;
            }
        }
        public static long longs(this object target)
        {
            if (target == null)
                return 0;
            try
            {
                return long.Parse(target.ToString());
            }
            catch (Exception)
            {

                return 0;
            }
        }
        public static double doubls(this object target)
        {
            if (target == null)
                return 0;
            try
            {
                return double.Parse(target.ToString());
            }
            catch (Exception)
            {

                return 0;
            }
        }

        public static string  Join(this string [] target)
        {
            if (target == null)
                return "—";
            try
            {
                return string.Join("—", target);
            }
            catch (Exception)
            {

                return "—";
            }
        }
        public static bool IsNotNull(this string [] target)
        {
            if (target != null && target.Length > 0)
                return true;
            else return false;
        }


        public static string timeUpper(this DateTime target)
        {
            return Base_date.dateToUpper(target);
        }
        /// <summary>
        /// 获取成交备案名称
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static string _ConfigCjbaMc(this string target)
        {
            try
            {
                if (!string.IsNullOrEmpty(target)) {
                    var s = target.Substring(target.IndexOf('_',0)+1, target.Length - target.IndexOf('_', 0)-1);
                    return s;
                }
                else return target;
            }
            catch (Exception)
            {

                return "";
            }

        }
        /// <summary>
        /// 获取认购数据名称
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static string _ConfigRgsjMc(this string target)
        {
            try
            {
                if (!string.IsNullOrEmpty(target))
                {
                    var s = target.Substring(target.IndexOf('_', 0) + 1, target.Length - target.IndexOf('_', 0) - 1);
                    return s;
                }
                else return target;
            }
            catch (Exception)
            {

                return "";
            }

        }

    }


    public static class DateTimeExtensions
    {
        /// <summary>
        /// 将日期转换为整型标识，如20160630
        /// </summary>
        /// <param name="dateTime">日期</param>
        public static int ToDateInt(this DateTime dateTime)
        {
            return int.Parse(dateTime.ToString("yyyyMMdd"));
        }
        public static bool IsNullOrEmpty(this string target)
        {
            return string.IsNullOrEmpty(target);
        }

        /// <summary>
        /// 将日期转换为字符串标识（yyyyMMddHHmm ：201606301230）
        /// </summary>
        /// <param name="dateTime">日期</param>
        public static string ToDateStr(this DateTime dateTime)
        {
            return dateTime.ToString("yyyyMMddHHmm");
        }

        /// <summary>
        /// 判断时间1是否大于时间2
        /// </summary>
        /// <param name="date1"></param>
        /// <param name="date2"></param>
        /// <returns></returns>
        public static bool DateDiff(string date1, string date2)
        {
            DateTime dt1 = Convert.ToDateTime(date1.DateTimeFormat());
            DateTime dt2 = Convert.ToDateTime(date2.DateTimeFormat());
            return dt1 >= dt2;
        }

        public static string ToDateStr(this string dateTime)
        {
            try
            {
                return DateTime.Parse(dateTime).ToString("yyyyMMddHHmm");
            }
            catch (Exception)
            {
                DateTime dt = DateTime.ParseExact(dateTime, "yyyyMMddHHmm", null);
                return dt.ToString("yyyy-MM-dd");
            }

        }
        public static string ToDateStr1(this string dateTime)
        {
            DateTime dt = DateTime.ParseExact(dateTime, "yyyyMMddHHmm", null);
            return dt.ToString("yyyy-MM-dd");
        }
        public static string ToDateStr2(this string dateTime)
        {
            DateTime dt = DateTime.ParseExact(dateTime, "yyyy-MM-dd", null);
            return dt.ToString("yyyyMMdd");
        }
    }


    public static class DateTimeStringExtensions
    {
        /// <summary>
        /// 将形同“201705191230”12位字符串转换为“2017-05-19 12：30”格式
        /// </summary>
        public static string DateTimeFormat(this string target)
        {
            if (target.IsNullOrEmpty() || target.Length != 12)
            {
                return target;
            }
            return string.Format("{0}-{1}-{2} {3}:{4}",
                target.Substring(0, 4),
                target.Substring(4, 2),
                target.Substring(6, 2),
                target.Substring(8, 2),
                target.Substring(10, 2));
        }


        /// <summary>
        /// 将形同“201705191230”12位字符串转换为日期类型
        /// </summary>
        public static DateTime ToDate(this string target)
        {
            return DateTime.Parse(target);
        }


        /// <summary>
        /// 将形同“201705191230”12位字符串转换为“2017-05-19”格式
        /// </summary>
        public static string DateTimeFormat1(this string target)
        {
            if (target.IsNullOrEmpty() || target.Length != 12)
            {
                return target;
            }
            return string.Format("{0}-{1}-{2}",
                target.Substring(0, 4),
                target.Substring(4, 2),
                target.Substring(6, 2));
        }



        /// <summary>
        /// 将形同“201705191230”12位字符串转换为“2017-05-19 12：30”或“05-19 12：30”格式
        /// </summary>
        public static string DateTimeFormat2(this string target)
        {
            if (target.IsNullOrEmpty() || target.Length != 12)
            {
                return target;
            }
            string temp = target.DateTimeFormat();
            string year = temp.Substring(0, 4);
            if (year == DateTime.Today.Year.ToString())
            {
                return temp.Substring(5);
            }
            else
            {
                return temp;
            }
        }



        /// <summary>
        /// 将形同“201705191230”12位字符串转换为“05-19”格式
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static string DateTimeFormatRQ(this string target)
        {
            if (target.IsNullOrEmpty() || target.Length != 12)
            {
                return target;
            }

            return string.Format("{0}/{1}",
                target.Substring(4, 2),
                target.Substring(6, 2));
        }
        public static string DateTimeFormatRQ1(this string target)
        {
            if (target.IsNullOrEmpty() || target.Length != 12)
            {
                target = target.ToDateStr();
            }
            return string.Format("{0}/{1}",
                target.Substring(4, 2),
                target.Substring(6, 2));
        }
        /// <summary>
        /// 将形同“201705191230”12位字符串转换为“12：30”格式
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static string DateTimeFormatSJ(this string target)
        {
            if (target.IsNullOrEmpty() || target.Length != 12)
            {
                return target;
            }
            return string.Format("{0}:{1}",
                target.Substring(8, 2),
                target.Substring(10, 2));
        }
    }
}
