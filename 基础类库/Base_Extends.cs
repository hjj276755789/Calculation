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

        public static int ints(this object target)
        {
            if (target == null)
                return 0;
            try
            {
                return Int32.Parse(target.ToString());
            }
            catch (Exception e)
            {
                
                return (int)target;
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

        public static string timeUpper(this DateTime target)
        {
            return Base_date.dateToUpper(target);
        }



}
}
