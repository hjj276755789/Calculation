using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public static class Base_dw
    {
        /// <summary>
        /// 面积
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static double mj(this double target)
        {
            return Math.Round(target, 0);
        }
        
        /// <summary>
        /// 面积_万方
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static double mj_wf(this double target)
        {
            return Math.Round(target / 10000, 2);
        }
        /// <summary>
        /// 面积_万方_描述
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static string mj_wf_ms(this double target)
        {

            try
            {
                string str = (target > 0 ? "增加" : "减少") + Math.Abs(Math.Round(target / 10000, 2)) +"万方";
                return str;
            }
            catch (Exception)
            {

                return "0";
            }
        }
        /// <summary>
        /// 面积_方转亩
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static double mj_f_to_m(this double target)
        {
            //方转亩
            return Math.Round(target, 2);
        }
        /// <summary>
        /// 面积_亩
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static double mj_m(this double target)
        {
            return Math.Round(target, 2);
        }
        /// <summary>
        /// 面积_亩转方
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static double mj_m_to_f(this double target)
        {
            //亩转方
            return Math.Round(target, 2);
        }

        /// <summary>
        /// 金额_万元
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static double je_wy(this double target)
        {
            return Math.Round(target / 10000, 0);
        }
        public static double je_wy_2(this double target)
        {
            return Math.Round(target / 10000, 2);
        }
        public static double je_wy(this long target)
        {
            return Math.Round(target.doubls() / 10000, 0);
        }
        /// <summary>
        /// 金额_亿元
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static double je_yy(this long target)
        {
            double temp = target / 100000000.00;
            return Math.Round(temp, 2);
        }
        public static double je_y(this long target)
        {
            double temp = target / 1.0;
            return Math.Round(temp, 0);
        }

        public static double je_w_to_y(this double target)
        {
            return Math.Round(target * 10000.00, 2);
        }
        public static double je_w_to_yy(this double target)
        {
            return Math.Round(target / 10000.00, 2);
        }

        public static double je_w_to_yy(this long target)
        {
            return Math.Round(target/ 10000.00, 2);
        }
        /// <summary>
        /// 金额_元
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static double je_y(this double target)
        {
            return Math.Round(target,0);
        }
        /// <summary>
        /// 算式_百分比
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static string ss_bfb(this double target)
        {
            try
            {
                double temp = Math.Abs(target) * 100;
                string str = (target > 0 ? "增加" : "减少") + Math.Round(temp, 2) + "%";
                return str;
            }
            catch (Exception)
            {

                return "0";
            }
            
        }

        public static double ss_bfb_ys(this double target)
        {
            try
            {
                return Math.Abs(Math.Round(target , 2));
            }
            catch (Exception)
            {
                return 0;
            }

        }
        /// <summary>
        /// 百分比 绝对值
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static string ss_bfb_jdz(this double target)
        {
            return Math.Abs(Math.Round(target*100, 2)) + "%";
        }
        /// <summary>
        /// 百分比 数值
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static string ss_bfb_sz(this double target)
        {
            return Math.Round(target * 100, 2) + "%";
        }
    }
}
