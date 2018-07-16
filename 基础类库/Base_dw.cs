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
            return Math.Round(target, 2);
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
        /// <summary>
        /// 金额_亿元
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public static double je_yy(this long target)
        {
            var t = target / 10000 / 10000;
            return Math.Round(t.doubls(), 2);
        }


        public static double je_w_to_y(this double target)
        {
            return Math.Round(target * 10000 , 2);
        }
        public static double je_w_to_yy(this double target)
        {
            return Math.Round(target / 10000, 2);
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
                string str = (target > 0 ? "增加" : "减少") + Math.Abs(Math.Round(target * 100, 2)) + "%";
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

    }
}
