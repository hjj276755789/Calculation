using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public class Base_Config
    {
        public enum 坐标方向
        {
            横向 = 1,
            纵向 = 2
        }

     

        public static string 备案套数 = "ts";
        public static string 成交金额 = "cjje";
        public static string 建筑面积 = "jzmj";
        public static string 套内面积 = "tnmj";
        public static string 建面均价 = "jmjj";
        public static string 套内均价 = "tnjj";

        public static string[] _备案数据 = {"ts", "cjje", "jzmj", "tnmj" , "jmjj", "tnjj" };


      
    }
    /// <summary>
    /// _认购数据
    /// </summary>
    public class Base_Config_Rgsj
    {
        public const string 项目名称 = "xm";
        public const string 业态 = "yt";
        public const string 新开套数 = "xkts";
        public const string 新开销售套数 = "xkxsts";
        public const string 主力建面区间态 = "zljmqj";
        public const string 主力套内面积区间 = "zltnqj";
        public const string 新开套内均价 = "xktnjj";
        public const string 新开建面均价 = "xkjmjj";

        public const string 认购套数 = "rgts";
        public const string 认购套内均价 = "rgtnjj";
        public const string 认购建面均价 = "rgjmjj";
        public const string 认购套内体量 = "rgtnjj";
        public const string 认购建面体量 = "rgjmtl";
        public const string 认购金额 = "rgje";
               
        public const string 成交套数环比 = "cjtnhb";
        public const string 套内均价环比 = "tnjjhb";
        public const string 变化原因 = "bhyy";
        public const string 本周库存 = "bzkc";
        public const string 本周来电 = "bzld";
        public const string 本周到访量 = "bzdfl";
        public const string 优惠 = "yh";
        public const string 营销动作 = "yxdz";
        public const string 活动 = "hd";
        public const string 下周加推预计 = "xzjtyj";

        public static string[] _认购数据 = {"xm","yt","xkts", "xkxsts", "zljmqj", "zltnqj", "xktnjj", "xkjmjj",
                                          "rgts", "rgtnjj", "rgjmjj", "rgtnjj", "rgjmtl", "rgje",
                                           "cjtnhb", "tnjjhb", "bhyy", "bzkc", "bzld", "bzdfl", "yh", "yxdz", "hd", "xzjtyj"
                                            };

    }

    public class Base_Config_Cjba
    {
        public const string 备案套数 = "ts";
        public const string 成交金额 = "cjje";
        public const string 建筑面积 = "jzmj";
        public const string 套内面积 = "tnmj";
        public const string 建面均价 = "jmjj";
        public const string 套内均价 = "tnjj";

        public static string[] _备案数据 = { "ts", "cjje", "jzmj", "tnmj", "jmjj", "tnjj" };
    }

}
