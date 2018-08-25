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

     

        //public static string 备案套数 = "ts";
        //public static string 成交金额 = "cjje";
        //public static string 建筑面积 = "jzmj";
        //public static string 套内面积 = "tnmj";
        //public static string 建面均价 = "jmjj";
        //public static string 套内均价 = "tnjj";

        //public static string[] _备案数据 = {"ts", "cjje", "jzmj", "tnmj" , "jmjj", "tnjj" };


      
    }
    /// <summary>
    /// _认购数据
    /// </summary>
    public class Base_Config_Rgsj
    {
        
        public const string 主力建面区间态 = "zljmqj";
        public const string 主力套内面积区间 = "zltnqj";
        public const string 新开套内均价 = "xktnjj";
        public const string 新开建面均价 = "xkjmjj";

        public const string 本周_新开套数     = "bz_xkts";
        public const string 本周_新开销售套数 = "bz_xkxsts";
        public const string 本周_新开建面均价 = "bz_xkjmjj";
        public const string 本周_新开套内均价 = "bz_xktnjj";
        public const string 本周_套均总价 = "bz_tjzj";
        public const string 本周_认购套数     = "bz_rgts";
        public const string 本周_认购套内均价 = "bz_rgtnjj";
        public const string 本周_认购建面均价 = "bz_rgjmjj";
        public const string 本周_认购套内体量 = "bz_rgtntl";
        public const string 本周_认购建面体量 = "bz_rgjmtl";
        public const string 本周_认购金额     = "bz_rxgje";
       

        public const string 上周_新开套数     = "sz_xkts";
        public const string 上周_新开销售套数 = "sz_xkxsts";
        public const string 上周_新开建面均价 = "sz_xkjmjj";
        public const string 上周_新开套内均价 = "sz_xktnjj";
        public const string 上周_套均总价 = "sz_tjzj";
        public const string 上周_认购套数     = "sz_rgts";
        public const string 上周_认购套内均价 = "sz_rgtnjj";
        public const string 上周_认购建面均价 = "sz_rgjmjj";
        public const string 上周_认购套内体量 = "sz_rgtntl";
        public const string 上周_认购建面体量 = "sz_rgjmtl";
        public const string 上周_认购金额     = "sz_rgje";
        


        public const string 成交套数环比 = "cjtshb";
        public const string 套内均价环比 = "tnjjhb";
        public const string 变化原因 = "bhyy";
        public const string 本周库存 = "bzkc";
        public const string 本周来电 = "bzld";
        public const string 本周到访量 = "bzdfl";
        public const string 优惠 = "yh";
        public const string 营销动作 = "yxdz";
        public const string 活动 = "hd";
        public const string 下周加推预计 = "xzjtyj";

        public static string[] _认购数据 = {"xkts", "xkxsts", "zljmqj", "zltnqj", "xktnjj", "xkjmjj",
                                            "bz_xkts","bz_xkxsts","bz_xkjmjj","bz_xktnjj","bz_tjzj","bz_rgts","bz_rgtnjj","bz_rgjmjj","bz_rgtntl","bz_rgjmtl","bz_rgje",
                                            "sz_xkts","sz_xkxsts","sz_xkjmjj","sz_xktnjj","sz_tjzj","sz_rgts","sz_rgtnjj","sz_rgjmjj","sz_rgtntl","sz_rgjmtl","sz_rgje",
                                            "cjtnhb", "tnjjhb", "bhyy", "bzkc", "bzld", "bzdfl", "yh", "yxdz", "hd", "xzjtyj"
                                            };

    }

    public class Base_Config_Cjba
    {
        public const string 本周_备案套数 = "bz_ts";
        public const string 本周_成交金额 = "bz_cjje";
        public const string 本周_建筑面积 = "bz_jzmj";
        public const string 本周_套内面积 = "bz_tnmj";
        public const string 本周_建面均价 = "bz_jmjj";
        public const string 本周_套内均价 = "bz_tnjj";
        public const string 本周_套均总价 = "bz_tjzj";

        public const string 上周_备案套数 = "sz_ts";
        public const string 上周_成交金额 = "sz_cjje";
        public const string 上周_建筑面积 = "sz_jzmj";
        public const string 上周_套内面积 = "sz_tnmj";
        public const string 上周_建面均价 = "sz_jmjj";
        public const string 上周_套内均价 = "sz_tnjj";
        public const string 上周_套均总价 = "sz_tjzj";

        public static string[] _备案数据 = { "bz_ts","bz_cjje","bz_jzmj","bz_tnmj","bz_jmjj","bz_tnjj", "bz_tjzj",
                                             "sz_ts","sz_cjje","sz_jzmj","sz_tnmj","sz_jmjj","sz_tnjj","sz_tjzj"
        };
    }

    public class Base_Config_Jzgj
    {
        public const string 竞争格局名称 = "jzgjmc";
        public const string 竞争格局_主力面积区间 = "zlmjqj";
        public const string 项目名称 = "lpmc";
        public const string 业态 = "yt";
        public const string 组团 = "zt";
        public static string [] _竞争格局参数名称 = { "jzgjmc", "zlmjqj", "lpmc", "yt" , "zt" };
    }

}
