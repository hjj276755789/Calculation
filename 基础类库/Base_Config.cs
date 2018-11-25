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
    }
    /// <summary>
    /// _认购数据
    /// </summary>
    public class Base_Config_Rgsj
    {
        
        public const string 本周_主力建面区间 = "bz_zljmqj";
        public const string 本周_主力套内面积区间 = "bz_zltnqj";
        public const string 新开套内均价 = "xktnjj";
        public const string 新开建面均价 = "xkjmjj";

        public const string 本周_新开套数     = "bz_xkts";
        public const string 本周_新开销售套数 = "bz_xkxsts";
        public const string 本周_新开建面均价 = "bz_xkjmjj";
        public const string 本周_新开套内均价 = "bz_xktnjj";
        public const string 本周_套均总价 = "bzrg_tjzj";
        public const string 本周_认购套数     = "bz_rgts";
        public const string 本周_认购套内均价 = "bz_rgtnjj";
        public const string 本周_认购建面均价 = "bz_rgjmjj";
        public const string 本周_认购套内体量 = "bz_rgtntl";
        public const string 本周_认购建面体量 = "bz_rgjmtl";
        public const string 本周_认购金额     = "bz_rgje";

        public const string 本周_认购套数环比 = "bz_rgtshb";
        public const string 本周_认购金额环比 = "bz_rgjehb";
        public const string 本周_认购建筑面积环比 = "bz_rgjzmjhb";
        public const string 本周_认购套内面积环比 = "bz_rgtnmjhb";
        public const string 本周_认购建面均价环比 = "bz_rgjmjjhb";
        public const string 本周_认购套内均价环比 = "bz_rgtnjjhb";
        public const string 本周_认购套均总价环比 = "bz_rgtjzjhb";


        public const string 本周_成交套数环比 = "bz_cjtshb";
        public const string 本周_套内均价环比 = "bz_tnjjhb";
        public const string 本周_变化原因 = "bz_bhyy";
        public const string 本周_本周库存 = "bz_bzkc";
        public const string 本周_本周来电 = "bz_bzld";
        public const string 本周_本周到访量 = "bz_bzdfl";
        public const string 本周_优惠 = "bz_yh";
        public const string 本周_营销动作 = "bz_yxdz";
        public const string 本周_活动 = "bz_hd";
        public const string 本周_下周加推预计 = "bz_xzjtyj";

        public const string 上周_新开套数     = "sz_xkts";
        public const string 上周_新开销售套数 = "sz_xkxsts";
        public const string 上周_新开建面均价 = "sz_xkjmjj";
        public const string 上周_新开套内均价 = "sz_xktnjj";
        public const string 上周_套均总价 = "szrg_tjzj";
        public const string 上周_认购套数     = "sz_rgts";
        public const string 上周_认购套内均价 = "sz_rgtnjj";
        public const string 上周_认购建面均价 = "sz_rgjmjj";
        public const string 上周_认购套内体量 = "sz_rgtntl";
        public const string 上周_认购建面体量 = "sz_rgjmtl";
        public const string 上周_认购金额     = "sz_rgje";
        public const string 上周_本周到访量 = "sz_bzdfl";
        public const string 上周_主力建面区间 = "sz_zljmqj";
        public const string 上周_主力套内面积区间 = "sz_zltnqj";
        public const string 上周_本周来电 = "sz_bzld";
        

        public const string 上上周_新开套数 = "ssz_xkts";
        public const string 上上周_新开销售套数 = "ssz_xkxsts";
        public const string 上上周_新开建面均价 = "ssz_xkjmjj";
        public const string 上上周_新开套内均价 = "ssz_xktnjj";
        public const string 上上周_套均总价 = "sszrg_tjzj";
        public const string 上上周_认购套数 = "sszrgts";
        public const string 上上周_认购套内均价 = "ssz_rgtnjj";
        public const string 上上周_认购建面均价 = "ssz_rgjmjj";
        public const string 上上周_认购套内体量 = "ssz_rgtntl";
        public const string 上上周_认购建面体量 = "ssz_rgjmtl";
        public const string 上上周_认购金额 = "ssz_rgje";
        public const string 上上周_主力建面区间 = "ssz_zljmqj";
        public const string 上上周_主力套内面积区间 = "ssz_zltnqj";

        public const string 上上上周_新开套数 = "sssz_xkts";
        public const string 上上上周_新开销售套数 = "sssz_xkxsts";
        public const string 上上上周_新开建面均价 = "sssz_xkjmjj";
        public const string 上上上周_新开套内均价 = "sssz_xktnjj";
        public const string 上上上周_套均总价 = "ssszrg_tjzj";
        public const string 上上上周_认购套数 = "ssszrgts";
        public const string 上上上周_认购套内均价 = "sssz_rgtnjj";
        public const string 上上上周_认购建面均价 = "sssz_rgjmjj";
        public const string 上上上周_认购套内体量 = "sssz_rgtntl";
        public const string 上上上周_认购建面体量 = "sssz_rgjmtl";
        public const string 上上上周_认购金额 = "sssz_rgje";
        public const string 上上上周_主力建面区间 = "sssz_zljmqj";
        public const string 上上上周_主力套内面积区间 = "sssz_zltnqj";



        public static string[] _认购数据 = {"xkts", "xkxsts", "bz_zljmqj", "bz_zltnqj", "xktnjj", "xkjmjj",
                                            "bz_xkts","bz_xkxsts","bz_xkjmjj","bz_xktnjj","bzrg_tjzj","bz_rgts","bz_rgtnjj","bz_rgjmjj","bz_rgtntl","bz_rgjmtl","bz_rgje",
                                            "sz_xkts","sz_xkxsts","sz_xkjmjj","sz_xktnjj","sszrg_tjzj","sz_rgts","sz_rgtnjj","sz_rgjmjj","sz_rgtntl","sz_rgjmtl","sz_rgje","sz_bzdfl","sz_zljmqj", "sz_zltnqj","sz_bzld",
                                            "ssz_xkts","ssz_xkxsts","ssz_xkjmjj","ssz_xktnjj","sszrg_tjzj","ssz_rgts","ssz_rgtnjj","ssz_rgjmjj","ssz_rgtntl","ssz_rgjmtl","ssz_rgje","ssz_zljmqj", "ssz_zltnqj",
                                            "sssz_xkts","sssz_xkxsts","sssz_xkjmjj","sssz_xktnjj","ssszrg_tjzj","sssz_rgts","sssz_rgtnjj","sssz_rgjmjj","sssz_rgtntl","sssz_rgjmtl","sssz_rgje","sssz_zljmqj", "sssz_zltnqj",
                                            "bz_cjtnhb", "bz_tnjjhb", "bz_bhyy", "bz_bzkc", "bz_bzld", "bz_bzdfl", "bz_yh", "bz_yxdz", "bz_hd", "bz_xzjtyj"
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

        public const string 本周_备案套数环比 = "bz_cjtshb";
        public const string 本周_成交金额环比 = "bz_cjjehb";
        public const string 本周_建筑面积环比 = "bz_jzmjhb";
        public const string 本周_套内面积环比 = "bz_tnmjhb";
        public const string 本周_建面均价环比 = "bz_jmjjhb";
        public const string 本周_套内均价环比 = "bz_tnjjhb";
        public const string 本周_套均总价环比 = "bz_tjzjhb";

        public const string 上周_备案套数 = "sz_ts";
        public const string 上周_成交金额 = "sz_cjje";
        public const string 上周_建筑面积 = "sz_jzmj";
        public const string 上周_套内面积 = "sz_tnmj";
        public const string 上周_建面均价 = "sz_jmjj";
        public const string 上周_套内均价 = "sz_tnjj";
        public const string 上周_套均总价 = "sz_tjzj";

        public const string 上上周_备案套数 = "ssz_ts";
        public const string 上上周_成交金额 = "ssz_cjje";
        public const string 上上周_建筑面积 = "ssz_jzmj";
        public const string 上上周_套内面积 = "ssz_tnmj";
        public const string 上上周_建面均价 = "ssz_jmjj";
        public const string 上上周_套内均价 = "ssz_tnjj";
        public const string 上上周_套均总价 = "ssz_tjzj";

        public const string 上上上周_备案套数 = "sssz_ts";
        public const string 上上上周_成交金额 = "sssz_cjje";
        public const string 上上上周_建筑面积 = "sssz_jzmj";
        public const string 上上上周_套内面积 = "sssz_tnmj";
        public const string 上上上周_建面均价 = "sssz_jmjj";
        public const string 上上上周_套内均价 = "sssz_tnjj";
        public const string 上上上周_套均总价 = "sssz_tjzj";


        public static string[] _备案数据 = { "bz_ts","bz_cjje","bz_jzmj","bz_tnmj","bz_jmjj","bz_tnjj", "bz_tjzj",
                                             "sz_ts","sz_cjje","sz_jzmj","sz_tnmj","sz_jmjj","sz_tnjj","sz_tjzj",
                                             "ssz_ts","ssz_cjje","ssz_jzmj","ssz_tnmj","ssz_jmjj","ssz_tnjj","ssz_tjzj",
                                             "sssz_ts","sssz_cjje","sssz_jzmj","sssz_tnmj","sssz_jmjj","sssz_tnjj","sssz_tjzj",
                                             "bz_cjtshb","bz_cjjehb","bz_jzmjhb","bz_tnmjhb","bz_jmjjhb","bz_tnjjhb","bz_tjzjhb"
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
    /// <summary>
    /// 统计项目
    /// </summary>
    public class Base_Config_TJXM
    {
        public const string 区域 = "qy";
        public const string 组团 = "zt";
        public const string 开发商 = "kfs";
        public const string 楼盘名称 = "lpmc";
        public const string 备案套数 = "ts";
        public const string 建筑面积 = "jzmj";
        public const string 套内面积 = "tnmj";
        public const string 成交金额 = "cjje";
        public const string 建面均价 = "jmjj";
        public const string 套内均价 = "tnjj";
        public const string 套均总价 = "tjzj";
        public static string[] _统计项目参数名称 = { "qy","zt", "kfs","lpmc","ts", "jzmj", "tnmj", "cjje", "jmjj","tnjj","tjzj" };
    }

}

