using Calculation.Models.Enums;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models
{
    public class ParamModels
    {
        /// <summary>
        ///参数编号
        /// </summary>
        public int csid { get; set; }
        /// <summary>
        ///  插件编号
        /// </summary>
        public int cjid { get; set; }
        /// <summary>
        /// 插件名称
        /// </summary>
        public string cjmc{ get; set; }
        /// <summary>
        /// 插件描述
        /// </summary>
        public string csms { get; set; }
        /// <summary>
        /// 插件类型
        /// </summary>
        public CS_LX cslx { get; set; }
        /// <summary>
        /// 是否并列
        /// </summary>
        public int sfbl { get; set; }
    }

    public class ParamValueModel
    {
        /// <summary>
        /// 任务参数ID
        /// </summary>
        public int rwcsid { get; set; }
        /// <summary>
        /// 任务ID
        /// </summary>
        public int rwid { get; set; }
        /// <summary>
        /// 插件ID
        /// </summary>
        public int cjid { get; set; }
        /// <summary>
        /// 参数描述
        /// </summary>
        public string csms { get; set; }
        /// <summary>
        /// 参数内容
        /// </summary>
        public string csnr { get; set; }

    }

    public class JP_ParamValueModel
    {
        public string[] zt { get; set; }
        public string [] qy { get; set; }
        public string[] lpmc { get; set; }
        public string[] yt { get; set; }
        public string[] xfyt { get; set; }
        public string[] hx { get; set; }
        /// <summary>
        /// 主力面积区间，不参与筛选
        /// </summary>
        public string zlmjqj { get; set; }
    }
    /// <summary>
    /// 竞品-本案
    /// </summary>
    public class JP_BA
    {
        public int rwid { get; set; }
        public int id { get; set; }
        public string bamc { get; set; }
        public string ztcs { get; set; }
        public string qycs { get; set; }
        public string lpcs { get; set; }
        public string ytcs { get; set; }
        public string xfytcs { get; set; }
        public string hxcs { get; set; }
        public string zlmjqj { get; set; }

    }
    /// <summary>
    /// 竞品-竞品项目
    /// </summary>
    public class JP_JPXM
    {
        public int id { get; set; }
        public int baid { get; set; }
        public int jzgjid { get; set; }
        public string jzgjmc { get; set; }
        public string ztcs { get; set; }
        public string qycs { get; set; }
        public string lpcs { get; set; }
        public string ytcs { get; set; }
        public string xfytcs { get; set; }
        public string hxcs { get; set; }
        public string zlmjqj { get; set; }
    }

    /// <summary>
    /// 竞品-竞争格局
    /// </summary>
    public class JP_JZGJ
    {
        public int id { get; set; }
        public string jzgjmc { get; set; }
        public int px { get; set; }
    }

    public class JP_BA_INFO 
    {
        public int cjid { get; set; }
        public int id { get; set; }
        public int rwid { get; set; }
        public string bamc { get; set; }
        public string [] qycs { get; set; }
        public string [] ztcs { get; set; }
        public string [] lpcs { get; set; }
        public string [] ytcs { get; set; }
        public string [] xfytcs { get; set; }
        public string [] hxcs { get; set; }
        public string zlmjqj { get; set; }
        public List<JP_JPXM_INFO> jpxmlb { get; set; }
    }
    public class JP_JPXM_INFO
    {
        public int id { get; set; }
        public int baid { get; set; }
        public int jzgjid { get; set; }
        public string jzgjmc { get; set; }
        public int px { get; set; }
        public string[] qycs { get; set; }
        public string[] ztcs { get; set; }
        public string[] lpcs { get; set; }
        public string[] ytcs { get; set; }
        public string[] xfytcs { get; set; }
        public string[] hxcs { get; set; }
        public string zlmjqj { get; set; }

    }
}
