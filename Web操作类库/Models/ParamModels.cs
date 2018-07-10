using Calculation.Models.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models.Models
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
}
