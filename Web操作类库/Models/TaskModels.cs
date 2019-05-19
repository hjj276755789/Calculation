using Calculation.Base;
using Calculation.Models.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models.Models
{
    /// <summary>
    /// 周报任务
    /// </summary>
    public class ZB_TaskModels :Data_Item<ZB_TaskModels>
    {
        public int kfsbh { get; set; }
        /// <summary>
        /// 开发商名称
        /// </summary>
        public string kfsmc { get; set; }
        /// <summary>
        /// 未启动
        /// </summary>
        public string wqd { get; set; }
        /// <summary>
        /// 生成中
        /// </summary>
        public string scz { get; set; }
        /// <summary>
        /// 已完成
        /// </summary>
        public string ywc { get; set; }
    }
}
