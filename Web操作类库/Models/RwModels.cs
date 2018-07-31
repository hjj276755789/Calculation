using Calculation.Models.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models
{
    public class Rw_Item_Model
    {
        public int rwid { get; set; }
        public string rwmc { get; set; }
        public int nf { get; set; }
        public int zc { get; set; }
        /// <summary>
        ///  下载地址
        /// </summary>
        public string xzdz { get; set; }
        /// <summary>
        /// 定稿下载地址
        /// </summary>
        public string xzdz2 { get; set; }
    }

    public class Rw_List
    {
        public int rwid { get; set; }
        public string rwmc { get; set; }
        public int mbid { get; set; }
        public int nf { get; set; }
        public int  zc { get; set; }
        public RW_ZT zt { get; set; }
        /// <summary>
        /// 定稿下载地址
        /// </summary>
        public string xzdz2 { get; set; }
    }


    public class Rw_Cofirm_data
    {
        public int rwid { get; set; }
        public string rwmc { get; set; }
        public int nf { get; set; }
        public int zc { get; set; }
        public DATA_ZT cjjl_zt { get; set; }
        public DATA_ZT xzys_zt { get; set; }
        public DATA_ZT tdcj_zt { get; set; }
        public DATA_ZT rgsj_zt { get; set; }
    }
}
