using Calculation.Base;
using Calculation.Models.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models.Models
{
    public class DataModels
    {

    }

    public class Data_JHSQXQ
    {
        public int jhbh { get; set; }
        public int nf { get; set; }
        public int zc { get; set; }
        public string zcmc
        {
            get
            {
                if (this.nf != 0 && this.zc != 0)
                    return Base_date.GET_ZCMC(this.nf, this.zc);
                else return "空缺";
            }
        }
        public Int64 cjjl { get; set; }
        public Int64 xzys { get; set; }
        public Int64 tdcj { get; set; }
        public Int64 rgsj { get; set; }
    }
}