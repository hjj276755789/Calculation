using Calculation.Base;
using Calculation.Models.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models
{

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



    public class Data_Cjba_Default : Data_Item<Data_Cjba_Default>
    {
        public int id {get; set; }
        public string cjrq { get; set; }
        public string qy { get; set; }
        public string zt { get; set; }
        public string kfs { get; set; }
        public string lpmc { get; set; }
        public string yt { get; set; }
        public string xfyt { get; set; }
        public string hx { get; set; }
        public  double jzmj { get; set; }
        public  double tnmj { get; set; }
        public long cjje { get; set; } 
        public int ts { get; set; }
        public string zlmjqj { get; set; }
        public int nf { get; set; }
        public int zc { get; set; }
        public string zcmc { get; set; }
    }

    public class Data_JHNF : Data_Item<Data_JHNF>
    {
       public int nf { get; set; }
    }




}