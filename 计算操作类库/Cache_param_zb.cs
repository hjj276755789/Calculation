using Calculation.Models.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    public class Cache_param_zb
    {
        private static Cache_param_zb uniqueInstance;
        public static Cache_param_zb ini_zb(int mbid,int nf,int zc)
        {
            if (uniqueInstance == null)
            {
                uniqueInstance = new Cache_param_zb();
                value = Dal.Param_DataProvider.GET_MBCJCSNR(mbid,nf,zc);
            }
            return uniqueInstance;
        }

        public static List<ParamValueModel> value { get; set; }
    }
}
