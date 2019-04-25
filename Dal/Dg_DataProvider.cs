using Calculation.Models;
using Calculation.Models.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Dal
{
    public class Dg_DataProvider : MySqlDbhelper
    {
        public IPageList<DGModels> GET_DG(string tj, string pagesize, string pagenow)
        {
            string sql = "";
            return null;
        }
    }
}
