using Calculation.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Dal
{
    public class ZB_Param_JP_DataProvider
    {
        public static List<JP_JZGJ> GET_JPGJ()
        {
            string sql = "select * from calculation. dmb_jzgj order by px";
            return Models.Modelhelper.类列表赋值<JP_JZGJ>(new JP_JZGJ(), MySqlDbhelper.GetDataSet(sql).Tables[0]);
        }
    }
}
