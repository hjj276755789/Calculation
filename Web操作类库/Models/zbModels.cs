using Calculation.Models.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models
{
    public class Zb_Item_Model
    {
        public int mbid { get; set; }
        public string mbmc { get; set; }
        public MB_Enums mblx { get; set; }
        public MB_XFLX xflx { get; set; }
    }
}
