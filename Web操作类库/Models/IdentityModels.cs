using Calculation.Models.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models
{
    public class YHXX
    {
        public int id { get; set; }
        public string yhm { get; set; }
        public YH_LX yhlx { get; set; }
    }
   
    public class JSXX
    {
        public int id { get; set; }
        public string jsmc { get; set; }
        public string jsms { get; set; }
        
    }
    public class QXXX {
        public int id { get; set; }
        public string qxmc { get; set; }
        public string qxms { get; set; }
        public string qxkzq { get; set; }
        public string qxst { get; set; }

    }
    public class YHJS
    {
        public int yhid { get; set; }
        public string yhm { get; set; }
        public int jsid { get; set; }
        public string jsmc { get; set; }
    }
    public class JSQX
    {
        public int jsid { get; set; }
        public string jsmc { get; set; }
        public int qxid { get; set; }
        public string qxmc { get; set; }
        public string qxkzq { get; set; }
        public string qxst { get; set; }
    }
}
