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
        public int yhbh { get; set; }
        public string yhmc { get; set; }
        public string yhlx { get; set; }

    }
   
    public class JSXX
    {
        public string jsbh { get; set; }
        public string jsmc { get; set; }
        public string jsms { get; set; }

    }
    public class QXXX {
        public string qxbh { get; set; }
        public string qxmc { get; set; }
        public string qxms { get; set; }
        public string qxkzq { get; set; }
        public string qxst { get; set; }
        /// <summary>
        /// 权限类型
        /// </summary>
        public string qxlx { get; set; }
        public string fqxbh { get; set; }
        public string tb { get; set; }
        public List<QXXX> xjqx { get; set; }

    }
    public class YHJS
    {
        public int yhbh { get; set; }
        public string yhm { get; set; }
        public int jsbh { get; set; }
        public string jsmc { get; set; }
    }
    public class JSQX
    {
        public int jsbh { get; set; }
        public string jsmc { get; set; }
        public int qxbh { get; set; }
        public string qxmc { get; set; }
        public string qxkzq { get; set; }
        public string qxst { get; set; }
    }
}
