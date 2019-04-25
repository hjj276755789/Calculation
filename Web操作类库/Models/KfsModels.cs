using Calculation.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models.Models
{
    public class KfsModels : Data_Item<KfsModels>
    {
        public string kfsbh { get; set; }
        public string kfsmc { get; set; }
        public string kfslx { get; set; }
        public string kfslxr { get; set; }
        public string kfslxrdh { get; set; }
        public string kfscjsj { get; set; }
        public string fkszt { get; set; }
        public string bz { get; set; }

    }
    public class KFSMBModels :Data_Item<KFSMBModels>
    {
        public string kfsbh { get; set; }
        public string mbbh { get; set; }
        public string mbmc { get; set; }
        public string rwcs { get; set; }
    }
    public class YHFZKFSModels : Data_Item<YHFZKFSModels>
    {
        public string kfsbh { get; set; }

        public string kfsmc { get; set; }
        public string kfsxr { get; set; }
        public string kfslxrdh { get; set; }

        public string sffp { get; set; }
    }
}
