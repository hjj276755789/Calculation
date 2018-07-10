using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public class SlideFactory
    {

        public SlideFactory()
        {
            //破解aspose.slide；
            
        }

        private static SlideInstance obj;

        public static SlideInstance GetInstance()
        {
            if (obj == null) {
                //aspose.slide          
                //AsposeSlideCrack.Crack();
                //aspose.cell
                Office_AsposeSlideCrack.SlideCrack();
                obj = new SlideInstance();
            }
            return obj;
        }
    }
}
