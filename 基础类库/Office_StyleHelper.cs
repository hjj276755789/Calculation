using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public class Office_StyleHelper
    {
        public static void setdefaultstyle(IAutoShape shape)
        {
            ITextFrame tf1 = shape.TextFrame;
            foreach (var item in tf1.Paragraphs)
            {
                IPortion port = item.Portions[0];
                port.PortionFormat.LatinFont = new FontData("微软雅黑");
                port.PortionFormat.FontBold = NullableBool.NotDefined;
                port.PortionFormat.FontHeight = 12;
            }
        }
    }
}
