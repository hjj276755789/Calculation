using Aspose.Slides;
using Aspose.Slides.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public class Office_ChartStyle
    {
        public Base_Config.坐标方向 坐标方向 { get; set; }
        public LegendDataLabelPosition 文字位置 { get; set; }
        public bool 是否显示文字 { get; set; }

        public   TextVerticalType 文字旋转方向 { get; set; }

    }
}
