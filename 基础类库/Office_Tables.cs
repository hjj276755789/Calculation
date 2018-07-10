using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public class Office_Tables
    {
        public static void SetChart(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable chart = (ITable)sld.Shapes[index];

            for (int i = 0; i < chart.Columns.Count; i++)
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int j = 0; j < (xsts.HasValue ? xsts : dt.Rows.Count); j++)
                    {
                        chart.Columns[i][j + 1].TextFrame.Text = dt.Rows[j][i].ToString();
                    }
                }
            }
        }
    }
}
