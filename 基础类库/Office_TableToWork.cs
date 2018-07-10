using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public class Office_TableToWork
    {
        public static Workbook GetWorkBooxFromDataTable(DataTable dt)
        {
            Workbook workbook = new Workbook();

            Worksheet cellSheet = workbook.Worksheets[0];
            int rowIndex = 0;
            int colIndex = 0;
            int colCount = dt.Columns.Count;
            int rowCount = dt.Rows.Count;
            //Head 列名处理
            for (int i = 0; i < colCount; i++)
            {
                cellSheet.Cells[rowIndex, colIndex].PutValue(dt.Columns[i].ColumnName);
                //cellSheet.Cells[rowIndex, colIndex].SetStyle(headStyle);
                colIndex++;
            }
            rowIndex++;
            //Cell 其它单元格处理
            for (int i = 0; i < rowCount; i++)
            {
                colIndex = 0;
                for (int j = 0; j < colCount; j++)
                {
                    cellSheet.Cells[rowIndex, colIndex].PutValue(dt.Rows[i][j]);
                    //cellSheet.Cells[rowIndex, colIndex].SetStyle(cellStyle);
                    colIndex++;
                }
                rowIndex++;
            }
            return workbook;

        }
    }
}
