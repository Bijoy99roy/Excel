using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Text;

namespace Madam_Nanomaterial_Extract.Utils
{
    class UtilsCell
    {
        public static string getCellStringValue(Worksheet worksheet, int rowIndex, int colIndex)
        {
            Cell cell = worksheet.Cells[rowIndex, colIndex];
            if (cell != null) return cell.StringValue;
            else return "";
        }
    }
}
