using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Chitrich_1.Models
{
    abstract class BaseClass
    {
        public abstract BaseClass OdjFromRow(Row row, SpreadsheetDocument spreadsheetDocument);

        protected static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (document.WorkbookPart != null && document.WorkbookPart.SharedStringTablePart != null && cell.CellValue != null)
            {
                string value = cell.CellValue.InnerText;

                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    SharedStringTablePart stringTable = document.WorkbookPart.SharedStringTablePart;
                    return stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
                else if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
                {
                    return value == "0" ? "FALSE" : "TRUE";
                }
                else
                {
                    return value;
                }
            }
            else
            {
                throw new Exception();
            }

        }
    }

    
}
