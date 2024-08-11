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
    }

    interface IBase<T>
    {
         T OdjFromRow(Row row);
    }
}
