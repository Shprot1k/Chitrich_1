using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Chitrich_1.Models
{
    class People : BaseClass
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public int Age { get; set; }
        public int Salary { get; set; }
        public string? Department { get; set; } 
        public People() { }

        public override People OdjFromRow(Row row, SpreadsheetDocument spreadsheetDocument)
        {
            People people = new People();
            int i = 0;
            foreach (Cell cell in row.Elements<Cell>())
            {
                switch (i)
                {
                    case 0:
                        Id = int.Parse(cell.InnerText);
                        break;
                    case 1:
                        Name = GetCellValue(spreadsheetDocument, cell);
                        break;
                    case 2:
                        Age = int.Parse(cell.InnerText);
                        break;
                    case 3:
                        Salary = int.Parse(cell.InnerText);
                        break;
                    case 4:
                        Department = GetCellValue(spreadsheetDocument, cell);
                        break;
                    default:
                        break;
                }
                i++;
            }
            return people;
        }

    }
}
        