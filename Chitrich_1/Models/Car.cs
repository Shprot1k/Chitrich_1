using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Chitrich_1.Models
{
    internal class Car : BaseClass
    {
        public string? Brand { get; set; }
        public string? Model { get; set; }
        public int Year { get; set; }
        public int Price { get; set; }
        public string? Color { get; set; }

        public Car() { }
        public override Car OdjFromRow(Row row, SpreadsheetDocument spreadsheetDocument)
        {
            Car car = new Car();
            int i = 0;
            foreach (Cell cell in row.Elements<Cell>())
            {
                switch (i)
                {
                    case 0:
                        Brand = GetCellValue(spreadsheetDocument, cell);
                        break;
                    case 1:
                        Model = GetCellValue(spreadsheetDocument, cell);
                        break;
                    case 2:
                        Year = int.Parse(cell.InnerText);
                        break;
                    case 3:
                        Price = int.Parse(cell.InnerText);
                        break;
                    case 4:
                        Color = GetCellValue(spreadsheetDocument, cell);
                        break;
                    default:
                        break;
                }
                i++;
            }
            return car;
        }
    }
}
