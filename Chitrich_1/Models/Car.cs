using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Chitrich_1.Models
{
    internal class Car
    {
        public string Brand { get; set; }
        public string Model { get; set; }
        public int Year { get; set; }
        public int Price { get; set; }
        public string Color { get; set; }

        public Car() { }
        public Car(List<string> fields)
        {
            for (int i = 0; i < fields.Count; i++)
            {
                switch (i)
                {
                    case 0:
                        Brand = fields[i];
                        break;
                    case 1:
                        Model = fields[i];
                        break;
                    case 2:
                        Year = int.Parse(fields[i]);
                        break;
                    case 3:
                        Price = int.Parse(fields[i]);
                        break;
                    case 4:
                        Color = fields[i];
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
