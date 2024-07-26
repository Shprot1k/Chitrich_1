using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Chitrich_1.Models
{
    class People
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public int Age { get; set; }
        public int Salary { get; set; }
        public string? Department { get; set; } 
        public People() { }
        
        public People(List<string> fields)
        {
            for (int i = 0; i < fields.Count; i++)
            {
                switch (i)
                {
                    case 0:
                        Id = int.Parse(fields[i]);
                        break;
                    case 1:
                        Name = fields[i];
                        break;
                    case 2:
                        Age = int.Parse(fields[i]);
                        break;
                    case 3:
                        Salary = int.Parse(fields[i]);
                        break;
                    case 4:
                        Department = fields[i];
                        break;
                    default:
                        break;
                }
            }

        }
        
    }
}
        