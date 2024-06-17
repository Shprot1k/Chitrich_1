using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Chitrich_1.Models
{
    class People
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public int Age { get; set; }
        public int Salary { get; set; }
        public string Department { get; set; }


         public static List<People> Read()
        {
            Workbook wb = new Workbook("generated_excel_data.xlsx");
            WorksheetCollection collection = wb.Worksheets;
            List<People> peoples = new List<People>();
            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {
                Worksheet worksheet = collection[worksheetIndex];
                //Console.WriteLine("Worksheet: " + worksheet.Name);
                int rows = worksheet.Cells.MaxDataRow;
                int cols = worksheet.Cells.MaxDataColumn;
                for (int row = 1; row < rows; row++)
                {
                    People people = new People();
                    for (int column = 0; column < cols; column++)
                    {
                        if (column == 0)
                        {
                            people.ID = int.Parse(worksheet.Cells[row, column].Value.ToString());
                            
                        }else if (column == 1)
                        {
                            people.Name = worksheet.Cells[row, column].Value.ToString();
                        }else if (column == 2)
                        {
                            people.Age = int.Parse(worksheet.Cells[row, column].Value.ToString());
                        }
                        else if (column == 3)
                        {
                            people.Salary = int.Parse(worksheet.Cells[row, column].Value.ToString());
                        }else if (row == 4)
                        {
                            people.Department = worksheet.Cells[row, column].Value.ToString();
                        }
                    }
                    peoples.Add(people);
                }
            }
            return peoples;
        }

        public static void Write(List<People> peoples)
        {
            
            foreach (People people in peoples)
            {
                Console.WriteLine(people.ID + " " 
                    + people.Name + " " 
                    + people.Age + " " 
                    + people.Salary + " "
                    + people.Department);
            }
        }
    }
}
