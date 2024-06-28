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

        public static void Print(List<People> peoples)
        {  
            foreach (People people in peoples)
            {
                Console.WriteLine(people.ID + "\t" 
                    + people.Name + "\t" 
                    + people.Age + "\t" 
                    + people.Salary + "\t"
                    + people.Department);
            }
        }

        public static List<People> SortByAge(List<People> peoples)
        {
            for (int i = 0; i < peoples.Count; i++)
            {
                for (int j = 0; j < peoples.Count; j++)
                {
                    if (peoples[i].Age < peoples[j].Age)
                    {
                        var temp = peoples[i];
                        peoples[i] = peoples[j];
                        peoples[j] = temp;
                    }
                }
            }
            return peoples;
        }

        public static void Save(List<People> peoples) // мало б працювати, але нє
        {
            
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            worksheet.Cells[0, 0].PutValue("ID");
            worksheet.Cells[0, 1].PutValue("Name");
            worksheet.Cells[0, 2].PutValue("Age");
            worksheet.Cells[0, 3].PutValue("Salary");
            worksheet.Cells[0, 4].PutValue("Department");

            for (int i = 0; i < peoples.Count; i++)
            {
                worksheet.Cells[i + 1, 0].PutValue(peoples[i].ID);
                worksheet.Cells[i + 1, 1].PutValue(peoples[i].Name);
                worksheet.Cells[i + 1, 2].PutValue(peoples[i].Age);
                worksheet.Cells[i + 1, 3].PutValue(peoples[i].Salary);
                worksheet.Cells[i + 1, 4].PutValue(peoples[i].Department);
            }
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

        }
        public static void TestSave() //тестовий метод для перевірки зберігання
        {
            Workbook wb = new Workbook("generated_excel_data.xlsx");
            wb.Save("output1.xlsx");
        }
    }
}
