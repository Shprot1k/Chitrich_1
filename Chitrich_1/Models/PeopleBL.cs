using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Chitrich_1.Models
{
    internal class PeopleBL
    {
        public static List<People> Read(string filePath, bool firstRowIsHeader = true)
        {
            List<People> peoples = new List<People>();
            //List<string> Headers = new List<string>();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
            {
                //Read the first Sheets 
                Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                
                
                int counter = 0;
                foreach (Row row in rows)
                {
                    //Read the first row as header
                    if (counter == 0 && firstRowIsHeader ==true)
                    {
                        /* поки непотрібно
                        var j = 1;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            var colunmName = firstRowIsHeader ? GetCellValue(doc, cell) : "Field" + j++;
                            Console.WriteLine(colunmName);
                            Headers.Add(colunmName);
                            dt.Columns.Add(colunmName);                            
                        }
                        */
                    }
                    else
                    {
                        People people = new People();
                        int i = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            switch (i)
                            {
                                case 0:
                                    //people.Id = int.Parse(GetCellValue(doc, cell));
                                    people.Id = int.Parse(cell?.CellValue?.Text);
                                    i++;
                                    break;
                                case 1:
                                    people.Name = GetCellValue(doc, cell);
                                    i++;
                                    break;
                                case 2:
                                    people.Age = int.Parse(GetCellValue(doc, cell));
                                    i++;
                                    break;
                                case 3:
                                    people.Salary = int.Parse(GetCellValue(doc, cell));
                                    i++;
                                    break;
                                case 4:
                                    people.Department = GetCellValue(doc, cell);
                                    i++;
                                    break;

                            }

                        }
                        peoples.Add(people);
                        /* ісходнік, якщо є в коміті то забув видалити
                        dt.Rows.Add();
                        int i = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = GetCellValue(doc, cell);
                            i++;
                        }
                        */
                    }
                    counter++;
                }

            }
            return peoples;
        }

        public static void Print(List<People> peoples)
        {
            foreach (People people in peoples)
            {
                Console.WriteLine(people.Id + "\t"
                    + people.Name + "\t"
                    + people.Age + "\t"
                    + people.Salary + "\t"
                    + people.Department);
            }
        }

        private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                //return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
                return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.ToString();
            }
            return value;
        }

        /*
        Workbook wb = new Workbook(filePath);
        WorksheetCollection collection = wb.Worksheets;
        List<People> peoples = new List<People>();
        for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
        {
            Worksheet worksheet = collection[worksheetIndex];
            int rows = worksheet.Cells.MaxDataRow;
            int cols = worksheet.Cells.MaxDataColumn;
            for (int row = 1; row < rows; row++)
            {
                People people = new People();
                for (int column = 0; column < cols; column++)
                {
                    if (column == 0)
                    {
                        people.Id = int.Parse(worksheet.Cells[row, column].Value.ToString());

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
        */
    }
        /*

        
        
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

        public static void Save(List<People> peoples)
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
                worksheet.Cells[i + 1, 0].PutValue(peoples[i].Id);
                worksheet.Cells[i + 1, 1].PutValue(peoples[i].Name);
                worksheet.Cells[i + 1, 2].PutValue(peoples[i].Age);
                worksheet.Cells[i + 1, 3].PutValue(peoples[i].Salary);
                worksheet.Cells[i + 1, 4].PutValue(peoples[i].Department);
            }
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }    
    }
    */
    
}