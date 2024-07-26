using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Chitrich_1.Models;
using System.Reflection;

namespace Chitrich_1.Services
{
    class ExelServices
    {
        public static List<T> Read<T>(string fileName) where T : class
        {
            List<People> peoples = new List<People>();
            if (SpreadsheetDocument.Open(fileName, false) == null)
            {
                throw new Exception("spreadsheetDocument is null");
            }
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
                {
                    if (spreadsheetDocument.WorkbookPart == null)
                    {
                        throw new Exception();
                    }
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault()!;
                    if (sheet.Id == null)
                    {
                        throw new Exception("Sheet not found");
                    }
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    ConstructorInfo constructor = typeof(T).GetConstructor(new[] { typeof(List<string>) })!;

                    int rowNum = 0;
                    List<T> ret = new List<T>();
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        T obj;
                        if (rowNum > 0 && constructor != null)
                        {
                            List<string> fields = new List<string>();
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                fields.Add(GetCellValue(spreadsheetDocument, cell));
                            }
                            obj = (T)constructor.Invoke(new object[] { fields });
                            ret.Add(obj);
                        }
                        rowNum++;
                    }
                    return ret;
                }
            
        }

        public static void Save(List<People> peoples)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(@"C:\Users\taras\source\repos\Chitrich_1\Chitrich_1\Files\output.xlsx", SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = spreadsheetDocument.WorkbookPart!.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                if (sheetData == null)
                {
                    throw new Exception();
                }

                int choise = 1;
                List<int> choises = new List<int>();

                while (choise != 0)
                {
                    Console.Write("Select the columns to save\n"
                        + "1 - Id\n"
                        + "2 - Name\n"
                        + "3 - Age\n"
                        + "4 - Salary\n"
                        + "5 - Department\n"
                        + "Else - exit\n"
                        + "Your choise: ");
                    choise = int.Parse(Console.ReadLine() ?? "0");
                    Console.WriteLine();
                    if (choise >= 1 && choise <= 5)
                    {
                        choises.Add(choise);
                    }
                    else
                    {
                        choise = 0;
                    }
                }
                #region header

                Row headerRow = new Row();
                int i = 1;
                foreach (int chois in choises)
                {
                    switch (chois)
                    {
                        case 1:
                            headerRow.Append(CreateTextCell("Id"));
                            break;
                        case 2:
                            headerRow.Append(CreateTextCell("Name"));
                            break;
                        case 3:
                            headerRow.Append(CreateTextCell("Age"));
                            break;
                        case 4:
                            headerRow.Append(CreateTextCell("Salary"));
                            break;
                        case 5:
                            headerRow.Append(CreateTextCell("Department"));
                            break;
                    }
                    i++;
                }
                #endregion
                #region calls
                sheetData.AppendChild(headerRow);
                foreach (People people in peoples)
                {
                    Row dataRow = new Row();
                    foreach (int chois in choises)
                    {
                        switch (chois)
                        {
                            case 1:
                                dataRow.Append(CreateIntCell(people.Id));
                                break;
                            case 2:
                                if (people.Name != null)
                                {
                                    dataRow.Append(CreateTextCell(people.Name));
                                    break;
                                }
                                else
                                {
                                    throw new Exception();
                                }
                            case 3:
                                dataRow.Append(CreateIntCell(people.Age));
                                break;
                            case 4:
                                dataRow.Append(CreateIntCell(people.Salary));
                                break;
                            case 5:
                                if (people.Department != null)
                                {
                                    dataRow.Append(CreateTextCell(people.Department));
                                    break;
                                }else
                                {
                                    throw new Exception();
                                }
                                
                        }
                    }
                    sheetData.AppendChild(dataRow);
                }
                #endregion

                workbookPart.Workbook.Save();
            }
        }

        private static Cell CreateTextCell(string cellValue)
        {
            Cell cell = new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(cellValue)
            };
            return cell;
        }

        private static Cell CreateIntCell( int cellValue)
        {
            Cell cell = new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(cellValue)
            };
            return cell;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
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
            }else
            {
                throw new Exception();
            }
            
        }

        private static string CellNum(int i)
        {
            switch (i)
            {
                case 1:
                    return "A1";
                case 2:
                    return "B1";
                case 3:
                    return "C1";
                case 4:
                    return "D1";
                case 5:
                    return "E1";
                default:
                    return "Eror";
            }
        }
    }
}
